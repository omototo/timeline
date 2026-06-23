/**
 * `TimelineEngineImpl` — the concrete Timeline Engine (Wave 1: value path).
 *
 * Functional core, imperative shell (ADR-0013): synchronous and stateful, it
 * holds the {@link ShadowState}, the HEAD, and the current branch's in-memory
 * delta log. Mutators RETURN an {@link EffectEnvelope} describing the I/O the
 * shell must perform; the engine never calls the store or does I/O itself.
 *
 * Wave 1 implements `ingest` for the `value` Observation kind in Present mode:
 * diff the after-slab against the Shadow State, and — if anything changed —
 * append a Step to the branch log, advance `HEAD.stepIndex`, update the Shadow
 * State, and emit `{ persist: [appendDelta, setHead] }`. There is no reconcile
 * in Present: the user already typed the value (ADR-0008 frozen-value capture
 * happens at the Delta, not via a write-back). A no-op observation records no
 * Step and returns an empty envelope.
 *
 * Pure: no Office.js, DOM, or React.
 */
import { ShadowState } from './shadow-state.ts';
import type { ShadowSnapshot } from './shadow-state.ts';
import { reconstruct } from './reconstruct.ts';
import { projectionDiff } from './project.ts';
import type { TimelineEngine } from './engine.ts';
import type {
  BranchId,
  Delta,
  EffectEnvelope,
  Head,
  Observation,
  PersistedHead,
  PersistOp,
  ReconcileOp,
  ReconcilePlan,
  StepDetail,
  StepRef,
  StructuralDelta,
  StructuralObservation,
  TimelineQuery,
  TimelineView,
  ValueDelta,
  WorkbookSnapshot,
  WorksheetDelta,
  WorksheetObservation,
} from './types.ts';

/** The default branch id every workbook starts on. */
const MAIN_BRANCH: BranchId = 'main';

/** The sheet id of the single Preview Sheet the engine projects onto. */
const PREVIEW_SHEET: BranchId = '__preview__';

/** Default adaptive-keyframe cadence (Q6): step count + cumulative delta bytes. */
const DEFAULT_KEYFRAME_STEP_INTERVAL = 100;
const DEFAULT_KEYFRAME_BYTE_THRESHOLD = 64 * 1024;

/**
 * Engine construction options. The adaptive keyframe cadence is configurable
 * (Q6): a keyframe is written when EITHER `keyframeStepInterval` steps OR
 * `keyframeByteThreshold` cumulative delta bytes have accrued since the last
 * keyframe — whichever fires first.
 */
export interface TimelineEngineOptions {
  /** Steps since last keyframe that triggers a new keyframe (default 100). */
  keyframeStepInterval?: number;
  /** Cumulative delta bytes since last keyframe that triggers one (default 64 KiB). */
  keyframeByteThreshold?: number;
}

/** A resident keyframe: a serialized Shadow State snapshot at a step boundary. */
interface Keyframe {
  readonly stepIndex: number;
  readonly snapshot: ShadowSnapshot;
}

/**
 * Estimate the serialized byte size of a {@link Delta} for the adaptive
 * keyframe cadence. Uses the JSON encoding length as a cheap, deterministic
 * proxy for on-disk delta size (the store serializes deltas as JSON; ADR-0007).
 */
function deltaBytes(delta: Delta): number {
  return JSON.stringify(delta).length;
}

/**
 * A recorded Step on a branch: the Delta plus its 0-based ordinal in the log.
 * `HEAD.stepIndex` is not part of the frozen `Head` shape (the spec's `Head`
 * tracks branch + mode + previewStepIndex); the engine tracks the Present tip's
 * next-write index internally and exposes it via {@link TimelineEngineImpl.tipStepIndex}.
 */
interface Step {
  readonly stepIndex: number;
  readonly delta: Delta;
}

/**
 * Diagnostic raised when an Observation arrives while HEAD is in Preview mode.
 * The shell must lock the real sheet during Preview; if one slips through, the
 * engine refuses it (no-op envelope) and records why.
 */
export interface IngestDiagnostic {
  code: 'ingestInPreview' | 'unsupportedKind';
  message: string;
}

export class TimelineEngineImpl implements TimelineEngine {
  readonly #shadow = new ShadowState();
  /** Per-branch append-only Step log (resident; deltas are small — ADR-0007). */
  readonly #log = new Map<BranchId, Step[]>();
  #head: Head = { branchId: MAIN_BRANCH, mode: 'present' };
  /** Last diagnostic from a refused/no-op `ingest`, for the shell to surface. */
  #lastDiagnostic: IngestDiagnostic | null = null;

  // --- Adaptive keyframe cadence (Q6) -------------------------------------
  readonly #keyframeStepInterval: number;
  readonly #keyframeByteThreshold: number;
  /** Per-branch resident keyframes (single branch for now), step-ascending. */
  readonly #keyframes = new Map<BranchId, Keyframe[]>();
  /** Steps appended on the current branch since its last keyframe. */
  #stepsSinceKeyframe = 0;
  /** Cumulative delta bytes appended since the last keyframe. */
  #bytesSinceKeyframe = 0;

  // --- Navigation: currently-projected Preview state ----------------------
  /**
   * The Shadow State the engine last projected onto the Preview Sheet, or
   * `null` when no Preview is active. `goto` diffs the target against this to
   * produce a MINIMAL value-mode plan; `returnToPresent` clears it.
   */
  #projected: ShadowState | null = null;

  constructor(options: TimelineEngineOptions = {}) {
    this.#keyframeStepInterval = options.keyframeStepInterval ?? DEFAULT_KEYFRAME_STEP_INTERVAL;
    this.#keyframeByteThreshold = options.keyframeByteThreshold ?? DEFAULT_KEYFRAME_BYTE_THRESHOLD;
  }

  // -------------------------------------------------------------------------
  // Recording — Present mode only (Wave 1: value path)
  // -------------------------------------------------------------------------

  ingest(obs: Observation): EffectEnvelope {
    this.#lastDiagnostic = null;

    // Present-only: an Observation in Preview is refused (shell must lock the
    // real sheet during Preview). No Step, empty envelope.
    if (this.#head.mode === 'preview') {
      this.#lastDiagnostic = {
        code: 'ingestInPreview',
        message:
          'Observation arrived while HEAD is in Preview; the shell must lock the real sheet.',
      };
      return {};
    }

    switch (obs.kind) {
      case 'value':
        return this.#ingestValue(obs);
      case 'structural':
        return this.#ingestStructural(obs);
      case 'worksheet':
        return this.#ingestWorksheet(obs);
    }
  }

  /**
   * Value path (Wave 1): diff the after-slab against the Shadow State and, if
   * anything changed, record a {@link ValueDelta} Step. A no-op observation
   * records no Step and returns an empty envelope.
   */
  #ingestValue(obs: Extract<Observation, { kind: 'value' }>): EffectEnvelope {
    const changed = this.#shadow.diff(obs);
    if (changed.length === 0) {
      // No-op observation: nothing changed -> no Step, empty envelope.
      return {};
    }

    const delta: ValueDelta = { kind: 'value', sheetId: obs.sheetId, cells: changed };
    return this.#recordStep(delta, () => {
      this.#shadow.apply(delta);
    });
  }

  /**
   * Structural path (Wave 2): a row/column/cell insert or delete is a
   * COORDINATE REMAP, not a value change. Build a {@link StructuralDelta} from
   * the Observation, apply the remap to the Shadow State, and record one Step.
   * Emits NO value diff (ADR-0001 suppress rule) and never rewrites formula
   * text (ADR-0003). Always records a Step (a structural op is a real event
   * even if it moves no currently-populated cells).
   */
  #ingestStructural(obs: StructuralObservation): EffectEnvelope {
    const delta: StructuralDelta = {
      kind: 'structural',
      sheetId: obs.sheetId,
      changeType: obs.changeType,
      address: obs.address,
      ...(obs.shiftDirection !== undefined ? { shiftDirection: obs.shiftDirection } : {}),
    };
    return this.#recordStep(delta, () => {
      this.#shadow.applyStructural(delta);
    });
  }

  /**
   * Worksheet path (Wave 2): add/delete/rename/reorder (ADR-0005). Build a
   * {@link WorksheetDelta}, apply it to the Shadow State's sheet map, and record
   * one Step.
   */
  #ingestWorksheet(obs: WorksheetObservation): EffectEnvelope {
    const delta: WorksheetDelta = {
      kind: 'worksheet',
      op: obs.op,
      sheetId: obs.sheetId,
      ...(obs.newName !== undefined ? { newName: obs.newName } : {}),
      ...(obs.newPosition !== undefined ? { newPosition: obs.newPosition } : {}),
    };
    return this.#recordStep(delta, () => {
      this.#shadow.applyWorksheet(delta);
    });
  }

  /**
   * Append `delta` as a Step on the current branch, run `applyToShadow` to push
   * the Shadow State forward, and return the Present-mode envelope
   * (`appendDelta` + `setHead`, no reconcile — the user already performed the
   * change in the live workbook).
   */
  #recordStep(delta: Delta, applyToShadow: () => void): EffectEnvelope {
    const branchId = this.#head.branchId;
    const stepIndex = this.#nextStepIndex(branchId);
    this.#appendStep(branchId, { stepIndex, delta });
    applyToShadow();

    const persist: PersistOp[] = [
      { op: 'appendDelta', branchId, delta },
      { op: 'setHead', head: this.head() },
    ];

    // Adaptive keyframe cadence (Q6): account this Step's cost, then — if EITHER
    // the step-count OR the byte threshold is crossed — snapshot the Shadow
    // State (now at `stepIndex`) and emit a writeKeyframe op. The snapshot is
    // also kept resident so reconstruction can replay forward from it.
    this.#stepsSinceKeyframe += 1;
    this.#bytesSinceKeyframe += deltaBytes(delta);
    if (
      this.#stepsSinceKeyframe >= this.#keyframeStepInterval ||
      this.#bytesSinceKeyframe >= this.#keyframeByteThreshold
    ) {
      const snapshot = this.#shadow.snapshot();
      this.#storeKeyframe(branchId, { stepIndex, snapshot });
      this.#stepsSinceKeyframe = 0;
      this.#bytesSinceKeyframe = 0;
      persist.push({ op: 'writeKeyframe', branchId, stepIndex, state: snapshot });
    }

    return { persist };
  }

  /** Append a keyframe to the branch's resident, step-ascending keyframe list. */
  #storeKeyframe(branchId: BranchId, keyframe: Keyframe): void {
    const list = this.#keyframes.get(branchId);
    if (list === undefined) {
      this.#keyframes.set(branchId, [keyframe]);
    } else {
      list.push(keyframe);
    }
  }

  /** The nearest resident keyframe with `stepIndex <= target`, or null. */
  #keyframeAtOrBefore(branchId: BranchId, target: number): Keyframe | null {
    const list = this.#keyframes.get(branchId);
    if (list === undefined) return null;
    let best: Keyframe | null = null;
    for (const kf of list) {
      if (kf.stepIndex <= target && (best === null || kf.stepIndex > best.stepIndex)) {
        best = kf;
      }
    }
    return best;
  }

  /**
   * Reconstruct the Shadow State at `ref` by forward-replay (Q6): seed from the
   * nearest resident keyframe ≤ `ref.stepIndex`, then apply the deltas in the
   * window `(keyframeStepIndex, ref.stepIndex]` forward. Never inverts a delta.
   */
  #reconstructAt(ref: StepRef): ShadowState {
    const keyframe = this.#keyframeAtOrBefore(ref.branchId, ref.stepIndex);
    const from = keyframe === null ? -1 : keyframe.stepIndex;
    const log = this.#log.get(ref.branchId) ?? [];
    const deltas: Delta[] = [];
    for (const step of log) {
      if (step.stepIndex > from && step.stepIndex <= ref.stepIndex) {
        deltas.push(step.delta);
      }
    }
    return reconstruct(keyframe === null ? null : keyframe.snapshot, deltas);
  }

  // -------------------------------------------------------------------------
  // Queries (pure)
  // -------------------------------------------------------------------------

  head(): Head {
    // Return a defensive copy; exactOptionalPropertyTypes-safe (omit preview).
    return this.#head.mode === 'preview' && this.#head.previewStepIndex !== undefined
      ? {
          branchId: this.#head.branchId,
          mode: 'preview',
          previewStepIndex: this.#head.previewStepIndex,
        }
      : { branchId: this.#head.branchId, mode: this.#head.mode };
  }

  timeline(_opts?: TimelineQuery): TimelineView {
    throw new Error('TimelineEngineImpl.timeline is not implemented in Wave 1.');
  }

  inspectStep(_ref: StepRef): StepDetail {
    throw new Error('TimelineEngineImpl.inspectStep is not implemented in Wave 1.');
  }

  // -------------------------------------------------------------------------
  // Lifecycle / navigation — not in Wave 1 scope
  // -------------------------------------------------------------------------

  attach(_observed: WorkbookSnapshot, _persisted: PersistedHead | null): EffectEnvelope {
    throw new Error('TimelineEngineImpl.attach is not implemented in Wave 1.');
  }

  detachToCoauthoring(): EffectEnvelope {
    throw new Error('TimelineEngineImpl.detachToCoauthoring is not implemented in Wave 1.');
  }

  /**
   * Enter (or move within) Preview at `ref` (ADR-0008, Q3). Reconstruct the
   * target state by forward-replay, then emit a {@link ReconcilePlan} targeting
   * the Preview Sheet that is the MINIMAL value-mode diff between the engine's
   * currently-projected state and the target. On first entry the plan is
   * prefixed with `createPreviewSheet` + `activateSheet`; HEAD flips to
   * `preview` and the projected state is tracked so a subsequent `goto`/scrub
   * writes only what changed.
   */
  goto(ref: StepRef): EffectEnvelope {
    const firstEntry = this.#projected === null;
    const target = this.#reconstructAt(ref);
    const from = this.#projected ?? new ShadowState();

    const ops: ReconcileOp[] = [];
    if (firstEntry) {
      ops.push({ op: 'createPreviewSheet', previewSheetId: PREVIEW_SHEET });
      ops.push({ op: 'activateSheet', sheetId: PREVIEW_SHEET });
    }
    ops.push(...projectionDiff(from, target, PREVIEW_SHEET));

    this.#projected = target;
    this.#head = { branchId: ref.branchId, mode: 'preview', previewStepIndex: ref.stepIndex };

    const reconcile: ReconcilePlan = { target: 'previewSheet', ops };
    return {
      reconcile,
      persist: [{ op: 'setHead', head: this.head() }],
    };
  }

  /**
   * Discard the Preview Sheet and reactivate the real sheet (Q3), returning
   * HEAD to Present. Emits a {@link ReconcilePlan} targeting the real sheet
   * that deletes the Preview Sheet then activates the present branch's surface.
   * A no-op (empty envelope) when not currently in Preview.
   */
  returnToPresent(): EffectEnvelope {
    if (this.#projected === null || this.#head.mode !== 'preview') {
      return {};
    }
    const branchId = this.#head.branchId;
    this.#projected = null;
    this.#head = { branchId, mode: 'present' };

    const reconcile: ReconcilePlan = {
      target: 'realSheet',
      ops: [
        { op: 'deletePreviewSheet', previewSheetId: PREVIEW_SHEET },
        { op: 'activateSheet', sheetId: branchId },
      ],
    };
    return {
      reconcile,
      persist: [{ op: 'setHead', head: this.head() }],
    };
  }

  branch(_from: StepRef): EffectEnvelope {
    throw new Error('TimelineEngineImpl.branch is not implemented in Wave 1.');
  }

  switch(_branch: BranchId): EffectEnvelope {
    throw new Error('TimelineEngineImpl.switch is not implemented in Wave 1.');
  }

  // -------------------------------------------------------------------------
  // Additive query methods (NON-breaking — noted in docs/engine-interface.md)
  // -------------------------------------------------------------------------

  /**
   * Read the current Shadow State at an absolute coordinate. Additive query so
   * tests (and the shell's inspectors) can assert the mirror without exposing
   * the internal store.
   */
  readShadow(sheetId: string, row: number, col: number) {
    return this.#shadow.read(sheetId, row, col);
  }

  /** Number of non-empty Shadow State cells held for a sheet. Additive query. */
  shadowCellCount(sheetId: string): number {
    return this.#shadow.cellCount(sheetId);
  }

  /** Sheet metadata (name + tab order) for a sheet, or undefined. Additive query. */
  sheetMeta(sheetId: string) {
    return this.#shadow.sheetMeta(sheetId);
  }

  /** All tracked sheets in tab order. Additive query. */
  shadowSheets() {
    return this.#shadow.sheets();
  }

  /**
   * The next stepIndex that a Present write on a branch would take — i.e. the
   * current tip + 1, or 0 on an empty branch. Additive query; `Head` itself is
   * the frozen shape and does not carry the tip index.
   */
  tipStepIndex(branchId: BranchId = this.#head.branchId): number {
    return this.#nextStepIndex(branchId) - 1;
  }

  /** The recorded Steps on a branch (defensive copy). Additive query. */
  steps(branchId: BranchId = this.#head.branchId): readonly Step[] {
    return [...(this.#log.get(branchId) ?? [])];
  }

  /** The last `ingest` diagnostic, or null. Additive query. */
  lastDiagnostic(): IngestDiagnostic | null {
    return this.#lastDiagnostic;
  }

  /**
   * The stepIndexes at which the engine has written keyframes on a branch, in
   * ascending order. Additive query — lets tests assert the adaptive cadence
   * fired on both the step-count and byte-threshold triggers.
   */
  keyframeIndices(branchId: BranchId = this.#head.branchId): number[] {
    return (this.#keyframes.get(branchId) ?? []).map((kf) => kf.stepIndex);
  }

  /**
   * Reconstruct the Shadow State at `ref` by forward-replay and read a single
   * cell from it. Additive query — lets tests assert replay correctness at
   * arbitrary steps (including across a keyframe boundary) without mutating the
   * live mirror.
   */
  readReconstructed(ref: StepRef, sheetId: string, row: number, col: number) {
    return this.#reconstructAt(ref).read(sheetId, row, col);
  }

  // -------------------------------------------------------------------------
  // Internals
  // -------------------------------------------------------------------------

  #nextStepIndex(branchId: BranchId): number {
    return this.#log.get(branchId)?.length ?? 0;
  }

  #appendStep(branchId: BranchId, step: Step): void {
    const log = this.#log.get(branchId);
    if (log === undefined) {
      this.#log.set(branchId, [step]);
    } else {
      log.push(step);
    }
  }
}
