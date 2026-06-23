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
import { ShadowState, EMPTY_CELL, cellStateEquals } from './shadow-state.ts';
import type { ShadowSnapshot } from './shadow-state.ts';
import { reconstruct } from './reconstruct.ts';
import { previewSheetIdFor, projectionDiff, realSheetDiff } from './project.ts';
import type { TimelineEngine } from './engine.ts';
import type {
  BranchId,
  BranchMeta,
  CellSlab,
  CellState,
  Delta,
  EffectEnvelope,
  Head,
  Observation,
  PersistedHead,
  PersistOp,
  ReconcileOp,
  ReconcilePlan,
  ReconciliationDelta,
  Rect,
  SheetDiff,
  SheetId,
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

/** Default adaptive-keyframe cadence (Q6): step count + cumulative delta bytes. */
const DEFAULT_KEYFRAME_STEP_INTERVAL = 100;
const DEFAULT_KEYFRAME_BYTE_THRESHOLD = 64 * 1024;

/**
 * The stepIndex of a branch's BASE keyframe — the seed state captured at a fork
 * point, before the branch records its first Step (which is index 0). A
 * negative index keeps it strictly below every real step so `keyframeAtOrBefore`
 * picks it up for any target ≥ 0 but it never collides with a recorded Step.
 */
const BASE_KEYFRAME_INDEX = -1;

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
 * Diagnostic raised when an operation is refused by the engine's state machine
 * (returns a no-op envelope, never corrupts state). Wave 1 carried the
 * `ingest`-in-Preview and `unsupportedKind` codes; Wave 4 adds:
 *
 * - `ingestSuspended` — an Observation arrived while tracking is suspended for
 *   co-authoring (`detachToCoauthoring`, ADR-0006). Ingest is a no-op until the
 *   engine re-attaches.
 * - `coauthoringSuspended` — emitted by `detachToCoauthoring` itself to tell the
 *   shell tracking has been disabled (a `source: 'remote'` edit was seen).
 */
export interface IngestDiagnostic {
  code: 'ingestInPreview' | 'unsupportedKind' | 'ingestSuspended' | 'coauthoringSuspended';
  message: string;
}

export class TimelineEngineImpl implements TimelineEngine {
  #shadow = new ShadowState();
  /** Per-branch append-only Step log (resident; deltas are small — ADR-0007). */
  readonly #log = new Map<BranchId, Step[]>();
  #head: Head = { branchId: MAIN_BRANCH, mode: 'present' };
  /** Last diagnostic from a refused/no-op `ingest`, for the shell to surface. */
  #lastDiagnostic: IngestDiagnostic | null = null;

  // --- Branching & lifecycle (Wave 4) -------------------------------------
  /**
   * Per-branch metadata (ADR-0005). `main` is implicit and non-provisional; it
   * is registered lazily (a fork records its parent, so `main` must exist as a
   * `BranchMeta` once any branch is created). A PROVISIONAL branch (from
   * `branch`) is held resident but NOT persisted until its first `ingest`
   * (`saveBranch`), and is garbage-collected if `switch`-ed away from before it
   * records a Step (ADR-0005 — provisional branches are cheap, disposable forks).
   */
  readonly #branches = new Map<BranchId, BranchMeta>();
  /** Monotonic counter assigning each new branch its tab `order`. */
  #branchOrder = 0;
  /**
   * Branch ids whose `saveBranch` has been emitted. A provisional branch is
   * persisted lazily on its first ingest; tracking this avoids re-emitting
   * `saveBranch` on every subsequent Step.
   */
  readonly #persistedBranches = new Set<BranchId>();
  /**
   * Tracking suspended for co-authoring (ADR-0006). Set by
   * `detachToCoauthoring`; while true, `ingest` is a no-op + diagnostic. Cleared
   * by a clean `attach`.
   */
  #suspended = false;
  /** Monotonic counter behind the deterministic minted branch ids. */
  #branchSeq = 0;

  // --- Adaptive keyframe cadence (Q6) -------------------------------------
  readonly #keyframeStepInterval: number;
  readonly #keyframeByteThreshold: number;
  /** Per-branch resident keyframes, step-ascending. */
  readonly #keyframes = new Map<BranchId, Keyframe[]>();
  /** Per-branch steps appended since that branch's last keyframe. */
  readonly #stepsSinceKeyframe = new Map<BranchId, number>();
  /** Per-branch cumulative delta bytes appended since its last keyframe. */
  readonly #bytesSinceKeyframe = new Map<BranchId, number>();

  // --- Navigation: currently-projected Preview state ----------------------
  /**
   * The Shadow State the engine last projected onto the Preview surfaces, or
   * `null` when no Preview is active. `goto` diffs the target against this to
   * produce a MINIMAL value-mode plan; `returnToPresent` clears it. This is the
   * FULL multi-sheet target (every logical sheet), diffed per logical sheet.
   */
  #projected: ShadowState | null = null;

  /**
   * The per-sheet Preview surfaces created during the current Preview session
   * (one per logical sheet projected so far), in creation order. `goto` adds a
   * `createPreviewSheet` for any logical sheet not yet seen; `returnToPresent`
   * emits a `deletePreviewSheet` for each and clears this.
   */
  #previewSurfaces: SheetId[] = [];

  /**
   * The real worksheet that was active when the current Preview session began,
   * captured at first `goto`. `returnToPresent` reactivates it. `null` when the
   * present state has no populated/known sheet to reactivate, in which case
   * `returnToPresent` omits the `activateSheet` op and lets the shell restore
   * the previously-active real sheet itself.
   */
  #activeRealSheetId: SheetId | null = null;

  constructor(options: TimelineEngineOptions = {}) {
    this.#keyframeStepInterval = options.keyframeStepInterval ?? DEFAULT_KEYFRAME_STEP_INTERVAL;
    this.#keyframeByteThreshold = options.keyframeByteThreshold ?? DEFAULT_KEYFRAME_BYTE_THRESHOLD;
  }

  // -------------------------------------------------------------------------
  // Recording — Present mode only (Wave 1: value path)
  // -------------------------------------------------------------------------

  ingest(obs: Observation): EffectEnvelope {
    this.#lastDiagnostic = null;

    // Suspended for co-authoring (ADR-0006): tracking is disabled until a clean
    // re-attach. Every Observation is a no-op while suspended.
    if (this.#suspended) {
      this.#lastDiagnostic = {
        code: 'ingestSuspended',
        message: 'Tracking is suspended for co-authoring; ingest is a no-op until re-attach.',
      };
      return {};
    }

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
   *
   * A PROVISIONAL branch persists lazily (ADR-0005): its `saveBranch` PersistOp
   * is prepended to the envelope on its FIRST recorded Step, not at `branch()`
   * time — a fork that is abandoned before any edit never touches the store.
   */
  #recordStep(delta: Delta, applyToShadow: () => void): EffectEnvelope {
    const branchId = this.#head.branchId;
    const stepIndex = this.#nextStepIndex(branchId);
    this.#appendStep(branchId, { stepIndex, delta });
    applyToShadow();

    const persist: PersistOp[] = [];

    // First Step on a not-yet-persisted branch: persist its BranchMeta now. A
    // provisional fork promotes to a real, saved branch on first edit.
    const saveBranchOp = this.#saveBranchOnFirstStep(branchId);
    if (saveBranchOp !== null) persist.push(saveBranchOp);

    persist.push({ op: 'appendDelta', branchId, delta }, { op: 'setHead', head: this.head() });

    // Adaptive keyframe cadence (Q6): account this Step's cost, then — if EITHER
    // the step-count OR the byte threshold is crossed — snapshot the Shadow
    // State (now at `stepIndex`) and emit a writeKeyframe op. The snapshot is
    // also kept resident so reconstruction can replay forward from it.
    const steps = (this.#stepsSinceKeyframe.get(branchId) ?? 0) + 1;
    const bytes = (this.#bytesSinceKeyframe.get(branchId) ?? 0) + deltaBytes(delta);
    if (steps >= this.#keyframeStepInterval || bytes >= this.#keyframeByteThreshold) {
      const snapshot = this.#shadow.snapshot();
      this.#storeKeyframe(branchId, { stepIndex, snapshot });
      this.#stepsSinceKeyframe.set(branchId, 0);
      this.#bytesSinceKeyframe.set(branchId, 0);
      persist.push({ op: 'writeKeyframe', branchId, stepIndex, state: snapshot });
    } else {
      this.#stepsSinceKeyframe.set(branchId, steps);
      this.#bytesSinceKeyframe.set(branchId, bytes);
    }

    return { persist };
  }

  /**
   * Return the `saveBranch` PersistOp to emit on a branch's first Step, or null
   * if there is nothing to persist. The implicit `main` root is NOT a saved
   * `BranchMeta` — it always exists, so `main`'s first Step emits no
   * `saveBranch`. A PROVISIONAL fork, by contrast, persists lazily on its first
   * Step (ADR-0005): we emit its `saveBranch` once and mark it persisted +
   * non-provisional so subsequent Steps do not re-emit.
   */
  #saveBranchOnFirstStep(branchId: BranchId): Extract<PersistOp, { op: 'saveBranch' }> | null {
    if (branchId === MAIN_BRANCH) return null;
    if (this.#persistedBranches.has(branchId)) return null;
    const meta = this.#branches.get(branchId);
    if (meta === undefined) return null;
    this.#persistedBranches.add(branchId);
    // The branch is no longer provisional once it is persisted.
    const persistedMeta: BranchMeta = { ...meta, provisional: false };
    this.#branches.set(branchId, persistedMeta);
    return { op: 'saveBranch', meta: persistedMeta };
  }

  /**
   * The {@link BranchMeta} for a branch, registering implicit `main` lazily.
   * `main` is the workbook's root branch (non-provisional, order 0); it has no
   * `BranchMeta` until something needs one (a fork records `main` as a parent,
   * or `main`'s first Step persists it).
   */
  #branchMeta(branchId: BranchId): BranchMeta | null {
    const existing = this.#branches.get(branchId);
    if (existing !== undefined) return existing;
    if (branchId === MAIN_BRANCH) {
      const meta: BranchMeta = { id: MAIN_BRANCH, order: 0, provisional: false };
      this.#branches.set(MAIN_BRANCH, meta);
      this.#branchOrder = Math.max(this.#branchOrder, 1);
      return meta;
    }
    return null;
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
  // Lifecycle (Wave 4)
  // -------------------------------------------------------------------------

  /**
   * Attach to a workbook on launch / pane-open (ADR-0006). Hash the observed
   * live state and compare it to the persisted Tip hash:
   *
   * - **No persisted head** — a fresh workbook. Seed the Shadow State from the
   *   observed slabs (so future ingests diff against real content), restore
   *   nothing, empty reconcile.
   * - **Clean match** (`observed.contentHash === persisted.tipHash`) — the
   *   workbook is exactly where history left it. Restore HEAD from the persisted
   *   head, resume tracking, empty reconcile (no writes — the live state already
   *   equals the tip).
   * - **Drift** — the workbook changed behind the engine's back. Compute an
   *   itemized, per-sheet before/after {@link ReconciliationDelta} (the engine's
   *   Shadow State is "before"; the observed slabs are "after"), append it as a
   *   single inspectable **Reconciliation Step** (ADR-0006), advance the Shadow
   *   State to the observed state, and restore HEAD. We capture WHAT changed,
   *   not the sequence of untracked edits. Empty reconcile (no write-back — the
   *   user's current work is authoritative; pre-drift history stays previewable).
   *
   * A successful attach always clears any co-authoring suspension — re-attaching
   * resumes tracking.
   */
  attach(observed: WorkbookSnapshot, persisted: PersistedHead | null): EffectEnvelope {
    this.#lastDiagnostic = null;
    this.#suspended = false;

    if (persisted === null) {
      // Fresh workbook: seed the mirror from the observed content; no Step.
      this.#shadow = snapshotToShadow(observed);
      return {};
    }

    if (observed.contentHash === persisted.tipHash) {
      // Clean resume: restore HEAD, no writes.
      this.#head = restoreHead(persisted.head);
      return { persist: [{ op: 'setHead', head: this.head() }] };
    }

    // Drift: build the itemized per-sheet diff (Shadow "before" vs observed
    // "after"), append it as a Reconciliation Step, and advance the mirror.
    const perSheet = this.#computeDriftDiff(observed);
    const delta: ReconciliationDelta = { kind: 'reconciliation', perSheet };
    // A drift reconciliation is a Present-mode history mutation: restore the
    // branch from the persisted head but FORCE Present mode. A persisted PREVIEW
    // head is reachable (goto emits a preview setHead the shell persists), and
    // appending the Reconciliation Step under a stale preview head would emit an
    // inconsistent preview setHead (its previewStepIndex now points BEFORE the
    // just-appended Step) and leave the engine stuck refusing the next ingest
    // with `ingestInPreview`. Landing in Present keeps tracking usable.
    this.#head = { branchId: restoreHead(persisted.head).branchId, mode: 'present' };
    const envelope = this.#recordStep(delta, () => {
      this.#shadow = snapshotToShadow(observed);
    });
    return envelope;
  }

  /**
   * Compute the per-sheet reconciliation diff between the engine's Shadow State
   * (the "before" the engine last witnessed) and the observed live workbook
   * (the "after"), over the union of sheets present in either. Only cells that
   * actually changed are recorded; a sheet with no changes is omitted. Value
   * changes only — drift reconciliation captures content, not the untracked
   * coordinate moves that produced it (ADR-0006).
   */
  #computeDriftDiff(observed: WorkbookSnapshot): SheetDiff[] {
    const observedById = new Map(observed.sheets.map((s) => [s.sheetId, s.slab]));
    const sheetIds = new Set<SheetId>([
      ...this.#shadow.populatedSheetIds(),
      ...observed.sheets.map((s) => s.sheetId),
    ]);

    const perSheet: SheetDiff[] = [];
    for (const sheetId of [...sheetIds].sort()) {
      const after = observedById.get(sheetId);
      const cells = this.#sheetDriftCells(sheetId, after);
      if (cells.length > 0) {
        perSheet.push({ sheetId, cells, structural: [] });
      }
    }
    return perSheet;
  }

  /**
   * Per-sheet drift cells: compare every coordinate populated in either the
   * Shadow State or the observed slab, recording `{ addr, before, after }` for
   * those that differ. A coordinate present in the slab is read from it; a
   * coordinate only in the Shadow State is now empty (it was cleared by the
   * untracked edits).
   */
  #sheetDriftCells(
    sheetId: SheetId,
    afterSlab: CellSlab | undefined,
  ): { addr: Rect; before: CellState; after: CellState }[] {
    const before = new Map<string, CellState>();
    for (const c of this.#shadow.cells(sheetId)) {
      before.set(`${String(c.row)},${String(c.col)}`, c.state);
    }
    const after = afterSlab === undefined ? new Map<string, CellState>() : slabToCellMap(afterSlab);

    const coords = new Set<string>([...before.keys(), ...after.keys()]);
    const out: { addr: Rect; before: CellState; after: CellState }[] = [];
    for (const coord of [...coords].sort(byRowMajor)) {
      const beforeState = before.get(coord) ?? { ...EMPTY_CELL };
      const afterState = after.get(coord) ?? { ...EMPTY_CELL };
      if (!cellStateEquals(beforeState, afterState)) {
        const [row, col] = coord.split(',').map(Number) as [number, number];
        out.push({
          addr: { startRow: row, startCol: col, rowCount: 1, colCount: 1 },
          before: beforeState,
          after: afterState,
        });
      }
    }
    return out;
  }

  /**
   * Suspend tracking for co-authoring (ADR-0006). A `source: 'remote'` edit
   * means the file is shared; branching history + multi-author merge are two
   * products, so v1 disables tracking with a clear diagnostic rather than
   * corrupting state. Subsequent `ingest` calls are no-ops until a clean
   * `attach` resumes tracking. NOT a Step (no history mutation).
   */
  detachToCoauthoring(): EffectEnvelope {
    this.#suspended = true;
    this.#lastDiagnostic = {
      code: 'coauthoringSuspended',
      message: 'Co-authoring detected (source: remote); tracking suspended for this session.',
    };
    return {};
  }

  /**
   * Enter (or move within) Preview at `ref` (ADR-0008, Q3). Reconstruct the
   * target state by forward-replay, then emit a {@link ReconcilePlan} that is
   * the MINIMAL value-mode diff between the engine's currently-projected state
   * and the target.
   *
   * MULTI-SHEET (ADR-0005): each logical source sheet projects onto its OWN
   * preview surface (`previewSheetIdFor(sheetId)`), so coordinates from
   * different sheets never collide. The first time a given logical sheet is
   * projected, the plan is prefixed with a `createPreviewSheet` for its surface;
   * the very first surface created is also `activateSheet`d. HEAD flips to
   * `preview` and the FULL multi-sheet projected state is tracked so a
   * subsequent `goto`/scrub writes only what changed, per logical sheet.
   */
  goto(ref: StepRef): EffectEnvelope {
    const firstEntry = this.#projected === null;
    if (firstEntry) {
      // Capture the active real sheet so returnToPresent can reactivate it.
      this.#activeRealSheetId = this.#presentActiveSheetId();
    }
    const target = this.#reconstructAt(ref);
    const from = this.#projected ?? new ShadowState();

    // Create any preview surface we have not created yet this session — one per
    // logical sheet touched by either the projected `from` or the `to` target.
    const ops: ReconcileOp[] = [];
    const touchedSheets = new Set<SheetId>([
      ...from.populatedSheetIds(),
      ...target.populatedSheetIds(),
    ]);
    for (const sheetId of [...touchedSheets].sort()) {
      if (!this.#previewSurfaces.includes(sheetId)) {
        const previewSheetId = previewSheetIdFor(sheetId);
        ops.push({ op: 'createPreviewSheet', previewSheetId });
        // Activate the first preview surface created in this session so the
        // shell lands the user on a visible preview sheet.
        if (this.#previewSurfaces.length === 0) {
          ops.push({ op: 'activateSheet', sheetId: previewSheetId });
        }
        this.#previewSurfaces.push(sheetId);
      }
    }
    ops.push(...projectionDiff(from, target));

    this.#projected = target;
    this.#head = { branchId: ref.branchId, mode: 'preview', previewStepIndex: ref.stepIndex };

    const reconcile: ReconcilePlan = { target: 'previewSheet', ops };
    return {
      reconcile,
      persist: [{ op: 'setHead', head: this.head() }],
    };
  }

  /**
   * The real worksheet the engine considers active in Present, used to restore
   * focus on `returnToPresent`. Prefers a registered sheet in tab order, else
   * the lexicographically-first populated sheet, else `null` (nothing to
   * reactivate — the shell restores the previously-active sheet itself).
   */
  #presentActiveSheetId(): SheetId | null {
    const registered = this.#shadow.sheets();
    if (registered.length > 0 && registered[0] !== undefined) {
      return registered[0].sheetId;
    }
    const populated = [...this.#shadow.populatedSheetIds()].sort();
    return populated[0] ?? null;
  }

  /**
   * Discard every Preview surface and reactivate the real sheet (Q3), returning
   * HEAD to Present. Emits a {@link ReconcilePlan} targeting the real sheet that
   * deletes each per-sheet Preview surface created this session, then activates
   * the worksheet that was active when Preview began.
   *
   * The `activateSheet` op carries a REAL {@link SheetId} (the active sheet
   * captured at `goto`-time), never a {@link BranchId} — branch ids and sheet
   * ids are distinct namespaces. If no active real sheet was knowable, the
   * `activateSheet` op is omitted and the shell restores the previously-active
   * sheet itself. A no-op (empty envelope) when not currently in Preview.
   */
  returnToPresent(): EffectEnvelope {
    if (this.#projected === null || this.#head.mode !== 'preview') {
      return {};
    }
    const branchId = this.#head.branchId;
    const surfaces = this.#previewSurfaces;
    const activeRealSheetId = this.#activeRealSheetId;

    this.#projected = null;
    this.#previewSurfaces = [];
    this.#activeRealSheetId = null;
    this.#head = { branchId, mode: 'present' };

    const ops: ReconcileOp[] = surfaces.map((sheetId) => ({
      op: 'deletePreviewSheet',
      previewSheetId: previewSheetIdFor(sheetId),
    }));
    if (activeRealSheetId !== null) {
      ops.push({ op: 'activateSheet', sheetId: activeRealSheetId });
    }

    const reconcile: ReconcilePlan = { target: 'realSheet', ops };
    return {
      reconcile,
      persist: [{ op: 'setHead', head: this.head() }],
    };
  }

  /**
   * "Branch from here" (ADR-0005): fork a new PROVISIONAL branch at `from` and
   * promote it to an editable Present (HEAD → the new branch tip).
   *
   * The fork's state is reconstructed at `from` by forward-replay and seeded as
   * a base keyframe (at stepIndex −1) on the new branch, so the new branch's own
   * reconstruction/replay starts from the fork point. The Shadow State becomes
   * that reconstructed state immediately (the user edits the fork live). HEAD
   * flips to the new branch in `present`.
   *
   * Provisional means NOT yet persisted: `branch()` emits NO `saveBranch` (only
   * a `setHead`). The branch persists lazily on its first `ingest` (ADR-0005),
   * or is garbage-collected if `switch`-ed away from before recording a Step.
   * If a Preview is active, it is implicitly returned to Present first (a fork
   * is a Present operation).
   */
  branch(from: StepRef): EffectEnvelope {
    // A fork is a Present op; abandon any active Preview projection silently.
    this.#clearPreview();

    // Ensure the parent branch is registered so the fork records a real parent.
    const parentMeta = this.#branchMeta(from.branchId);
    const parentId = parentMeta?.id ?? from.branchId;

    const newBranchId = this.#mintBranchId();
    const order = this.#branchOrder++;
    const meta: BranchMeta = {
      id: newBranchId,
      parentBranchId: parentId,
      forkedAt: { branchId: from.branchId, stepIndex: from.stepIndex },
      order,
      provisional: true,
    };
    this.#branches.set(newBranchId, meta);

    // Reconstruct the fork point and seed it as the new branch's base keyframe
    // (stepIndex −1: the state before the branch's first Step) so replay on the
    // new branch starts from the fork.
    const forkState = this.#reconstructAt(from);
    this.#log.set(newBranchId, []);
    this.#storeKeyframe(newBranchId, {
      stepIndex: BASE_KEYFRAME_INDEX,
      snapshot: forkState.snapshot(),
    });
    this.#stepsSinceKeyframe.set(newBranchId, 0);
    this.#bytesSinceKeyframe.set(newBranchId, 0);

    this.#shadow = forkState;
    this.#head = { branchId: newBranchId, mode: 'present' };

    return { persist: [{ op: 'setHead', head: this.head() }] };
  }

  /**
   * Checkout another branch's tip (ADR-0005). NAVIGATION, not a Step: it
   * reconstructs the target branch's tip state, makes it the live Shadow State,
   * and emits a `formula`-mode {@link ReconcilePlan} onto `realSheet` (live —
   * Present is editable, so the diff writes formulas, not frozen values). HEAD
   * flips to the target in `present`. Never appends a delta.
   *
   * PROVISIONAL GC (ADR-0005): if we are switching AWAY from a provisional
   * branch that has recorded zero Steps, it is discarded — the branch + its
   * resident log/keyframes are dropped and a `deleteBranch` PersistOp is
   * emitted. A no-op (empty envelope) when the target is already the current
   * branch in Present.
   */
  switch(branch: BranchId): EffectEnvelope {
    // A switch is a Present op; abandon any active Preview projection first.
    this.#clearPreview();

    const fromBranchId = this.#head.branchId;
    if (branch === fromBranchId) {
      // Already on this branch in Present: nothing to do.
      return {};
    }

    const persist: PersistOp[] = [];

    // Provisional GC: discard a zero-Step provisional branch we are leaving.
    const gcOp = this.#gcProvisionalBranch(fromBranchId);
    if (gcOp !== null) persist.push(gcOp);

    // Reconstruct the target tip and diff the live sheets onto it (formula mode).
    const from = this.#shadow;
    const target = this.#reconstructTip(branch);
    const ops = realSheetDiff(from, target);

    this.#shadow = target;
    this.#head = { branchId: branch, mode: 'present' };

    persist.push({ op: 'setHead', head: this.head() });
    return { reconcile: { target: 'realSheet', ops }, persist };
  }

  /**
   * If `branchId` is a PROVISIONAL branch with zero recorded Steps, drop its
   * resident state (log, keyframes, cadence counters, meta) and return the
   * `deleteBranch` PersistOp to emit; otherwise null. Used by `switch` to GC a
   * fork abandoned before its first edit.
   */
  #gcProvisionalBranch(branchId: BranchId): Extract<PersistOp, { op: 'deleteBranch' }> | null {
    const meta = this.#branches.get(branchId);
    if (!meta?.provisional) return null;
    if ((this.#log.get(branchId)?.length ?? 0) > 0) return null; // has Steps: keep
    this.#branches.delete(branchId);
    this.#log.delete(branchId);
    this.#keyframes.delete(branchId);
    this.#stepsSinceKeyframe.delete(branchId);
    this.#bytesSinceKeyframe.delete(branchId);
    this.#persistedBranches.delete(branchId);
    return { op: 'deleteBranch', branchId };
  }

  /** Reconstruct a branch's TIP state by forward-replay from its last keyframe. */
  #reconstructTip(branchId: BranchId): ShadowState {
    const tip = this.#nextStepIndex(branchId) - 1;
    if (tip < 0) {
      // No Steps: seed from the branch's base keyframe if present, else empty.
      const base = this.#keyframeAtOrBefore(branchId, BASE_KEYFRAME_INDEX);
      return base === null ? new ShadowState() : ShadowState.fromSnapshot(base.snapshot);
    }
    return this.#reconstructAt({ branchId, stepIndex: tip });
  }

  /**
   * Tear down any active Preview projection WITHOUT emitting effects (internal).
   * Used by `branch`/`switch`, which are Present operations: they leave Preview
   * implicitly. The shell will overwrite the surfaces via the new plan; the
   * Preview surfaces themselves are the shell's concern on a mode switch.
   */
  #clearPreview(): void {
    this.#projected = null;
    this.#previewSurfaces = [];
    this.#activeRealSheetId = null;
    if (this.#head.mode === 'preview') {
      this.#head = { branchId: this.#head.branchId, mode: 'present' };
    }
  }

  /** Mint the next deterministic branch id (`branch-1`, `branch-2`, …). */
  #mintBranchId(): BranchId {
    return `branch-${String(++this.#branchSeq)}`;
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

  /**
   * The {@link BranchMeta} of every branch the engine knows (resident),
   * tab-order ascending (defensive copies). Additive query (Wave 4) — lets the
   * shell/tests inspect the branch graph (fork points, provisional flags)
   * without exposing internals.
   */
  branches(): BranchMeta[] {
    return [...this.#branches.values()].sort((a, b) => a.order - b.order).map((m) => ({ ...m }));
  }

  /** Whether a branch is currently held resident (not GC'd). Additive query. */
  hasBranch(branchId: BranchId): boolean {
    return this.#branches.has(branchId);
  }

  /** Whether tracking is suspended for co-authoring (ADR-0006). Additive query. */
  isSuspended(): boolean {
    return this.#suspended;
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

// ---------------------------------------------------------------------------
// Module-level helpers (Wave 4 lifecycle)
// ---------------------------------------------------------------------------

/** Restore a persisted {@link Head}, defensively normalized (omit preview index when absent). */
function restoreHead(head: Head): Head {
  return head.mode === 'preview' && head.previewStepIndex !== undefined
    ? { branchId: head.branchId, mode: 'preview', previewStepIndex: head.previewStepIndex }
    : { branchId: head.branchId, mode: head.mode };
}

/**
 * Seed a fresh {@link ShadowState} from an observed {@link WorkbookSnapshot} —
 * every non-empty cell of every observed sheet slab. Used by `attach` to align
 * the mirror with the live workbook (fresh open, or after drift reconciliation).
 */
function snapshotToShadow(observed: WorkbookSnapshot): ShadowState {
  const state = new ShadowState();
  for (const sheet of observed.sheets) {
    const cells = slabToCellMap(sheet.slab);
    const valueCells: ValueDelta['cells'] = [];
    for (const [coord, cellState] of cells) {
      const [row, col] = coord.split(',').map(Number) as [number, number];
      valueCells.push({
        addr: { startRow: row, startCol: col, rowCount: 1, colCount: 1 },
        before: { ...EMPTY_CELL },
        after: cellState,
      });
    }
    if (valueCells.length > 0) {
      state.apply({ kind: 'value', sheetId: sheet.sheetId, cells: valueCells });
    }
  }
  return state;
}

/**
 * Flatten a {@link CellSlab} (anchored at A1, row-major) into a `"row,col" ->
 * CellState` map, dropping cells that are equal to the canonical empty cell.
 * The slab is assumed to be a single rectangle starting at (0,0) — the shape
 * `attach`'s {@link WorkbookSnapshot} carries per sheet.
 */
function slabToCellMap(slab: CellSlab): Map<string, CellState> {
  const out = new Map<string, CellState>();
  for (let r = 0; r < slab.values.length; r++) {
    const valuesRow = slab.values[r] ?? [];
    const formulasRow = slab.formulas[r] ?? [];
    const numberFormatsRow = slab.numberFormats[r] ?? [];
    const valueTypesRow = slab.valueTypes[r] ?? [];
    for (let c = 0; c < valuesRow.length; c++) {
      const cell: CellState = {
        value: valuesRow[c] ?? '',
        formula: formulasRow[c] ?? null,
        valueType: valueTypesRow[c] ?? 'empty',
        numberFormat: numberFormatsRow[c] ?? 'General',
      };
      if (!cellStateEquals(cell, EMPTY_CELL)) {
        out.set(`${String(r)},${String(c)}`, cell);
      }
    }
  }
  return out;
}

/** Compare two `"row,col"` keys row-major (ascending row, then column). */
function byRowMajor(a: string, b: string): number {
  const [ar, ac] = a.split(',').map(Number) as [number, number];
  const [br, bc] = b.split(',').map(Number) as [number, number];
  return ar - br !== 0 ? ar - br : ac - bc;
}
