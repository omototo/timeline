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
import type { TimelineEngine } from './engine.ts';
import type {
  BranchId,
  Delta,
  EffectEnvelope,
  Head,
  Observation,
  PersistedHead,
  StepDetail,
  StepRef,
  TimelineQuery,
  TimelineView,
  ValueDelta,
  WorkbookSnapshot,
} from './types.ts';

/** The default branch id every workbook starts on. */
const MAIN_BRANCH: BranchId = 'main';

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

    if (obs.kind !== 'value') {
      // Wave 1 implements the value path only; other kinds are recorded as a
      // diagnostic and produce no Step (additive waves implement them).
      this.#lastDiagnostic = {
        code: 'unsupportedKind',
        message: `ingest does not yet handle Observation kind '${obs.kind}' (Wave 1: value path only).`,
      };
      return {};
    }

    const changed = this.#shadow.diff(obs);
    if (changed.length === 0) {
      // No-op observation: nothing changed -> no Step, empty envelope.
      return {};
    }

    const delta: ValueDelta = { kind: 'value', sheetId: obs.sheetId, cells: changed };

    // Append the Step, advance the tip, update the Shadow State forward.
    const branchId = this.#head.branchId;
    const stepIndex = this.#nextStepIndex(branchId);
    this.#appendStep(branchId, { stepIndex, delta });
    this.#shadow.apply(delta);

    // No reconcile in Present (the user already typed it). Persist the delta
    // and the advanced HEAD.
    return {
      persist: [
        { op: 'appendDelta', branchId, delta },
        { op: 'setHead', head: this.head() },
      ],
    };
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

  goto(_ref: StepRef): EffectEnvelope {
    throw new Error('TimelineEngineImpl.goto is not implemented in Wave 1.');
  }

  returnToPresent(): EffectEnvelope {
    throw new Error('TimelineEngineImpl.returnToPresent is not implemented in Wave 1.');
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
