/**
 * Timeline Engine — verb surface (Q4).
 *
 * Functional core, imperative shell (ADR-0013): the engine is a synchronous,
 * stateful, in-memory instance. Every mutator returns an `EffectEnvelope`
 * describing the I/O the shell must perform; the engine never performs or
 * awaits I/O itself. Queries are pure.
 *
 * This is the interface ONLY — the engine algorithm is implemented behind it
 * later.
 */
import type {
  BranchId,
  EffectEnvelope,
  Head,
  Observation,
  PersistedHead,
  StepDetail,
  StepRef,
  TimelineQuery,
  TimelineView,
  WorkbookSnapshot,
} from './types.ts';

export interface TimelineEngine {
  // Lifecycle
  /** hash+compare; clean resume OR drift -> Reconciliation Step. */
  attach(observed: WorkbookSnapshot, persisted: PersistedHead | null): EffectEnvelope;
  /** on source:'remote' -> suspend tracking (ADR-0006). */
  detachToCoauthoring(): EffectEnvelope;

  // Recording — Present mode only; if an Observation arrives in Preview the
  // engine returns a no-op + diagnostic (shell must lock the real sheet during
  // Preview).
  ingest(obs: Observation): EffectEnvelope;

  // Navigation
  /** enter Preview (frozen values, fresh Preview Sheet). */
  goto(ref: StepRef): EffectEnvelope;
  /** discard Preview Sheet, reactivate real sheet. */
  returnToPresent(): EffectEnvelope;
  /** "Branch from here" -> provisional editable Present. */
  branch(from: StepRef): EffectEnvelope;
  /** checkout another branch tip (non-destructive, NOT a Step). */
  switch(branch: BranchId): EffectEnvelope;

  // Queries (pure)
  head(): Head;
  /** histogram model: steps, bar magnitudes, branch splits. */
  timeline(opts?: TimelineQuery): TimelineView;
  /** formula text metadata for Preview. */
  inspectStep(ref: StepRef): StepDetail;
}
