/**
 * `@timeline/engine` — public surface.
 *
 * Pure, host-agnostic Timeline Engine types, the engine interface, the
 * persistence-seam ports, and the first concrete store (`InMemoryStore`).
 * No Office.js, DOM, or React (engine-purity wall).
 */

// Core types (geometry, observation, delta, effect, identity).
export type {
  // geometry & cell primitives
  SheetId,
  Rect,
  Area,
  CellValue,
  ValueType,
  CellSlab,
  CellState,
  // observation
  ObservationMeta,
  StructuralChangeType,
  ShiftDirection,
  WorksheetOp,
  ValueObservation,
  StructuralObservation,
  WorksheetObservation,
  Observation,
  // identity & history
  BranchId,
  StepRef,
  Head,
  BranchMeta,
  // deltas
  ValueDelta,
  StructuralDelta,
  WorksheetDelta,
  SheetDiff,
  ReconciliationDelta,
  Delta,
  // effects
  WriteMode,
  ReconcileOp,
  ReconcilePlan,
  PersistOp,
  EffectEnvelope,
  // lifecycle / query placeholders
  WorkbookSnapshot,
  PersistedHead,
  RehydratedBranch,
  RehydrationData,
  TimelineQuery,
  TimelineView,
  StepDetail,
} from './types.ts';

// Engine interface.
export type { TimelineEngine } from './engine.ts';

// Concrete engine (Wave 1: value path; Wave 3: keyframes + navigation) plus its
// additive diagnostic and construction-option shapes.
export {
  TimelineEngineImpl,
  stepMagnitude,
  stepFormulaCells,
  type IngestDiagnostic,
  type TimelineEngineOptions,
} from './timeline-engine.ts';

// Shadow State — in-memory workbook mirror (ADR-0001).
export {
  ShadowState,
  type ChangedCell,
  type SheetMeta,
  type ShadowSnapshot,
} from './shadow-state.ts';

// Reconstruction — forward-replay rebuild (Wave 3).
export { applyDelta, reconstruct } from './reconstruct.ts';

// Projection diff — engine-side minimal ReconcilePlan diff (Wave 3).
// Multi-sheet: one Preview surface per logical sheet (ADR-0005).
// `realSheetDiff` is the formula-mode live counterpart used by `switch` (Wave 4).
export {
  PREVIEW_SHEET_PREFIX,
  previewSheetIdFor,
  projectionDiff,
  realSheetDiff,
} from './project.ts';

// Persistence-seam ports.
export type { HistoryStore, WorkbookStamp, WorkbookStampData } from './ports.ts';

// Concrete store.
export { InMemoryStore } from './in-memory-store.ts';

// Pure cell-diff helper.
export { diffCell, type CellDiff } from './diff-cell.ts';
