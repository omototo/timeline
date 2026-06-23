/**
 * Timeline Engine — core types.
 *
 * Engine-neutral, host-agnostic shapes resolved during interface grilling
 * (see `docs/engine-interface.md`, Q1–Q6) and the decisions of record in the
 * ADRs (esp. ADR-0013 and ADR-0001). These types carry NO Office.js, DOM, or
 * React dependency — the engine package stays pure.
 */

// ---------------------------------------------------------------------------
// Geometry & cell primitives (Q2)
// ---------------------------------------------------------------------------

export type SheetId = string;

export interface Rect {
  startRow: number;
  startCol: number;
  rowCount: number;
  colCount: number;
}

/** RangeAreas-aware: one logical change may span disjoint rectangles. */
export type Area = Rect[];

/** Engine-neutral scalar | rich-value JSON. */
export type CellValue = unknown;

export type ValueType = 'empty' | 'string' | 'number' | 'boolean' | 'error' | 'richValue';

export interface CellSlab {
  values: CellValue[][];
  formulas: (string | null)[][];
  numberFormats: string[][];
  valueTypes: ValueType[][];
}

/** Lossless per-cell state used in Value Deltas (before/after). */
export interface CellState {
  value: CellValue;
  formula: string | null;
  valueType: ValueType;
  numberFormat: string;
}

// ---------------------------------------------------------------------------
// Observation — the input boundary (Q2)
// ---------------------------------------------------------------------------

export interface ObservationMeta {
  /** ExcelApi 1.14; 'unknown' on older hosts. */
  triggerSource: 'thisLocalAddin' | 'unknown';
  /** Co-authoring signal. */
  source: 'local' | 'remote';
}

export type StructuralChangeType =
  | 'rowInserted'
  | 'rowDeleted'
  | 'columnInserted'
  | 'columnDeleted'
  | 'cellInserted'
  | 'cellDeleted';

export type ShiftDirection = 'down' | 'right' | 'up' | 'left';

export type WorksheetOp = 'add' | 'delete' | 'rename' | 'reorder';

export interface ValueObservation extends ObservationMeta {
  kind: 'value';
  sheetId: SheetId;
  area: Area;
  after: CellSlab;
}

export interface StructuralObservation extends ObservationMeta {
  kind: 'structural';
  sheetId: SheetId;
  changeType: StructuralChangeType;
  address: Rect;
  /** From changeDirectionState (1.14); inferred on older hosts. */
  shiftDirection?: ShiftDirection;
  // no slab: a structural op is a coordinate transform, not a value change
}

export interface WorksheetObservation extends ObservationMeta {
  kind: 'worksheet';
  op: WorksheetOp;
  sheetId: SheetId;
  newName?: string; // rename
  newPosition?: number; // reorder
}

export type Observation = ValueObservation | StructuralObservation | WorksheetObservation;

// ---------------------------------------------------------------------------
// Identity & history (Q6)
// ---------------------------------------------------------------------------

export type BranchId = string;

/** Ordinal; stable (branches are append-only). */
export interface StepRef {
  branchId: BranchId;
  stepIndex: number;
}

export interface Head {
  branchId: BranchId;
  mode: 'present' | 'preview';
  previewStepIndex?: number;
}

export interface BranchMeta {
  id: BranchId;
  parentBranchId?: BranchId;
  forkedAt?: StepRef;
  order: number;
  name?: string;
  provisional: boolean;
}

// ---------------------------------------------------------------------------
// Deltas (Q6)
// ---------------------------------------------------------------------------

export interface ValueDelta {
  kind: 'value';
  sheetId: SheetId;
  cells: { addr: Rect; before: CellState; after: CellState }[];
}

export interface StructuralDelta {
  kind: 'structural';
  sheetId: SheetId;
  changeType: StructuralChangeType;
  address: Rect;
  shiftDirection?: ShiftDirection;
}

export interface WorksheetDelta {
  kind: 'worksheet';
  op: WorksheetOp;
  sheetId: SheetId;
  newName?: string;
  newPosition?: number;
}

/**
 * One sheet's worth of drift reconciliation (ADR-0006), inspectable. Pinned
 * (Wave 4): per-sheet coordinate-keyed before/after cell states (a Value Delta
 * scoped to one sheet) plus any structural ops applied during reconciliation.
 * Drift currently populates `cells` only (value changes — content, not the
 * untracked coordinate moves that produced it); `structural` is reserved for
 * future structural-drift capture and is currently always empty.
 */
export interface SheetDiff {
  sheetId: SheetId;
  cells: { addr: Rect; before: CellState; after: CellState }[];
  structural: {
    changeType: StructuralChangeType;
    address: Rect;
    shiftDirection?: ShiftDirection;
  }[];
}

/** ADR-0006, inspectable. */
export interface ReconciliationDelta {
  kind: 'reconciliation';
  perSheet: SheetDiff[];
}

export type Delta = ValueDelta | StructuralDelta | WorksheetDelta | ReconciliationDelta;

// ---------------------------------------------------------------------------
// Effects — the output boundary (Q3)
// ---------------------------------------------------------------------------

export type WriteMode = 'value' | 'formula';

export type ReconcileOp =
  | { op: 'setCells'; sheetId: SheetId; area: Area; slab: CellSlab; mode: WriteMode }
  | {
      op: 'applyStructural';
      sheetId: SheetId;
      changeType: StructuralChangeType;
      address: Rect;
      shiftDirection?: ShiftDirection;
    }
  | { op: 'createPreviewSheet'; previewSheetId: SheetId }
  | { op: 'activateSheet'; sheetId: SheetId }
  | { op: 'deletePreviewSheet'; previewSheetId: SheetId };

export interface ReconcilePlan {
  target: 'realSheet' | 'previewSheet';
  ops: ReconcileOp[];
}

export type PersistOp =
  | { op: 'appendDelta'; branchId: BranchId; delta: Delta }
  | { op: 'writeKeyframe'; branchId: BranchId; stepIndex: number; state: unknown /* serialized */ }
  | { op: 'setHead'; head: Head }
  | { op: 'saveBranch'; meta: BranchMeta }
  | { op: 'deleteBranch'; branchId: BranchId };

export interface EffectEnvelope {
  reconcile?: ReconcilePlan;
  persist?: PersistOp[];
}

// ---------------------------------------------------------------------------
// Lifecycle / query placeholders (referenced by the interface; not yet pinned)
// ---------------------------------------------------------------------------

/**
 * `WorkbookSnapshot` — the hashed, observed full-workbook state handed to
 * `attach` for drift comparison (Wave 4, ADR-0006). Pinned: the shell computes
 * the canonical `contentHash`; the engine compares it to the persisted tip hash
 * and itemizes the per-cell diff on drift. Per-sheet slabs are anchored at A1,
 * row-major.
 */
export interface WorkbookSnapshot {
  workbookGuid: string;
  /** Content hash of the observed state, used for clean-resume vs drift. */
  contentHash: string;
  sheets: { sheetId: SheetId; slab: CellSlab }[];
}

/**
 * `PersistedHead` — the resume payload loaded from the store and passed into
 * `attach` (Wave 4): the persisted HEAD plus the stamped tip hash (ADR-0006).
 */
export interface PersistedHead {
  head: Head;
  tipHash: string;
}

/**
 * One branch's persisted history, loaded back from a {@link HistoryStore} for
 * `rehydrate`. `deltas` are in step order (their array index is their
 * stepIndex); `keyframes` are the periodic snapshots persisted via
 * `writeKeyframe` (a fork's base keyframe at stepIndex −1 is NOT persisted — the
 * engine recomputes it from the parent at the fork point).
 */
export interface RehydratedBranch {
  branchId: BranchId;
  deltas: Delta[];
  keyframes: { stepIndex: number; state: unknown }[];
}

/**
 * `RehydrationData` — the full persisted timeline loaded from a
 * {@link HistoryStore} on launch, handed to `rehydrate` to restore the engine's
 * in-memory history (log, branches, keyframes, head) before `attach` reseeds the
 * Shadow State from the live workbook. `branches` excludes the implicit `main`
 * root (which is never persisted as a BranchMeta); `perBranch` includes it.
 */
export interface RehydrationData {
  head: Head | null;
  branches: BranchMeta[];
  perBranch: RehydratedBranch[];
}

/**
 * `TimelineQuery` filters the histogram model (Wave 5 — pinned). Optional branch
 * scope plus an inclusive `[fromStepIndex, toStepIndex]` step window. Filters the
 * returned Steps only; the branch graph is always the full resident fork graph.
 */
export interface TimelineQuery {
  branchId?: BranchId;
  fromStepIndex?: number;
  toStepIndex?: number;
}

/**
 * `TimelineView` — the histogram model returned by `timeline()` (Wave 5 —
 * pinned). `branches` is the full resident fork graph (parent + forkedAt, tab
 * order) so the renderer can draw branch splits; `steps` are the ordered Steps,
 * each with a per-Step `magnitude` (the histogram bar height — see
 * `stepMagnitude`).
 */
export interface TimelineView {
  branches: BranchMeta[];
  steps: {
    ref: StepRef;
    kind: Delta['kind'];
    /** Bar magnitude for the histogram: value=cell count; structural/worksheet=1. */
    magnitude: number;
  }[];
}

/**
 * `StepDetail` — the formula-text metadata returned by `inspectStep()` for the
 * Preview inspect/diff UI (Wave 5 — pinned). For every cell the Step touched,
 * its before/after formula text. `structural`/`worksheet` Steps never rewrite
 * formula text (ADR-0003), so their `cells` list is empty.
 */
export interface StepDetail {
  ref: StepRef;
  kind: Delta['kind'];
  /** Per-cell formula text for the inspect/diff UI. */
  cells: { addr: Rect; beforeFormula: string | null; afterFormula: string | null }[];
}
