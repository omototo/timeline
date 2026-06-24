/**
 * Projection diff — the engine-side minimal-diff for a {@link ReconcilePlan}.
 *
 * Per Q3 / ADR-0008, `goto` enters Preview by writing **Frozen Values** to a
 * Preview Sheet. The RenderTarget stays dumb: the engine computes the MINIMAL
 * diff between the state it last projected onto the Preview Sheet and the
 * target reconstructed state, here, headlessly. Every cell op is `value` mode
 * (Frozen Values) — Preview never writes live formulas.
 *
 * Pure: no Office.js, DOM, or React.
 */
import { cellStateEquals, EMPTY_CELL, type ShadowState } from './shadow-state.ts';
import type { CellSlab, CellState, ReconcileOp, SheetId } from './types.ts';

/** One cell that differs between the projected and target states. */
interface CellChange {
  row: number;
  col: number;
  /** The target state to write (the empty cell when the cell was cleared). */
  state: CellState;
}

/** A `"row,col"` key for set arithmetic over the union of populated cells. */
function key(row: number, col: number): string {
  return `${String(row)},${String(col)}`;
}

/**
 * The prefix every per-sheet Preview surface id carries. A logical source sheet
 * `Sheet1` projects onto `__preview__::Sheet1`, keeping each sheet's coordinate
 * space distinct (ADR-0005, multi-sheet history) while staying recognizably a
 * preview surface for the shell.
 */
export const PREVIEW_SHEET_PREFIX = '__preview__::';

/** Derive the per-sheet Preview surface id for a logical source `sheetId`. */
export function previewSheetIdFor(sheetId: SheetId): SheetId {
  return `${PREVIEW_SHEET_PREFIX}${sheetId}`;
}

/**
 * Per-sheet minimal cell diff: every coordinate whose state differs between the
 * currently-projected `from` state and the `to` target. A cell present in
 * `from` but absent (empty) in `to` is emitted as a clear (target = empty cell)
 * so the Preview Sheet is brought *exactly* to the target — no stale values.
 */
function sheetCellChanges(from: ShadowState, to: ShadowState, sheetId: SheetId): CellChange[] {
  const fromCells = new Map<string, CellState>();
  for (const c of from.cells(sheetId)) fromCells.set(key(c.row, c.col), c.state);

  const toCells = new Map<string, CellState>();
  for (const c of to.cells(sheetId)) toCells.set(key(c.row, c.col), c.state);

  const changes: CellChange[] = [];

  // Cells in the target: write where they differ from what we projected.
  for (const c of to.cells(sheetId)) {
    const prev = fromCells.get(key(c.row, c.col)) ?? EMPTY_CELL;
    if (!cellStateEquals(prev, c.state)) {
      changes.push({ row: c.row, col: c.col, state: c.state });
    }
  }
  // Cells we projected that the target no longer has: clear them.
  for (const c of from.cells(sheetId)) {
    if (!toCells.has(key(c.row, c.col))) {
      changes.push({ row: c.row, col: c.col, state: { ...EMPTY_CELL } });
    }
  }

  // Deterministic order: row-major.
  changes.sort((a, b) => (a.row - b.row !== 0 ? a.row - b.row : a.col - b.col));
  return changes;
}

/** Wrap a single {@link CellState} as a 1×1 {@link CellSlab}. */
function singleCellSlab(state: CellState): CellSlab {
  return {
    values: [[state.value]],
    formulas: [[state.formula]],
    numberFormats: [[state.numberFormat]],
    valueTypes: [[state.valueType]],
  };
}

/**
 * Compute the minimal `setCells` reconcile ops (value mode) that bring the
 * `from` projection to the `to` target across every populated sheet of either
 * state. One op per changed cell (a single-cell area + 1×1 slab) — minimal and
 * unambiguous; the shell can coalesce adjacent cells if it chooses.
 *
 * MULTI-SHEET (ADR-0005): history is workbook-scoped. Each *logical* source
 * sheet is projected onto its **own** preview surface, whose id is derived from
 * the source sheet id by `previewSheetIdFor`. The diff is computed per logical
 * sheet and each op carries the per-sheet preview id, so coordinates from
 * different logical sheets never collide on a single flat surface (e.g.
 * `Sheet1!A1` and `Sheet2!A1` land on distinct preview sheets).
 *
 * `restrictTo` (full-workbook rollback) bounds the diff to a given set of logical
 * sheets — the sheets that exist at the target step. Without it, the diff spans
 * every populated sheet of either state. A sheet absent from `restrictTo` is
 * skipped: its Preview surface is being deleted, so its cells must not be written.
 */
export function projectionDiff(
  from: ShadowState,
  to: ShadowState,
  restrictTo?: ReadonlySet<SheetId>,
): ReconcileOp[] {
  const sheetIds =
    restrictTo ?? new Set<SheetId>([...from.populatedSheetIds(), ...to.populatedSheetIds()]);
  const ops: ReconcileOp[] = [];
  // Deterministic sheet order.
  for (const sheetId of [...sheetIds].sort()) {
    const previewSheetId = previewSheetIdFor(sheetId);
    for (const change of sheetCellChanges(from, to, sheetId)) {
      ops.push({
        op: 'setCells',
        sheetId: previewSheetId,
        area: [{ startRow: change.row, startCol: change.col, rowCount: 1, colCount: 1 }],
        slab: singleCellSlab(change.state),
        mode: 'value',
      });
    }
  }
  return ops;
}

/**
 * Compute the minimal `setCells` reconcile ops (FORMULA mode) that bring the
 * REAL worksheets from the `from` state to the `to` target — the live-write
 * counterpart of {@link projectionDiff}.
 *
 * Used by `switch` (ADR-0005): checking out another branch tip writes the
 * target state live onto the real sheets, in `formula` mode, so the workbook
 * recalculates (Present is editable; Preview is frozen values). Unlike the
 * Preview projector, ops carry the LOGICAL sheet id (the real worksheet), not a
 * per-sheet preview surface — `switch` lands on the live workbook, not a sandbox.
 */
export function realSheetDiff(from: ShadowState, to: ShadowState): ReconcileOp[] {
  const sheetIds = new Set<SheetId>([...from.populatedSheetIds(), ...to.populatedSheetIds()]);
  const ops: ReconcileOp[] = [];
  for (const sheetId of [...sheetIds].sort()) {
    for (const change of sheetCellChanges(from, to, sheetId)) {
      ops.push({
        op: 'setCells',
        sheetId,
        area: [{ startRow: change.row, startCol: change.col, rowCount: 1, colCount: 1 }],
        slab: singleCellSlab(change.state),
        mode: 'formula',
      });
    }
  }
  return ops;
}
