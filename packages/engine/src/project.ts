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
 * Targets the Preview Sheet conceptually: the caller stamps the actual
 * `sheetId` (the Preview Sheet's id) onto each op. We diff *logical* sheets
 * here but emit against the single preview surface, so the `previewSheetId`
 * argument overrides the source sheet id on every op.
 */
export function projectionDiff(
  from: ShadowState,
  to: ShadowState,
  previewSheetId: SheetId,
): ReconcileOp[] {
  const sheetIds = new Set<SheetId>([...from.populatedSheetIds(), ...to.populatedSheetIds()]);
  const ops: ReconcileOp[] = [];
  // Deterministic sheet order.
  for (const sheetId of [...sheetIds].sort()) {
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
