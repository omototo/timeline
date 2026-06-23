/**
 * Shadow State — the engine's in-memory workbook mirror (ADR-0001).
 *
 * The engine, not Excel, owns history; the Shadow State is its source of truth.
 * It mirrors the live workbook per-sheet, mapping a cell coordinate to a
 * lossless {@link CellState} (value, formula, valueType, numberFormat —
 * ADR-0008). Office.js change events are only a trigger + bounding box; the
 * adapter reads back the changed area and hands the engine a
 * {@link ValueObservation}, which the engine diffs against this mirror to
 * produce a sparse {@link ValueDelta} (only the cells that actually changed),
 * then applies forward.
 *
 * Pure: no Office.js, DOM, or React. Coordinates are absolute `(row, col)`,
 * 0-based, matching the {@link Rect} geometry of an Observation's `area`.
 */
import type {
  Area,
  CellSlab,
  CellState,
  Rect,
  SheetId,
  ValueDelta,
  ValueObservation,
  ValueType,
} from './types.ts';

/** A changed cell: its absolute address plus before/after lossless state. */
export interface ChangedCell {
  addr: Rect;
  before: CellState;
  after: CellState;
}

/** The lossless state of an empty (never-written / cleared) cell. */
const EMPTY_CELL: CellState = {
  value: '',
  formula: null,
  valueType: 'empty',
  numberFormat: 'General',
};

/** Coordinate key for the per-sheet cell map: `"row,col"`. */
type CellKey = string;

function cellKey(row: number, col: number): CellKey {
  return `${String(row)},${String(col)}`;
}

/** A single-cell {@link Rect} at an absolute coordinate. */
function cellRect(row: number, col: number): Rect {
  return { startRow: row, startCol: col, rowCount: 1, colCount: 1 };
}

/** Two {@link CellState}s are equal iff every lossless field matches. */
function cellStateEquals(a: CellState, b: CellState): boolean {
  return (
    Object.is(a.value, b.value) &&
    a.formula === b.formula &&
    a.valueType === b.valueType &&
    a.numberFormat === b.numberFormat
  );
}

/** Default an empty cell to the canonical {@link EMPTY_CELL} (immutable copy). */
function emptyCell(): CellState {
  return { ...EMPTY_CELL };
}

/**
 * In-memory workbook mirror. Per-sheet, coordinate-keyed lossless cell state.
 *
 * Provides the three Shadow State ops the value path needs:
 * - {@link ShadowState.read} — current state at a coordinate.
 * - {@link ShadowState.diff} — Observation vs mirror -> sparse changed cells.
 * - {@link ShadowState.apply} — push a {@link ValueDelta} forward into the mirror.
 */
export class ShadowState {
  /** sheetId -> ("row,col" -> CellState). */
  readonly #sheets = new Map<SheetId, Map<CellKey, CellState>>();

  /** The cell map for a sheet, creating it on first write. */
  #sheet(sheetId: SheetId): Map<CellKey, CellState> {
    let sheet = this.#sheets.get(sheetId);
    if (sheet === undefined) {
      sheet = new Map<CellKey, CellState>();
      this.#sheets.set(sheetId, sheet);
    }
    return sheet;
  }

  /**
   * Read the current Shadow State at an absolute coordinate.
   *
   * An unwritten (or cleared) cell reads as the canonical empty cell — the
   * Shadow State is dense by definition even though it is stored sparsely.
   */
  read(sheetId: SheetId, row: number, col: number): CellState {
    return this.#sheets.get(sheetId)?.get(cellKey(row, col)) ?? emptyCell();
  }

  /**
   * Diff a {@link ValueObservation}'s after-slab against the mirror.
   *
   * Walks every cell of the observed {@link Area} (the after-slab is laid out
   * rectangle-by-rectangle, row-major within each), reading the "before" from
   * the Shadow State and comparing to the "after". Returns ONLY the cells that
   * actually changed — a no-op observation yields an empty list, which the
   * engine treats as "no Step".
   */
  diff(obs: ValueObservation): ChangedCell[] {
    const sheet = this.#sheets.get(obs.sheetId);
    const changed: ChangedCell[] = [];
    let slabRow = 0;

    for (const rect of obs.area) {
      for (let r = 0; r < rect.rowCount; r++) {
        const valuesRow = rowAt(obs.after.values, slabRow + r);
        const formulasRow = rowAt(obs.after.formulas, slabRow + r);
        const numberFormatsRow = rowAt(obs.after.numberFormats, slabRow + r);
        const valueTypesRow = rowAt(obs.after.valueTypes, slabRow + r);

        for (let c = 0; c < rect.colCount; c++) {
          const absRow = rect.startRow + r;
          const absCol = rect.startCol + c;
          const after: CellState = {
            value: cellAt(valuesRow, c, ''),
            formula: cellAt(formulasRow, c, null),
            valueType: cellAt<ValueType>(valueTypesRow, c, 'empty'),
            numberFormat: cellAt(numberFormatsRow, c, 'General'),
          };
          const before = sheet?.get(cellKey(absRow, absCol)) ?? emptyCell();
          if (!cellStateEquals(before, after)) {
            changed.push({ addr: cellRect(absRow, absCol), before, after });
          }
        }
      }
      slabRow += rect.rowCount;
    }

    return changed;
  }

  /**
   * Apply a {@link ValueDelta} forward into the mirror.
   *
   * Each cell's `after` becomes the new Shadow State at its coordinate. A cell
   * that returns to the empty state is removed from the sparse store so the
   * mirror does not accumulate empties. Navigation never inverts a delta
   * (ADR/spec Q6): this is forward-only.
   */
  apply(delta: ValueDelta): void {
    const sheet = this.#sheet(delta.sheetId);
    for (const cell of delta.cells) {
      const key = cellKey(cell.addr.startRow, cell.addr.startCol);
      if (cellStateEquals(cell.after, EMPTY_CELL)) {
        sheet.delete(key);
      } else {
        sheet.set(key, { ...cell.after });
      }
    }
  }

  /**
   * Build a {@link CellSlab} of the current mirror over an {@link Area}.
   *
   * Used to materialize "currently-projected" state for reconcile plans and
   * tests. Rectangle-by-rectangle, row-major, matching Observation slab layout.
   */
  slab(sheetId: SheetId, area: Area): CellSlab {
    const values: CellSlab['values'] = [];
    const formulas: CellSlab['formulas'] = [];
    const numberFormats: CellSlab['numberFormats'] = [];
    const valueTypes: CellSlab['valueTypes'] = [];

    for (const rect of area) {
      for (let r = 0; r < rect.rowCount; r++) {
        const valuesRow: CellSlab['values'][number] = [];
        const formulasRow: CellSlab['formulas'][number] = [];
        const numberFormatsRow: CellSlab['numberFormats'][number] = [];
        const valueTypesRow: CellSlab['valueTypes'][number] = [];
        for (let c = 0; c < rect.colCount; c++) {
          const state = this.read(sheetId, rect.startRow + r, rect.startCol + c);
          valuesRow.push(state.value);
          formulasRow.push(state.formula);
          numberFormatsRow.push(state.numberFormat);
          valueTypesRow.push(state.valueType);
        }
        values.push(valuesRow);
        formulas.push(formulasRow);
        numberFormats.push(numberFormatsRow);
        valueTypes.push(valueTypesRow);
      }
    }

    return { values, formulas, numberFormats, valueTypes };
  }

  /** Number of non-empty cells held for a sheet (0 if the sheet is unknown). */
  cellCount(sheetId: SheetId): number {
    return this.#sheets.get(sheetId)?.size ?? 0;
  }
}

/** Safe row access into a slab (noUncheckedIndexedAccess): missing -> `[]`. */
function rowAt<T>(grid: T[][], index: number): T[] {
  return grid[index] ?? [];
}

/** Safe cell access into a slab row, defaulting a short/ragged row. */
function cellAt<T>(row: readonly T[], index: number, fallback: T): T {
  // Only a genuinely-absent (ragged-row) cell falls back; a present-but-falsy
  // value (0, '', false, null) is preserved, so this is `=== undefined`, not
  // `??` — `noUncheckedIndexedAccess` makes the indexed access `T | undefined`.
  const cell = row[index];
  // eslint-disable-next-line @typescript-eslint/prefer-nullish-coalescing
  return cell === undefined ? fallback : cell;
}
