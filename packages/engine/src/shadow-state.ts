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
  StructuralDelta,
  ValueDelta,
  ValueObservation,
  ValueType,
  WorksheetDelta,
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

/**
 * Per-sheet metadata mirrored alongside the cell map: the sheet's display
 * `name` and its `order` (left-to-right tab position). Tracked so the engine
 * can replay {@link WorksheetDelta}s (add/delete/rename/reorder — ADR-0005)
 * deterministically. The `sheetId` is the stable key (a sheet keeps its id
 * across a rename); `name` is the human-facing label.
 */
export interface SheetMeta {
  sheetId: SheetId;
  name: string;
  order: number;
}

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
  /** sheetId -> sheet metadata (name + tab order), for Worksheet Deltas. */
  readonly #sheetMeta = new Map<SheetId, SheetMeta>();

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
   * Apply a {@link StructuralDelta} forward into the mirror as a COORDINATE
   * REMAP (ADR-0001). A structural op is a coordinate transform, not a value
   * change: an insert opens blank space and shifts existing cells down/right;
   * a delete removes the spanned cells and shifts the rest up/left.
   *
   * Cell formulas are opaque strings — this NEVER rewrites formula-text
   * references (that is Excel's job — ADR-0003); it only moves whole
   * {@link CellState}s to new coordinates. Deterministic (needed for replay).
   *
   * Row/column ops span the full sheet on the orthogonal axis. Cell ops are
   * bounded to the address rectangle's rows (for a left/right shift) or columns
   * (for an up/down shift); the `shiftDirection` disambiguates.
   */
  applyStructural(delta: StructuralDelta): void {
    const sheet = this.#sheets.get(delta.sheetId);
    if (sheet === undefined) return; // nothing to move on an untouched sheet

    const remapped = new Map<CellKey, CellState>();
    for (const [key, st] of sheet) {
      const parsed = parseCellKey(key);
      const moved = remapCoordinate(delta, parsed.row, parsed.col);
      // `null` => the cell was deleted (it lived in the removed span).
      if (moved !== null) {
        remapped.set(cellKey(moved.row, moved.col), st);
      }
    }
    this.#sheets.set(delta.sheetId, remapped);
  }

  /**
   * Apply a {@link WorksheetDelta} forward into the sheet-metadata map
   * (ADR-0005). `add` registers a new sheet; `delete` drops it and its cells;
   * `rename` updates the display name (id is stable); `reorder` moves the tab
   * to `newPosition`, re-packing the surrounding `order` values. Deterministic.
   */
  applyWorksheet(delta: WorksheetDelta): void {
    switch (delta.op) {
      case 'add': {
        // Append at the end first, then (if a position was given) move it into
        // place so an explicit `newPosition` actually re-packs the neighbours.
        this.#sheetMeta.set(delta.sheetId, {
          sheetId: delta.sheetId,
          name: delta.newName ?? delta.sheetId,
          order: this.#sheetMeta.size,
        });
        this.#repackOrder();
        if (delta.newPosition !== undefined) {
          this.#reorderSheet(delta.sheetId, delta.newPosition);
        }
        return;
      }
      case 'delete': {
        this.#sheetMeta.delete(delta.sheetId);
        this.#sheets.delete(delta.sheetId);
        this.#repackOrder();
        return;
      }
      case 'rename': {
        const existing = this.#sheetMeta.get(delta.sheetId);
        const name = delta.newName ?? existing?.name ?? delta.sheetId;
        const order = existing?.order ?? this.#sheetMeta.size;
        this.#sheetMeta.set(delta.sheetId, { sheetId: delta.sheetId, name, order });
        return;
      }
      case 'reorder': {
        const existing = this.#sheetMeta.get(delta.sheetId);
        if (existing === undefined || delta.newPosition === undefined) return;
        this.#reorderSheet(delta.sheetId, delta.newPosition);
        return;
      }
    }
  }

  /** Sheet metadata for a sheet, or `undefined` if untracked. */
  sheetMeta(sheetId: SheetId): SheetMeta | undefined {
    const meta = this.#sheetMeta.get(sheetId);
    return meta === undefined ? undefined : { ...meta };
  }

  /** All tracked sheets in tab order (defensive copies). */
  sheets(): SheetMeta[] {
    return [...this.#sheetMeta.values()].sort((a, b) => a.order - b.order).map((m) => ({ ...m }));
  }

  /**
   * Move `sheetId` to `newPosition` and re-pack every sheet's `order` to a
   * dense 0-based sequence in the resulting tab order.
   */
  #reorderSheet(sheetId: SheetId, newPosition: number): void {
    const ordered = [...this.#sheetMeta.values()].sort((a, b) => a.order - b.order);
    const fromIndex = ordered.findIndex((m) => m.sheetId === sheetId);
    if (fromIndex === -1) return;
    const [moved] = ordered.splice(fromIndex, 1);
    if (moved === undefined) return;
    const clamped = Math.max(0, Math.min(newPosition, ordered.length));
    ordered.splice(clamped, 0, moved);
    ordered.forEach((m, i) => {
      this.#sheetMeta.set(m.sheetId, { ...m, order: i });
    });
  }

  /** Re-pack `order` to a dense 0-based sequence preserving relative order. */
  #repackOrder(): void {
    const ordered = [...this.#sheetMeta.values()].sort((a, b) => a.order - b.order);
    ordered.forEach((m, i) => {
      this.#sheetMeta.set(m.sheetId, { ...m, order: i });
    });
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

/** Parse a `"row,col"` {@link CellKey} back into numeric coordinates. */
function parseCellKey(key: CellKey): { row: number; col: number } {
  const comma = key.indexOf(',');
  return {
    row: Number(key.slice(0, comma)),
    col: Number(key.slice(comma + 1)),
  };
}

/**
 * Remap a single cell coordinate under a {@link StructuralDelta}.
 *
 * Returns the new `(row, col)`, or `null` if the cell was deleted (it lived in
 * the removed span). Row/column ops are sheet-wide on the orthogonal axis; cell
 * ops are bounded to the address rectangle on the orthogonal axis and shift in
 * the `shiftDirection`.
 */
function remapCoordinate(
  delta: StructuralDelta,
  row: number,
  col: number,
): { row: number; col: number } | null {
  const { address } = delta;
  switch (delta.changeType) {
    case 'rowInserted':
      // Open `rowCount` blank rows at `startRow`: rows at/after shift down.
      return row >= address.startRow ? { row: row + address.rowCount, col } : { row, col };
    case 'rowDeleted': {
      const end = address.startRow + address.rowCount;
      if (row >= address.startRow && row < end) return null; // in removed span
      return row >= end ? { row: row - address.rowCount, col } : { row, col };
    }
    case 'columnInserted':
      return col >= address.startCol ? { row, col: col + address.colCount } : { row, col };
    case 'columnDeleted': {
      const end = address.startCol + address.colCount;
      if (col >= address.startCol && col < end) return null;
      return col >= end ? { row, col: col - address.colCount } : { row, col };
    }
    case 'cellInserted':
      return remapCellInserted(delta, row, col);
    case 'cellDeleted':
      return remapCellDeleted(delta, row, col);
  }
}

/** Does a coordinate fall within the address rectangle's rows? */
function inRows(address: Rect, row: number): boolean {
  return row >= address.startRow && row < address.startRow + address.rowCount;
}

/** Does a coordinate fall within the address rectangle's columns? */
function inCols(address: Rect, col: number): boolean {
  return col >= address.startCol && col < address.startCol + address.colCount;
}

/** Insert a block of cells, shifting `down` (default) or `right`. */
function remapCellInserted(
  delta: StructuralDelta,
  row: number,
  col: number,
): { row: number; col: number } {
  const { address } = delta;
  if (delta.shiftDirection === 'right') {
    // Within the affected rows, cells at/after startCol shift right.
    return inRows(address, row) && col >= address.startCol
      ? { row, col: col + address.colCount }
      : { row, col };
  }
  // Default 'down': within the affected columns, cells at/after startRow shift down.
  return inCols(address, col) && row >= address.startRow
    ? { row: row + address.rowCount, col }
    : { row, col };
}

/** Delete a block of cells, shifting `up` (default) or `left`. */
function remapCellDeleted(
  delta: StructuralDelta,
  row: number,
  col: number,
): { row: number; col: number } | null {
  const { address } = delta;
  if (delta.shiftDirection === 'left') {
    // Within the affected rows, the spanned columns are removed and cells to
    // the right shift left; rows outside the address are untouched.
    if (!inRows(address, row)) return { row, col };
    const end = address.startCol + address.colCount;
    if (col >= address.startCol && col < end) return null; // removed
    return col >= end ? { row, col: col - address.colCount } : { row, col };
  }
  // Default 'up': within the affected columns, the spanned rows are removed and
  // cells below shift up.
  if (!inCols(address, col)) return { row, col };
  const end = address.startRow + address.rowCount;
  if (row >= address.startRow && row < end) return null; // removed
  return row >= end ? { row: row - address.rowCount, col } : { row, col };
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
