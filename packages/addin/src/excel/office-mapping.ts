/**
 * Pure translation between Office.js enum strings and the engine-neutral
 * vocabulary in `@timeline/engine`. No Office.js type crosses the seam — these
 * helpers are the only place the two namings meet.
 */
import type {
  ShiftDirection,
  StructuralChangeType,
  ValueType,
  CellValue,
  CellSlab,
} from '@timeline/engine';
import type {
  ChangeDirectionStateLike,
  ExcelDataChangeType,
  ExcelDeleteShiftDirection,
  ExcelInsertShiftDirection,
  RangeLike,
} from './office-types.ts';

/** A1-style address parsed into a zero-based rectangle. */
export interface ParsedRect {
  startRow: number;
  startCol: number;
  rowCount: number;
  colCount: number;
}

/** Maps `Excel.DataChangeType` to the engine's structural change type, or null for a value edit. */
export function toStructuralChangeType(
  changeType: ExcelDataChangeType,
): StructuralChangeType | null {
  switch (changeType) {
    case 'RowInserted':
      return 'rowInserted';
    case 'RowDeleted':
      return 'rowDeleted';
    case 'ColumnInserted':
      return 'columnInserted';
    case 'ColumnDeleted':
      return 'columnDeleted';
    case 'CellInserted':
      return 'cellInserted';
    case 'CellDeleted':
      return 'cellDeleted';
    case 'RangeEdited':
    case 'Unknown':
      return null;
  }
}

/** Maps the 1.14 `changeDirectionState` to the engine's shift direction, if present. */
export function toShiftDirection(
  state: ChangeDirectionStateLike | undefined,
): ShiftDirection | undefined {
  if (state === undefined) {
    return undefined;
  }
  if (state.insertShiftDirection !== undefined) {
    return state.insertShiftDirection === 'Down' ? 'down' : 'right';
  }
  if (state.deleteShiftDirection !== undefined) {
    return state.deleteShiftDirection === 'Up' ? 'up' : 'left';
  }
  return undefined;
}

/** Maps engine shift direction back to the Office insert enum (for reconcile writes). */
export function toInsertShift(direction: ShiftDirection | undefined): ExcelInsertShiftDirection {
  // Row/column inserts shift down/right; default to 'Down' for a row insert.
  return direction === 'right' ? 'Right' : 'Down';
}

/** Maps engine shift direction back to the Office delete enum (for reconcile writes). */
export function toDeleteShift(direction: ShiftDirection | undefined): ExcelDeleteShiftDirection {
  return direction === 'left' ? 'Left' : 'Up';
}

/** Maps an Excel `valueTypes` string to the engine `ValueType`. */
export function toValueType(raw: string): ValueType {
  switch (raw) {
    case 'Empty':
      return 'empty';
    case 'String':
      return 'string';
    case 'Double':
    case 'Integer':
      return 'number';
    case 'Boolean':
      return 'boolean';
    case 'Error':
      return 'error';
    default:
      // RichValue, Entity, LinkedEntity, etc. — flattened (Fidelity Caveat).
      return 'richValue';
  }
}

const COL_RE = /^([A-Z]+)(\d+)$/;

/** Converts a column letter run (A, B, …, AA) to a zero-based index. */
function colToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

/**
 * Parses an A1-style address (optionally sheet-qualified, optionally a range)
 * into a zero-based rectangle. Single cells yield a 1×1 rect.
 *
 * Examples: `"B2"`, `"B2:D5"`, `"Sheet1!B2:D5"`.
 */
function parseCell(token: string | undefined, address: string): { row: number; col: number } {
  const match = COL_RE.exec(token ?? '');
  const letters = match?.[1];
  const digits = match?.[2];
  if (letters === undefined || digits === undefined) {
    throw new Error(`parseAddress: cannot parse address "${address}".`);
  }
  return { row: Number(digits) - 1, col: colToIndex(letters) };
}

export function parseAddress(address: string): ParsedRect {
  const bang = address.lastIndexOf('!');
  const local = bang >= 0 ? address.slice(bang + 1) : address;
  const [start, end] = local.split(':');
  const from = parseCell(start, address);
  if (end === undefined) {
    return { startRow: from.row, startCol: from.col, rowCount: 1, colCount: 1 };
  }
  const to = parseCell(end, address);
  return {
    startRow: Math.min(from.row, to.row),
    startCol: Math.min(from.col, to.col),
    rowCount: Math.abs(to.row - from.row) + 1,
    colCount: Math.abs(to.col - from.col) + 1,
  };
}

/**
 * Builds an engine `CellSlab` from a loaded `RangeLike`'s 2-D arrays. The range
 * must already have been `load()`ed and `sync()`ed by the caller.
 */
export function slabFromRange(range: RangeLike): CellSlab {
  const values: CellValue[][] = range.values;
  const rawFormulas = range.formulas;
  const rawNumberFormats = range.numberFormat;
  const rawValueTypes = range.valueTypes;
  const formulas: (string | null)[][] = rawFormulas.map((row) =>
    row.map((f) => (typeof f === 'string' && f.startsWith('=') ? f : null)),
  );
  const numberFormats: string[][] = rawNumberFormats.map((row) =>
    row.map((n) => (typeof n === 'string' ? n : 'General')),
  );
  const valueTypes: ValueType[][] = rawValueTypes.map((row) => row.map((t) => toValueType(t)));
  return { values, formulas, numberFormats, valueTypes };
}
