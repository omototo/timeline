/**
 * Pure translation between Office.js enum strings and the engine-neutral
 * vocabulary in `@timeline/engine`. No Office.js type crosses the seam — these
 * helpers are the only place the two namings meet.
 */
import {
  PREVIEW_SHEET_PREFIX,
  type ShiftDirection,
  type StructuralChangeType,
  type ValueType,
  type CellValue,
  type CellSlab,
} from '@timeline/engine';
import type {
  ChangeDirectionStateLike,
  ExcelDataChangeType,
  ExcelDeleteShiftDirection,
  ExcelInsertShiftDirection,
  RangeLike,
} from './office-types.ts';

/**
 * Map an engine sheet id to a valid Excel worksheet NAME.
 *
 * Real sheet ids pass through (Office `getItem` accepts an id or a name). Engine
 * preview-surface ids carry the `__preview__::<id>` prefix, whose `:` is illegal
 * in an Excel sheet name and whose length (a 38-char worksheet GUID) blows the
 * 31-char limit — so they fold to a short, legal, deterministic name. The fold
 * is stable per logical sheet, so create/activate/delete/getItem all resolve the
 * same Excel sheet.
 */
/** Name prefix marking an engine-owned preview surface (never a user worksheet). */
export const INTERNAL_SHEET_PREFIX = '__tl_preview_';

export function toExcelSheetName(sheetId: string): string {
  if (!sheetId.startsWith(PREVIEW_SHEET_PREFIX)) {
    return sheetId;
  }
  const logical = sheetId.slice(PREVIEW_SHEET_PREFIX.length);
  let hash = 0x811c9dc5;
  for (let i = 0; i < logical.length; i += 1) {
    hash ^= logical.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  // prefix (13) + 8 hex = 21 chars: under Excel's 31-char limit, all legal.
  return `${INTERNAL_SHEET_PREFIX}${(hash >>> 0).toString(16).padStart(8, '0')}`;
}

/**
 * Is this an engine-owned preview surface (not a user worksheet)? The change
 * source uses this to ignore the add/delete of the add-in's own preview sheets,
 * which would otherwise pollute the timeline with phantom "worksheet change"
 * Steps every time the user previews.
 */
export function isInternalSheetName(name: string): boolean {
  return name.startsWith(INTERNAL_SHEET_PREFIX);
}

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

/**
 * Excel grid dimensions, used as the spanning extent for whole-row / whole-column
 * references (e.g. `"3:3"` spans all 16,384 columns; `"C:C"` spans all 1,048,576
 * rows). Real Excel emits these bare row/column forms on structural change events,
 * so `parseAddress` must accept them without throwing (it runs inside an awaited
 * `onChanged` handler). We represent the unbounded dimension as a sentinel extent
 * rather than adding a `wholeRow`/`wholeColumn` flag, keeping the `ParsedRect`
 * shape (and the engine `Rect` it maps to) unchanged — the least-invasive option.
 */
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;

const FULL_CELL_RE = /^([A-Z]+)(\d+)$/;
const COL_ONLY_RE = /^([A-Z]+)$/;
const ROW_ONLY_RE = /^(\d+)$/;

/** Converts a column letter run (A, B, …, AA) to a zero-based index. */
function colToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

/** One endpoint of an address: a full cell, a bare column, or a bare row. */
type Endpoint =
  | { kind: 'cell'; row: number; col: number }
  | { kind: 'col'; col: number }
  | { kind: 'row'; row: number };

/** Strip `$` anchors and any leading `Sheet!` qualifier from one endpoint token. */
function normalizeToken(token: string): string {
  return token.replace(/\$/g, '');
}

function parseEndpoint(token: string, address: string): Endpoint {
  const full = FULL_CELL_RE.exec(token);
  if (full?.[1] !== undefined && full[2] !== undefined) {
    return { kind: 'cell', row: Number(full[2]) - 1, col: colToIndex(full[1]) };
  }
  const colOnly = COL_ONLY_RE.exec(token);
  if (colOnly?.[1] !== undefined) {
    return { kind: 'col', col: colToIndex(colOnly[1]) };
  }
  const rowOnly = ROW_ONLY_RE.exec(token);
  if (rowOnly?.[1] !== undefined) {
    return { kind: 'row', row: Number(rowOnly[1]) - 1 };
  }
  throw new Error(`parseAddress: cannot parse address "${address}".`);
}

/**
 * Parses an A1-style address into a zero-based rectangle. Robust to the forms
 * real Excel emits on structural change events: single cells (`"B2"`), ranges
 * (`"B2:D5"`), absolute references (`"$B$2"`, `"$B$2:$D$4"`), sheet qualifiers
 * (`"Sheet1!A1"`), whole-row ranges (`"3:3"`, `"5:7"` → span all columns) and
 * whole-column ranges (`"C:C"`, `"B:D"` → span all rows). Single cells yield a
 * 1×1 rect. Must NOT throw on these structural forms — only on genuinely
 * unparseable input.
 */
export function parseAddress(address: string): ParsedRect {
  const bang = address.lastIndexOf('!');
  const local = bang >= 0 ? address.slice(bang + 1) : address;
  const [rawStart, rawEnd, extra] = local.split(':');
  if (extra !== undefined || rawStart === undefined) {
    throw new Error(`parseAddress: cannot parse address "${address}".`);
  }
  const from = parseEndpoint(normalizeToken(rawStart), address);

  if (rawEnd === undefined) {
    // A bare column ("C") or bare row ("3") with no `:` is still a full span.
    if (from.kind === 'cell') {
      return { startRow: from.row, startCol: from.col, rowCount: 1, colCount: 1 };
    }
    return spanFromEndpoints(from, from, address);
  }

  const to = parseEndpoint(normalizeToken(rawEnd), address);
  return spanFromEndpoints(from, to, address);
}

/** Combine two endpoints into a rect, spanning the unbounded dimension for whole-row/column forms. */
function spanFromEndpoints(from: Endpoint, to: Endpoint, address: string): ParsedRect {
  // Whole-column range: "C:C" / "B:D" — both endpoints are columns, span all rows.
  if (from.kind === 'col' && to.kind === 'col') {
    return {
      startRow: 0,
      startCol: Math.min(from.col, to.col),
      rowCount: MAX_ROWS,
      colCount: Math.abs(to.col - from.col) + 1,
    };
  }
  // Whole-row range: "3:3" / "5:7" — both endpoints are rows, span all columns.
  if (from.kind === 'row' && to.kind === 'row') {
    return {
      startRow: Math.min(from.row, to.row),
      startCol: 0,
      rowCount: Math.abs(to.row - from.row) + 1,
      colCount: MAX_COLS,
    };
  }
  // Otherwise both endpoints must be full cells (mixing a cell with a bare
  // row/column, e.g. "A1:zzz" or "A1:3", is not an address Excel emits).
  if (from.kind !== 'cell' || to.kind !== 'cell') {
    throw new Error(`parseAddress: cannot parse address "${address}".`);
  }
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
