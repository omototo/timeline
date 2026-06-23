/**
 * Minimal structural typings for the slice of the Office.js / Excel API surface
 * the Stream B adapters actually touch.
 *
 * Following the slice-1 pattern (`workbook-stamp.ts`), the adapters depend on
 * these injectable `*Like` interfaces rather than the ambient `Excel` global.
 * That keeps the adapters headlessly testable — a fake (`__mocks__/excel.ts`)
 * implements the same shapes — without requiring a real Office host or pulling
 * Office.js types across the engine seam.
 *
 * These mirror the real Office.js names/values closely enough to swap for the
 * genuine `Excel.*` types at the call site, but only the members in use are
 * declared. No Office.js type is re-exported from here; engine-neutral
 * `Observation` / `ReconcilePlan` values are produced by the adapters that
 * consume these.
 */

/** `Excel.DataChangeType` member strings (ExcelApi 1.7). */
export type ExcelDataChangeType =
  | 'RangeEdited'
  | 'RowInserted'
  | 'RowDeleted'
  | 'ColumnInserted'
  | 'ColumnDeleted'
  | 'CellInserted'
  | 'CellDeleted'
  | 'Unknown';

/** `Excel.EventSource` (Local/Remote) — ExcelApi 1.8. */
export type ExcelEventSource = 'Local' | 'Remote';

/** `WorksheetChangedEventArgs.triggerSource` — ExcelApi 1.14. */
export type ExcelTriggerSource = 'None' | 'ThisLocalAddin' | 'Unknown';

/** `Excel.InsertShiftDirection` (ExcelApi 1.1). */
export type ExcelInsertShiftDirection = 'Down' | 'Right';

/** `Excel.DeleteShiftDirection` (ExcelApi 1.1). */
export type ExcelDeleteShiftDirection = 'Up' | 'Left';

/** `Excel.SheetVisibility` (ExcelApi 1.1; `veryHidden` member usable from 1.1). */
export type ExcelSheetVisibility = 'Visible' | 'Hidden' | 'VeryHidden';

/** `changeDirectionState` (ExcelApi 1.14): the two members are mutually exclusive. */
export interface ChangeDirectionStateLike {
  insertShiftDirection?: ExcelInsertShiftDirection;
  deleteShiftDirection?: ExcelDeleteShiftDirection;
}

/** Single-cell `WorksheetChangedEventArgs.details` (ExcelApi 1.9). */
export interface ChangedEventDetailLike {
  valueAfter: unknown;
  valueBefore: unknown;
  valueTypeAfter: string;
  valueTypeBefore: string;
}

/**
 * The slice of `Excel.WorksheetChangedEventArgs` we read. `triggerSource` and
 * `changeDirectionState` are 1.14 and feature-detected, so optional here.
 */
export interface WorksheetChangedEventArgsLike {
  worksheetId: string;
  /** A1-style address of the changed range, scoped to the sheet (e.g. "B2:C3"). */
  address: string;
  changeType: ExcelDataChangeType;
  source: ExcelEventSource;
  triggerSource?: ExcelTriggerSource;
  changeDirectionState?: ChangeDirectionStateLike;
  details?: ChangedEventDetailLike | null;
}

/** A registered event handler we can later remove. */
export interface EventHandlerResultLike {
  remove(): void;
}

/** `Excel.WorksheetCollection`-level add/delete/name-change args. */
export interface WorksheetAddedEventArgsLike {
  worksheetId: string;
  source: ExcelEventSource;
}
export interface WorksheetDeletedEventArgsLike {
  worksheetId: string;
  source: ExcelEventSource;
}
export interface WorksheetNameChangedEventArgsLike {
  worksheetId: string;
  /** New worksheet name (ExcelApi 1.7+). */
  nameAfter?: string;
  source: ExcelEventSource;
}
export interface WorksheetPositionChangedEventArgsLike {
  worksheetId: string;
  positionAfter?: number;
  source: ExcelEventSource;
}

/** A toggleable event with `.add` returning a removable handle. */
export interface EventToggleLike<TArgs> {
  add(handler: (args: TArgs) => void | Promise<void>): EventHandlerResultLike;
}

/** The slice of an `Excel.Range` proxy the adapters read or write. */
export interface RangeLike {
  /** Loadable: number of cells in the range (ExcelApi 1.4). */
  cellCount: number;
  /** Loadable 2-D arrays (`[row][col]`). */
  values: unknown[][];
  formulas: unknown[][];
  numberFormat: unknown[][];
  valueTypes: string[][];
  /** Stage a load of the named properties before the next `sync()`. */
  load(properties: string[] | string): void;
  /** Release the proxy from tracked memory (ExcelApi 1.3). */
  untrack(): RangeLike;
  /** Sub-range by zero-based offset (used to tile a large read-back). */
  getCell(row: number, column: number): RangeLike;
  /** A sub-block of this range, zero-based within the range. */
  getOffsetRange?(rowOffset: number, columnOffset: number): RangeLike;
  getBoundingRect?(otherRange: RangeLike): RangeLike;
  /** Structural ops (ExcelApi 1.1). */
  insert(shift: ExcelInsertShiftDirection): void;
  delete(shift: ExcelDeleteShiftDirection): void;
}

/** The slice of an `Excel.Worksheet` proxy the adapters use. */
export interface WorksheetLike {
  id: string;
  name: string;
  position: number;
  visibility: ExcelSheetVisibility;
  onChanged: EventToggleLike<WorksheetChangedEventArgsLike>;
  onFormatChanged: EventToggleLike<WorksheetChangedEventArgsLike>;
  getRange(address?: string): RangeLike;
  getRangeByIndexes(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
  ): RangeLike;
  /** Make this the active worksheet (switches the user's view to it). */
  activate(): void;
  delete(): void;
}

/** The slice of `Excel.WorksheetCollection`. */
export interface WorksheetCollectionLike {
  onAdded: EventToggleLike<WorksheetAddedEventArgsLike>;
  onDeleted: EventToggleLike<WorksheetDeletedEventArgsLike>;
  onNameChanged: EventToggleLike<WorksheetNameChangedEventArgsLike>;
  onMoved?: EventToggleLike<WorksheetPositionChangedEventArgsLike>;
  add(name?: string): WorksheetLike;
  getItem(key: string): WorksheetLike;
  getItemOrNullObject(key: string): WorksheetLike & { isNullObject?: boolean };
}

/** The slice of `Excel.Workbook`. */
export interface WorkbookLike {
  worksheets: WorksheetCollectionLike;
}

/** The slice of `Excel.RequestContext` passed to an `Excel.run` body. */
export interface RequestContextLike {
  workbook: WorkbookLike;
  sync(): Promise<void>;
}

/**
 * The `Excel.run`-shaped entry point: runs a batch against a fresh context.
 * Injecting this (instead of the `Excel` global) is what makes the adapters
 * headless.
 */
export type ExcelRun = <T>(batch: (context: RequestContextLike) => Promise<T>) => Promise<T>;

/**
 * Feature-detection shim mirroring
 * `Office.context.requirements.isSetSupported('ExcelApi', '1.14')`.
 */
export type IsSetSupported = (name: string, minVersion?: string) => boolean;
