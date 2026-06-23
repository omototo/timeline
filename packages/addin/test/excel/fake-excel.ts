/**
 * A minimal headless fake of the Office.js Excel surface the Stream B adapters
 * use (`office-types.ts`). Enough to unit-test `OfficeChangeSource` and the
 * render targets without a real Excel host.
 *
 * The fake models a tiny workbook of sheets, each with a sparse cell grid, and
 * synthetic event hubs you can fire from a test (`fireChanged`, `fireAdded`,
 * etc.). `run(batch)` invokes the batch against a shared context; `sync()` is a
 * no-op resolve (mutations apply to the in-memory grid eagerly).
 */
import type {
  ChangeDirectionStateLike,
  ChangedEventDetailLike,
  EventHandlerResultLike,
  EventToggleLike,
  ExcelDataChangeType,
  ExcelEventSource,
  ExcelRun,
  ExcelSheetVisibility,
  ExcelTriggerSource,
  RangeLike,
  RequestContextLike,
  WorkbookLike,
  WorksheetAddedEventArgsLike,
  WorksheetChangedEventArgsLike,
  WorksheetCollectionLike,
  WorksheetDeletedEventArgsLike,
  WorksheetLike,
  WorksheetNameChangedEventArgsLike,
  WorksheetPositionChangedEventArgsLike,
} from '../../src/excel/office-types.ts';
import { parseAddress } from '../../src/excel/office-mapping.ts';

/** One cell's stored state in the fake grid. */
export interface FakeCell {
  value?: unknown;
  formula?: string;
  numberFormat?: string;
  valueType?: string;
}

/** A simple synchronous event hub. */
class EventHub<TArgs> implements EventToggleLike<TArgs> {
  readonly #handlers = new Set<(args: TArgs) => void | Promise<void>>();

  add(handler: (args: TArgs) => void | Promise<void>): EventHandlerResultLike {
    this.#handlers.add(handler);
    return {
      remove: () => {
        this.#handlers.delete(handler);
      },
    };
  }

  /** Fire all handlers; awaits any returned promises so async read-back settles. */
  async fire(args: TArgs): Promise<void> {
    for (const handler of [...this.#handlers]) {
      await handler(args);
    }
  }

  get handlerCount(): number {
    return this.#handlers.size;
  }
}

/**
 * A fake `Excel.Range`: a view onto a rectangular block of a sheet's grid.
 *
 * Reads (`load` + getter) project the grid into 2-D arrays. Writes go through
 * setters that flush back into the grid eagerly (mirroring the real proxy-queue
 * model closely enough for the adapter tests, where `sync()` is a no-op).
 */
class FakeRange implements RangeLike {
  cellCount = 0;
  valueTypes: string[][] = [];
  #values: unknown[][] = [];
  #formulas: unknown[][] = [];
  #numberFormat: unknown[][] = [];

  constructor(
    private readonly sheet: FakeWorksheet,
    private readonly startRow: number,
    private readonly startCol: number,
    private readonly rowCount: number,
    private readonly colCount: number,
  ) {}

  get values(): unknown[][] {
    return this.#values;
  }
  set values(grid: unknown[][]) {
    this.#values = grid;
    this.#flushValues(grid);
  }

  get formulas(): unknown[][] {
    return this.#formulas;
  }
  set formulas(grid: unknown[][]) {
    this.#formulas = grid;
    this.#flushFormulas(grid);
  }

  get numberFormat(): unknown[][] {
    return this.#numberFormat;
  }
  set numberFormat(grid: unknown[][]) {
    this.#numberFormat = grid;
    this.#flushNumberFormat(grid);
  }

  load(properties: string[] | string): void {
    const props = Array.isArray(properties) ? properties : [properties];
    if (props.includes('cellCount')) {
      // The host normally reports rowCount * colCount, but a test can install a
      // hook to make a probe over-report (modelling Excel returning a larger
      // actual count than the adapter predicted) to drive the defensive
      // sub-tiling branch.
      this.cellCount =
        this.sheet.cellCountFor?.(this.startRow, this.startCol, this.rowCount, this.colCount) ??
        this.rowCount * this.colCount;
    }
    if (props.includes('values')) {
      this.#values = this.#read((c) => c.value ?? '');
    }
    if (props.includes('formulas')) {
      this.#formulas = this.#read((c) => c.formula ?? c.value ?? '');
    }
    if (props.includes('numberFormat')) {
      this.#numberFormat = this.#read((c) => c.numberFormat ?? 'General');
    }
    if (props.includes('valueTypes')) {
      this.valueTypes = this.#read((c) => c.valueType ?? 'String') as string[][];
    }
  }

  #read(pick: (cell: FakeCell) => unknown): unknown[][] {
    const out: unknown[][] = [];
    for (let r = 0; r < this.rowCount; r++) {
      const row: unknown[] = [];
      for (let c = 0; c < this.colCount; c++) {
        row.push(pick(this.sheet.cellAt(this.startRow + r, this.startCol + c)));
      }
      out.push(row);
    }
    return out;
  }

  #cell(r: number, c: number): FakeCell {
    return this.sheet.cellAt(this.startRow + r, this.startCol + c);
  }

  #flushValues(grid: unknown[][]): void {
    grid.forEach((row, r) => {
      row.forEach((v, c) => {
        if (v !== undefined) {
          this.#cell(r, c).value = v;
        }
      });
    });
  }

  #flushFormulas(grid: unknown[][]): void {
    grid.forEach((row, r) => {
      row.forEach((f, c) => {
        const cell = this.#cell(r, c);
        if (typeof f === 'string' && f.startsWith('=')) {
          cell.formula = f;
        } else if (f !== undefined && f !== null) {
          // A non-formula entry in `.formulas` is a literal constant.
          cell.value = f;
        }
      });
    });
  }

  #flushNumberFormat(grid: unknown[][]): void {
    grid.forEach((row, r) => {
      row.forEach((n, c) => {
        if (typeof n === 'string') {
          this.#cell(r, c).numberFormat = n;
        }
      });
    });
  }

  untrack(): RangeLike {
    this.sheet.workbook.untrackCount++;
    return this;
  }

  getCell(row: number, column: number): RangeLike {
    return new FakeRange(this.sheet, this.startRow + row, this.startCol + column, 1, 1);
  }

  getOffsetRange(rowOffset: number, columnOffset: number): RangeLike {
    return new FakeRange(
      this.sheet,
      this.startRow + rowOffset,
      this.startCol + columnOffset,
      this.rowCount,
      this.colCount,
    );
  }

  getBoundingRect(): RangeLike {
    return this;
  }

  insert(shift: 'Down' | 'Right'): void {
    this.sheet.workbook.structuralOps.push({
      sheetId: this.sheet.id,
      kind: 'insert',
      shift,
      startRow: this.startRow,
      startCol: this.startCol,
    });
  }

  delete(shift: 'Up' | 'Left'): void {
    this.sheet.workbook.structuralOps.push({
      sheetId: this.sheet.id,
      kind: 'delete',
      shift,
      startRow: this.startRow,
      startCol: this.startCol,
    });
  }
}

export class FakeWorksheet implements WorksheetLike {
  visibility: ExcelSheetVisibility = 'Visible';
  readonly onChanged = new EventHub<WorksheetChangedEventArgsLike>();
  readonly onFormatChanged = new EventHub<WorksheetChangedEventArgsLike>();
  readonly #grid = new Map<string, FakeCell>();

  /**
   * Optional hook letting a test override the reported `cellCount` for a probed
   * rect — used to model the host returning a larger actual count than the
   * adapter predicted, exercising the defensive sub-tiling path.
   */
  cellCountFor?: (startRow: number, startCol: number, rowCount: number, colCount: number) => number;

  constructor(
    readonly workbook: FakeWorkbook,
    readonly id: string,
    public name: string,
    public position: number,
  ) {}

  cellAt(row: number, col: number): FakeCell {
    const key = `${String(row)},${String(col)}`;
    let cell = this.#grid.get(key);
    if (cell === undefined) {
      cell = {};
      this.#grid.set(key, cell);
    }
    return cell;
  }

  /** Seed a cell so a read-back returns meaningful values. */
  setCell(row: number, col: number, cell: FakeCell): void {
    this.#grid.set(`${String(row)},${String(col)}`, cell);
  }

  getRange(address?: string): RangeLike {
    const rect = parseA1(address ?? 'A1');
    return new FakeRange(this, rect.startRow, rect.startCol, rect.rowCount, rect.colCount);
  }

  getRangeByIndexes(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
  ): RangeLike {
    return new FakeRange(this, startRow, startColumn, rowCount, columnCount);
  }

  activate(): void {
    this.workbook.activeSheetId = this.id;
  }

  delete(): void {
    this.workbook.removeSheet(this.id);
  }
}

interface StructuralOpRecord {
  sheetId: string;
  kind: 'insert' | 'delete';
  shift: string;
  startRow: number;
  startCol: number;
}

class FakeWorksheetCollection implements WorksheetCollectionLike {
  readonly onAdded = new EventHub<WorksheetAddedEventArgsLike>();
  readonly onDeleted = new EventHub<WorksheetDeletedEventArgsLike>();
  readonly onNameChanged = new EventHub<WorksheetNameChangedEventArgsLike>();
  readonly onMoved = new EventHub<WorksheetPositionChangedEventArgsLike>();

  constructor(private readonly workbook: FakeWorkbook) {}

  add(name?: string): WorksheetLike {
    return this.workbook.addSheet(name);
  }

  getItem(key: string): WorksheetLike {
    const sheet = this.workbook.findSheet(key);
    if (sheet === undefined) {
      throw new Error(`FakeWorksheetCollection.getItem: no sheet "${key}".`);
    }
    return sheet;
  }

  getItemOrNullObject(key: string): WorksheetLike & { isNullObject?: boolean } {
    const sheet = this.workbook.findSheet(key);
    if (sheet === undefined) {
      return nullSheet();
    }
    return sheet;
  }
}

/** A null-object worksheet stand-in (mirrors `getItemOrNullObject`). */
function nullSheet(): WorksheetLike & { isNullObject: boolean } {
  const noopHub = new EventHub<WorksheetChangedEventArgsLike>();
  return {
    isNullObject: true,
    id: '',
    name: '',
    position: -1,
    visibility: 'Hidden',
    onChanged: noopHub,
    onFormatChanged: noopHub,
    getRange: () => {
      throw new Error('null sheet');
    },
    getRangeByIndexes: () => {
      throw new Error('null sheet');
    },
    activate: () => {
      /* no-op */
    },
    delete: () => {
      /* no-op */
    },
  };
}

export class FakeWorkbook implements WorkbookLike {
  readonly worksheets: FakeWorksheetCollection;
  readonly #sheets: FakeWorksheet[] = [];
  /** Side-channels a test can assert on. */
  untrackCount = 0;
  readonly structuralOps: StructuralOpRecord[] = [];
  syncCount = 0;
  /** The id of the worksheet most recently `.activate()`d, or null. */
  activeSheetId: string | null = null;

  constructor() {
    this.worksheets = new FakeWorksheetCollection(this);
  }

  addSheet(name?: string): FakeWorksheet {
    const id = name ?? `Sheet${String(this.#sheets.length + 1)}`;
    const sheet = new FakeWorksheet(this, id, id, this.#sheets.length);
    this.#sheets.push(sheet);
    return sheet;
  }

  removeSheet(id: string): void {
    const idx = this.#sheets.findIndex((s) => s.id === id);
    if (idx >= 0) {
      this.#sheets.splice(idx, 1);
    }
  }

  findSheet(key: string): FakeWorksheet | undefined {
    return this.#sheets.find((s) => s.id === key || s.name === key);
  }

  get sheets(): readonly FakeWorksheet[] {
    return this.#sheets;
  }
}

/** The mock entry point: an `Excel.run`-shaped function over one shared workbook. */
export function createFakeExcel(workbook: FakeWorkbook = new FakeWorkbook()): {
  run: ExcelRun;
  workbook: FakeWorkbook;
} {
  const context: RequestContextLike = {
    workbook,
    sync: () => {
      workbook.syncCount++;
      return Promise.resolve();
    },
  };
  const run: ExcelRun = (batch) => batch(context);
  return { run, workbook };
}

/** Build a synthetic `WorksheetChangedEventArgs`. */
export function changedEvent(args: {
  worksheetId: string;
  address: string;
  changeType?: ExcelDataChangeType;
  source?: ExcelEventSource;
  triggerSource?: ExcelTriggerSource;
  changeDirectionState?: ChangeDirectionStateLike;
  details?: ChangedEventDetailLike | null;
}): WorksheetChangedEventArgsLike {
  return {
    worksheetId: args.worksheetId,
    address: args.address,
    changeType: args.changeType ?? 'RangeEdited',
    source: args.source ?? 'Local',
    ...(args.triggerSource !== undefined ? { triggerSource: args.triggerSource } : {}),
    ...(args.changeDirectionState !== undefined
      ? { changeDirectionState: args.changeDirectionState }
      : {}),
    ...(args.details !== undefined ? { details: args.details } : {}),
  };
}

/** Parse an A1 address into a zero-based rect (reuses the adapter's parser). */
function parseA1(address: string): {
  startRow: number;
  startCol: number;
  rowCount: number;
  colCount: number;
} {
  return parseAddress(address);
}
