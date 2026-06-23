/**
 * `OfficeChangeSource` — turns Office.js worksheet/collection events into
 * engine-neutral `Observation`s, one batch per debounced user action
 * (ADR-0013, ADR-0001).
 *
 * Responsibilities (Stream B, slice 2):
 *   - Register `Worksheet.onChanged` / `onFormatChanged` plus worksheet-collection
 *     add / delete / rename (and reorder where available).
 *   - Debounce the burst of one user action into one Step's Observations. The
 *     window is a constructor option.
 *   - Read back the changed region to build the `after`-slab, chunked to respect
 *     the 5,000,000-cell (all-platform) and 5 MB (web) limits (findings §2):
 *     check `cellCount` before `sync()`, tile if over, `untrack()` proxies.
 *   - Echo-filter our own writes: drop events where `triggerSource ===
 *     'ThisLocalAddin'` (1.14, feature-detected); fall back to the shared
 *     expected-write set below 1.14 (ADR-0002).
 *   - Signal co-authoring: a `source: 'remote'` event invokes the supplied
 *     callback (so the engine can suspend tracking — ADR-0006).
 *
 * No Office.js type crosses the `ChangeSource` seam: events become `Observation`s
 * before they leave this class.
 */
import type {
  Observation,
  Rect,
  StructuralObservation,
  ValueObservation,
  WorksheetObservation,
} from '@timeline/engine';
import type { ChangeSource } from './seams.ts';
import type {
  ExcelRun,
  EventHandlerResultLike,
  IsSetSupported,
  WorksheetChangedEventArgsLike,
  WorksheetAddedEventArgsLike,
  WorksheetDeletedEventArgsLike,
  WorksheetNameChangedEventArgsLike,
  WorksheetPositionChangedEventArgsLike,
  WorksheetLike,
  RequestContextLike,
} from './office-types.ts';
import {
  parseAddress,
  slabFromRange,
  toShiftDirection,
  toStructuralChangeType,
  type ParsedRect,
} from './office-mapping.ts';
import type { ExpectedWriteSet } from './expected-write-set.ts';

/** Default debounce window for one user action's event burst. */
// TODO(spike): tune via event-fan-out spike (findings §A9 / Step debounce).
const DEFAULT_DEBOUNCE_MS = 300;

/**
 * Per-get cell ceiling. The real Excel limit is 5,000,000 cells (all platforms);
 * tests lower it to exercise the chunking path.
 */
const DEFAULT_MAX_CELLS_PER_READ = 5_000_000;

export interface OfficeChangeSourceOptions {
  /** Office.js `Excel.run` shim (injected for headless tests). */
  run: ExcelRun;
  /** Feature-detect ExcelApi 1.14 `triggerSource` (default: assume present). */
  isSetSupported?: IsSetSupported;
  /** Shared expected-write set for the sub-1.14 echo fallback. */
  expectedWrites?: ExpectedWriteSet;
  /** Debounce window in ms (default 300). */
  debounceMs?: number;
  /** Per-get cell cap; tests lower it to drive the chunked read-back. */
  maxCellsPerRead?: number;
  /** Invoked once per `source: 'remote'` event (co-authoring; ADR-0006). */
  onRemoteChange?: (observation: Observation) => void;
  /** Injectable timer (default `setTimeout`/`clearTimeout`) for deterministic tests. */
  setTimer?: (callback: () => void, ms: number) => ReturnType<typeof setTimeout>;
  clearTimer?: (handle: ReturnType<typeof setTimeout>) => void;
}

export class OfficeChangeSource implements ChangeSource {
  readonly #run: ExcelRun;
  readonly #isSetSupported: IsSetSupported;
  readonly #expectedWrites: ExpectedWriteSet | undefined;
  readonly #debounceMs: number;
  readonly #maxCellsPerRead: number;
  readonly #onRemoteChange: ((observation: Observation) => void) | undefined;
  readonly #setTimer: (callback: () => void, ms: number) => ReturnType<typeof setTimeout>;
  readonly #clearTimer: (handle: ReturnType<typeof setTimeout>) => void;

  #handler: ((obs: Observation) => void) | null = null;
  #registrations: EventHandlerResultLike[] = [];
  /** Observations accumulated within the current debounce window. */
  #pending: Observation[] = [];
  #flushTimer: ReturnType<typeof setTimeout> | null = null;
  #has114: boolean;

  constructor(options: OfficeChangeSourceOptions) {
    this.#run = options.run;
    this.#isSetSupported = options.isSetSupported ?? (() => true);
    this.#expectedWrites = options.expectedWrites;
    this.#debounceMs = options.debounceMs ?? DEFAULT_DEBOUNCE_MS;
    this.#maxCellsPerRead = options.maxCellsPerRead ?? DEFAULT_MAX_CELLS_PER_READ;
    this.#onRemoteChange = options.onRemoteChange;
    this.#setTimer = options.setTimer ?? ((cb, ms) => setTimeout(cb, ms));
    this.#clearTimer =
      options.clearTimer ??
      ((h) => {
        clearTimeout(h);
      });
    this.#has114 = this.#isSetSupported('ExcelApi', '1.14');
  }

  /**
   * Optionally list every existing worksheet so per-sheet value/format handlers
   * can be attached. Defaults to no extra sheets — a host that surfaces all
   * value edits through a single representative sheet may leave this unset.
   */
  #listWorksheets: ((ctx: RequestContextLike) => WorksheetLike[]) | undefined;

  async start(onObservation: (obs: Observation) => void): Promise<void> {
    this.#handler = onObservation;
    await this.#run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      // Collection-level handlers cover sheet add/delete/rename/reorder.
      this.#registrations.push(
        sheets.onAdded.add((args) => {
          this.#onWorksheetAdded(args);
        }),
        sheets.onDeleted.add((args) => {
          this.#onWorksheetDeleted(args);
        }),
        sheets.onNameChanged.add((args) => {
          this.#onWorksheetNameChanged(args);
        }),
      );
      if (sheets.onMoved !== undefined) {
        this.#registrations.push(
          sheets.onMoved.add((args) => {
            this.#onWorksheetMoved(args);
          }),
        );
      }
      // Per-sheet value/format handlers for every existing worksheet.
      for (const sheet of this.#listWorksheets?.(ctx) ?? []) {
        this.#registerSheetHandlers(sheet);
      }
      await ctx.sync();
    });
  }

  /** Provide the per-launch list of worksheets to attach value/format handlers to. */
  withWorksheetLister(lister: (ctx: RequestContextLike) => WorksheetLike[]): this {
    this.#listWorksheets = lister;
    return this;
  }

  /** Attach value + format `onChanged` handlers to one worksheet. */
  #registerSheetHandlers(sheet: WorksheetLike): void {
    // Office.js awaits a promise-returning handler; returning it (rather than
    // `void`-ing) lets the async read-back settle before the next event runs.
    this.#registrations.push(
      sheet.onChanged.add((args) => this.#onChanged(args)),
      sheet.onFormatChanged.add((args) => this.#onChanged(args)),
    );
  }

  stop(): Promise<void> {
    if (this.#flushTimer !== null) {
      this.#clearTimer(this.#flushTimer);
      this.#flushTimer = null;
    }
    for (const reg of this.#registrations) {
      reg.remove();
    }
    this.#registrations = [];
    this.#pending = [];
    this.#handler = null;
    return Promise.resolve();
  }

  /** Handle a value/format `onChanged` event: echo-filter, read back, enqueue. */
  async #onChanged(args: WorksheetChangedEventArgsLike): Promise<void> {
    if (this.#isEcho(args)) {
      return;
    }
    const structural = toStructuralChangeType(args.changeType);
    const observation =
      structural === null
        ? await this.#buildValueObservation(args)
        : this.#buildStructuralObservation(args, structural);
    this.#enqueue(observation, args.source === 'Remote');
  }

  /** Drop our own write-echoes (1.14 `triggerSource`, else expected-write set). */
  #isEcho(args: WorksheetChangedEventArgsLike): boolean {
    if (this.#has114 && args.triggerSource === 'ThisLocalAddin') {
      return true;
    }
    if (!this.#has114 && this.#expectedWrites !== undefined) {
      return this.#expectedWrites.consume(args.worksheetId, args.address);
    }
    return false;
  }

  /** Build a `ValueObservation`, reading back the changed region (chunked). */
  async #buildValueObservation(args: WorksheetChangedEventArgsLike): Promise<ValueObservation> {
    const rect = parseAddress(args.address);
    const after = await this.#readBack(args.worksheetId, rect);
    return {
      kind: 'value',
      sheetId: args.worksheetId,
      area: [toRect(rect)],
      after,
      triggerSource: args.triggerSource === 'ThisLocalAddin' ? 'thisLocalAddin' : 'unknown',
      source: args.source === 'Remote' ? 'remote' : 'local',
    };
  }

  /**
   * Read back the changed region into a single `CellSlab`, tiling the read into
   * sub-blocks whose cell count stays under the per-get limit (findings §2).
   * Each tile is loaded, `sync()`ed, and `untrack()`ed before the next.
   */
  async #readBack(sheetId: string, rect: ParsedRect): Promise<ValueObservation['after']> {
    const totalCells = rect.rowCount * rect.colCount;
    if (totalCells <= this.#maxCellsPerRead) {
      return this.#readTile(sheetId, rect);
    }
    // Tile by rows: the largest whole-row block that fits the cell budget.
    const rowsPerTile = Math.max(1, Math.floor(this.#maxCellsPerRead / rect.colCount));
    const values: ValueObservation['after']['values'] = [];
    const formulas: ValueObservation['after']['formulas'] = [];
    const numberFormats: ValueObservation['after']['numberFormats'] = [];
    const valueTypes: ValueObservation['after']['valueTypes'] = [];
    for (let row = 0; row < rect.rowCount; row += rowsPerTile) {
      const tileRows = Math.min(rowsPerTile, rect.rowCount - row);
      const tile = await this.#readTile(sheetId, {
        startRow: rect.startRow + row,
        startCol: rect.startCol,
        rowCount: tileRows,
        colCount: rect.colCount,
      });
      values.push(...tile.values);
      formulas.push(...tile.formulas);
      numberFormats.push(...tile.numberFormats);
      valueTypes.push(...tile.valueTypes);
    }
    return { values, formulas, numberFormats, valueTypes };
  }

  /** Read one tile that is known to fit under the cell budget. */
  #readTile(sheetId: string, rect: ParsedRect): Promise<ValueObservation['after']> {
    return this.#run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(sheetId);
      const range = sheet.getRangeByIndexes(
        rect.startRow,
        rect.startCol,
        rect.rowCount,
        rect.colCount,
      );
      range.load(['cellCount']);
      await ctx.sync();
      // Defensive: a multi-row tile that still exceeds the budget is split by
      // halving rows. A single row cannot be split further by row (Excel caps
      // columns at 16,384, well under the 5,000,000-cell limit), so it is read
      // as-is — this also avoids unbounded recursion.
      if (range.cellCount > this.#maxCellsPerRead && rect.rowCount > 1) {
        const half = Math.floor(rect.rowCount / 2);
        const top = await this.#readBack(sheetId, { ...rect, rowCount: half });
        const bottom = await this.#readBack(sheetId, {
          ...rect,
          startRow: rect.startRow + half,
          rowCount: rect.rowCount - half,
        });
        return {
          values: [...top.values, ...bottom.values],
          formulas: [...top.formulas, ...bottom.formulas],
          numberFormats: [...top.numberFormats, ...bottom.numberFormats],
          valueTypes: [...top.valueTypes, ...bottom.valueTypes],
        };
      }
      range.load(['values', 'formulas', 'numberFormat', 'valueTypes']);
      await ctx.sync();
      const slab = slabFromRange(range);
      range.untrack();
      await ctx.sync();
      return slab;
    });
  }

  /** Build a `StructuralObservation` from a structural `changeType`. */
  #buildStructuralObservation(
    args: WorksheetChangedEventArgsLike,
    changeType: StructuralObservation['changeType'],
  ): StructuralObservation {
    const shift = toShiftDirection(args.changeDirectionState);
    const base: StructuralObservation = {
      kind: 'structural',
      sheetId: args.worksheetId,
      changeType,
      address: toRect(parseAddress(args.address)),
      triggerSource: args.triggerSource === 'ThisLocalAddin' ? 'thisLocalAddin' : 'unknown',
      source: args.source === 'Remote' ? 'remote' : 'local',
    };
    return shift === undefined ? base : { ...base, shiftDirection: shift };
  }

  #onWorksheetAdded(args: WorksheetAddedEventArgsLike): void {
    this.#enqueue(
      this.#worksheetObservation('add', args.worksheetId, args.source),
      args.source === 'Remote',
    );
  }

  #onWorksheetDeleted(args: WorksheetDeletedEventArgsLike): void {
    this.#enqueue(
      this.#worksheetObservation('delete', args.worksheetId, args.source),
      args.source === 'Remote',
    );
  }

  #onWorksheetNameChanged(args: WorksheetNameChangedEventArgsLike): void {
    const base = this.#worksheetObservation('rename', args.worksheetId, args.source);
    const obs: WorksheetObservation =
      args.nameAfter === undefined ? base : { ...base, newName: args.nameAfter };
    this.#enqueue(obs, args.source === 'Remote');
  }

  #onWorksheetMoved(args: WorksheetPositionChangedEventArgsLike): void {
    const base = this.#worksheetObservation('reorder', args.worksheetId, args.source);
    const obs: WorksheetObservation =
      args.positionAfter === undefined ? base : { ...base, newPosition: args.positionAfter };
    this.#enqueue(obs, args.source === 'Remote');
  }

  #worksheetObservation(
    op: WorksheetObservation['op'],
    sheetId: string,
    source: 'Local' | 'Remote',
  ): WorksheetObservation {
    return {
      kind: 'worksheet',
      op,
      sheetId,
      triggerSource: 'unknown',
      source: source === 'Remote' ? 'remote' : 'local',
    };
  }

  /**
   * Add an Observation to the current debounce window. A remote-sourced change
   * additionally invokes the co-authoring callback so the engine can suspend.
   */
  #enqueue(observation: Observation, isRemote: boolean): void {
    if (this.#handler === null) {
      return;
    }
    if (isRemote) {
      this.#onRemoteChange?.(observation);
    }
    this.#pending.push(observation);
    if (this.#flushTimer !== null) {
      this.#clearTimer(this.#flushTimer);
    }
    this.#flushTimer = this.#setTimer(() => {
      this.#flush();
    }, this.#debounceMs);
  }

  /** Emit each debounced Observation to the handler, oldest first. */
  #flush(): void {
    this.#flushTimer = null;
    const batch = this.#pending;
    this.#pending = [];
    const handler = this.#handler;
    if (handler === null) {
      return;
    }
    for (const obs of batch) {
      handler(obs);
    }
    this.#expectedWrites?.prune();
  }
}

/** Convert a parsed rect into the engine's `Rect`. */
function toRect(rect: ParsedRect): Rect {
  return {
    startRow: rect.startRow,
    startCol: rect.startCol,
    rowCount: rect.rowCount,
    colCount: rect.colCount,
  };
}
