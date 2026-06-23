/**
 * `RealSheetRenderTarget` / `PreviewSheetRenderTarget` — apply a `ReconcilePlan`
 * to a render surface (ADR-0002 echo cancellation, ADR-0008 frozen values).
 *
 * Both route every write through a single **echo-cancelled choke point**
 * (`#applyOps`): before each `setCells` / structural op it registers the target
 * region in the shared `ExpectedWriteSet`, so the OfficeChangeSource swallows the
 * resulting `onChanged` echo on hosts below ExcelApi 1.14 (at 1.14+ the host's
 * `triggerSource === 'ThisLocalAddin'` does the same job, but registering keeps
 * the choke point uniform across versions).
 *
 * `RealSheet` writes live formulas (`mode: 'formula'`). `PreviewSheet` owns a
 * `veryHidden` worksheet and writes frozen values (`mode: 'value'`), and handles
 * the preview-sheet create / activate / delete lifecycle.
 *
 * Structural ops replay natively via `Range.insert` / `Range.delete` (ADR-0003).
 * Every plan batches its proxy mutations before a single `sync()` and
 * `untrack()`s the ranges it touched.
 */
import type { CellSlab, ReconcileOp, ReconcilePlan, Rect } from '@timeline/engine';
import type { RenderTarget } from './seams.ts';
import type { ExcelRun, RangeLike, RequestContextLike, WorksheetLike } from './office-types.ts';
import { toDeleteShift, toExcelSheetName, toInsertShift } from './office-mapping.ts';
import type { ExpectedWriteSet } from './expected-write-set.ts';

/** Convert a zero-based `Rect` to an A1 address (range or single cell). */
export function rectToAddress(rect: Rect): string {
  const a1 = (row: number, col: number): string => `${colLetters(col)}${String(row + 1)}`;
  const topLeft = a1(rect.startRow, rect.startCol);
  if (rect.rowCount === 1 && rect.colCount === 1) {
    return topLeft;
  }
  const bottomRight = a1(rect.startRow + rect.rowCount - 1, rect.startCol + rect.colCount - 1);
  return `${topLeft}:${bottomRight}`;
}

/** Zero-based column index → letters (0 → A, 26 → AA). */
function colLetters(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export interface RenderTargetOptions {
  run: ExcelRun;
  /** Shared with the OfficeChangeSource for sub-1.14 echo cancellation (ADR-0002). */
  expectedWrites?: ExpectedWriteSet;
}

/**
 * Shared write choke point. Concrete targets pick the write `mode` and decide
 * how to resolve a worksheet from a `sheetId`; everything else (registration,
 * batching, untrack) lives here.
 */
abstract class BaseRenderTarget implements RenderTarget {
  protected readonly run: ExcelRun;
  protected readonly expectedWrites: ExpectedWriteSet | undefined;

  constructor(options: RenderTargetOptions) {
    this.run = options.run;
    this.expectedWrites = options.expectedWrites;
  }

  abstract reconcile(plan: ReconcilePlan): Promise<void>;

  /**
   * Resolve a worksheet from an engine sheet id. Engine preview-surface ids are
   * folded to a legal Excel sheet name first; real sheet ids pass through.
   */
  protected resolveSheet(ctx: RequestContextLike, sheetId: string): WorksheetLike {
    return ctx.workbook.worksheets.getItem(toExcelSheetName(sheetId));
  }

  /**
   * Apply the preview-sheet lifecycle ops (create / activate / delete) carried by
   * a plan. Both targets run these: `goto` creates and activates a preview
   * surface, `returnToPresent` deletes the surfaces and reactivates the real
   * sheet — and that latter plan targets `realSheet`, so the real target must
   * honor them too.
   */
  protected async applyLifecycle(ops: ReconcileOp[]): Promise<void> {
    for (const op of ops) {
      if (op.op === 'createPreviewSheet') {
        await this.#createPreviewSheet(op.previewSheetId);
      } else if (op.op === 'activateSheet') {
        await this.#activateSheet(op.sheetId);
      } else if (op.op === 'deletePreviewSheet') {
        await this.#deletePreviewSheet(op.previewSheetId);
      }
    }
  }

  /** Create the preview sheet (legal name) as `VeryHidden` until activated. */
  #createPreviewSheet(previewSheetId: string): Promise<void> {
    return this.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.add(toExcelSheetName(previewSheetId));
      sheet.visibility = 'VeryHidden';
      await ctx.sync();
    });
  }

  /** Reveal and switch to a sheet — the engine's "land the user here" op. */
  #activateSheet(sheetId: string): Promise<void> {
    return this.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(toExcelSheetName(sheetId));
      sheet.visibility = 'Visible';
      sheet.activate();
      await ctx.sync();
    });
  }

  #deletePreviewSheet(previewSheetId: string): Promise<void> {
    return this.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItemOrNullObject(toExcelSheetName(previewSheetId));
      if (sheet.isNullObject !== true) {
        sheet.delete();
      }
      await ctx.sync();
    });
  }

  /**
   * Apply value/structural ops for one sheet-resolving target. Registers every
   * write region in the expected-write set, stages all proxy mutations, then
   * flushes with a single `sync()` and untracks the touched ranges.
   */
  protected async applyOps(ops: ReconcileOp[]): Promise<void> {
    if (ops.length === 0) {
      return;
    }
    await this.run(async (ctx) => {
      const touched: RangeLike[] = [];
      for (const op of ops) {
        if (op.op === 'setCells') {
          touched.push(...this.#stageSetCells(ctx, op.sheetId, op.area, op.slab, op.mode));
        } else if (op.op === 'applyStructural') {
          touched.push(this.#stageStructural(ctx, op));
        }
      }
      await ctx.sync();
      for (const range of touched) {
        range.untrack();
      }
      if (touched.length > 0) {
        await ctx.sync();
      }
    });
  }

  /** Stage a `setCells` op across each rectangle of its area; returns touched ranges. */
  #stageSetCells(
    ctx: RequestContextLike,
    sheetId: string,
    area: Rect[],
    slab: CellSlab,
    mode: 'value' | 'formula',
  ): RangeLike[] {
    const sheet = this.resolveSheet(ctx, sheetId);
    const ranges: RangeLike[] = [];
    let rowOffset = 0;
    for (const rect of area) {
      const address = rectToAddress(rect);
      this.expectedWrites?.register(sheet.id, address);
      const range = sheet.getRange(address);
      const block = sliceSlab(slab, rowOffset, rect);
      if (mode === 'formula') {
        // Live writes: prefer the formula text; fall back to the literal value
        // for cells that hold a constant (formula === null).
        range.formulas = block.formulas.map((row, r) =>
          row.map((f, c) => f ?? block.values[r]?.[c] ?? null),
        );
      } else {
        // Frozen values (ADR-0008): write the computed values, not formulas.
        range.values = block.values;
      }
      range.numberFormat = block.numberFormats;
      ranges.push(range);
      rowOffset += rect.rowCount;
    }
    return ranges;
  }

  /** Stage a structural insert/delete; returns the touched range. */
  #stageStructural(
    ctx: RequestContextLike,
    op: Extract<ReconcileOp, { op: 'applyStructural' }>,
  ): RangeLike {
    const sheet = this.resolveSheet(ctx, op.sheetId);
    const address = rectToAddress(op.address);
    this.expectedWrites?.register(sheet.id, address);
    const range = sheet.getRange(address);
    const inserting =
      op.changeType === 'rowInserted' ||
      op.changeType === 'columnInserted' ||
      op.changeType === 'cellInserted';
    if (inserting) {
      range.insert(toInsertShift(op.shiftDirection));
    } else {
      range.delete(toDeleteShift(op.shiftDirection));
    }
    return range;
  }
}

/** Writes live formulas to the real sheet (`mode: 'formula'`). */
export class RealSheetRenderTarget extends BaseRenderTarget {
  async reconcile(plan: ReconcilePlan): Promise<void> {
    // `returnToPresent` routes its delete-preview / activate-real ops through a
    // realSheet-targeted plan, so process the lifecycle here too.
    await this.applyLifecycle(plan.ops);
    const writeOps = plan.ops.filter((op) => op.op === 'setCells' || op.op === 'applyStructural');
    await this.applyOps(writeOps);
  }
}

/**
 * Owns an engine-owned `veryHidden` preview worksheet and writes frozen values
 * (`mode: 'value'`). Handles `createPreviewSheet` / `activateSheet` /
 * `deletePreviewSheet`.
 */
export class PreviewSheetRenderTarget extends BaseRenderTarget {
  async reconcile(plan: ReconcilePlan): Promise<void> {
    // Lifecycle ops run first (a setCells may target a just-created sheet),
    // then the frozen-value writes for the preview surface.
    await this.applyLifecycle(plan.ops);
    const writeOps = plan.ops.filter((op) => op.op === 'setCells' || op.op === 'applyStructural');
    await this.applyOps(writeOps);
  }
}

/** Slice the `rowOffset..rowOffset+rect.rowCount` rows of a slab into a block. */
function sliceSlab(slab: CellSlab, rowOffset: number, rect: Rect): CellSlab {
  const end = rowOffset + rect.rowCount;
  return {
    values: slab.values.slice(rowOffset, end),
    formulas: slab.formulas.slice(rowOffset, end),
    numberFormats: slab.numberFormats.slice(rowOffset, end),
    valueTypes: slab.valueTypes.slice(rowOffset, end),
  };
}
