// Full-workbook rollback chrome (ADR-0008): hide every real sheet while Preview
// is active so only the historical (preview) surfaces are visible, then restore
// them exactly on return. The engine flags the transitions (enterPreview /
// exitPreview on the ReconcilePlan); the data source drives enter()/exit().

import { isInternalSheetName } from './office-mapping.ts';
import type { ExcelRun, ExcelSheetVisibility } from './office-types.ts';

export interface PreviewChrome {
  /** Hide all real sheets (remembering their visibility) when Preview begins. */
  enter(): Promise<void>;
  /** Restore the real sheets to their captured visibility when Preview ends. */
  exit(): Promise<void>;
}

/** A no-op chrome for hosts/tests where the full-workbook hide is not wired. */
export const NOOP_PREVIEW_CHROME: PreviewChrome = {
  enter: () => Promise.resolve(),
  exit: () => Promise.resolve(),
};

export class OfficePreviewChrome implements PreviewChrome {
  readonly #run: ExcelRun;
  /** Real sheetId -> visibility captured before hiding (cleared on exit). */
  readonly #hidden = new Map<string, ExcelSheetVisibility>();

  constructor(run: ExcelRun) {
    this.#run = run;
  }

  async enter(): Promise<void> {
    await this.#run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/id,items/name,items/visibility');
      await ctx.sync();
      for (const sheet of sheets.items) {
        // Preview surfaces are already the historical view — leave them; hide the
        // user's real sheets so the workbook shows only the rolled-back state.
        if (isInternalSheetName(sheet.name) || sheet.visibility !== 'Visible') {
          continue;
        }
        this.#hidden.set(sheet.id, sheet.visibility);
        sheet.visibility = 'Hidden';
      }
      await ctx.sync();
    });
  }

  async exit(): Promise<void> {
    if (this.#hidden.size === 0) {
      return;
    }
    const captured = [...this.#hidden.entries()];
    this.#hidden.clear();
    await this.#run(async (ctx) => {
      for (const [sheetId, visibility] of captured) {
        const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetId);
        await ctx.sync();
        if (sheet.isNullObject !== true) {
          sheet.visibility = visibility;
        }
      }
      await ctx.sync();
    });
  }
}
