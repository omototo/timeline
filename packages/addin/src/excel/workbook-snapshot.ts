// Read the live workbook into the engine's `WorkbookSnapshot` for `attach`.
//
// On launch the engine needs the current state of every worksheet so future
// edits diff against real content (and, on resume, so it can detect drift). We
// read each sheet's used range into a `CellSlab` via the same `slabFromRange`
// mapping the change source uses for its read-backs, keeping the two paths
// consistent.

import type { CellSlab, WorkbookSnapshot } from '@timeline/engine';
import type { RealExcelRun } from './excel-host.ts';
import { isInternalSheetName, slabFromRange } from './office-mapping.ts';
import type { RangeLike } from './office-types.ts';

export interface WorkbookSnapshotResult {
  readonly snapshot: WorkbookSnapshot;
  readonly sheetIds: string[];
}

function emptySlab(): CellSlab {
  return { values: [], formulas: [], numberFormats: [], valueTypes: [] };
}

/**
 * A stable FNV-1a hash of the observed sheet slabs. Used by the engine to tell a
 * clean resume from drift; with an in-memory store (fresh each launch) it is
 * never compared, but we compute it honestly so a persistent store works later.
 */
function hashSheets(sheets: WorkbookSnapshot['sheets']): string {
  let hash = 0x811c9dc5;
  const json = JSON.stringify(sheets);
  for (let i = 0; i < json.length; i += 1) {
    hash ^= json.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16);
}

/** Read every worksheet's used range into a `WorkbookSnapshot`. */
export async function buildWorkbookSnapshot(run: RealExcelRun): Promise<WorkbookSnapshotResult> {
  return run(async (ctx) => {
    const collection = ctx.workbook.worksheets;
    collection.load('items/id,items/name');
    await ctx.sync();

    const sheetIds: string[] = [];
    const sheets: { sheetId: string; slab: CellSlab }[] = [];
    for (const sheet of collection.items) {
      // Engine-owned preview surfaces (e.g. a leftover after a crash) are NOT
      // part of the user's workbook — never seed the baseline from them.
      if (isInternalSheetName(sheet.name)) {
        continue;
      }
      sheetIds.push(sheet.id);
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load(['values', 'formulas', 'numberFormat', 'valueTypes', 'isNullObject']);
      await ctx.sync();
      sheets.push({
        sheetId: sheet.id,
        slab: used.isNullObject ? emptySlab() : slabFromRange(used as unknown as RangeLike),
      });
    }

    const snapshot: WorkbookSnapshot = {
      workbookGuid: 'workbook',
      contentHash: hashSheets(sheets),
      sheets,
    };
    return { snapshot, sheetIds };
  });
}
