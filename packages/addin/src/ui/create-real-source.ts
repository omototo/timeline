// Compose the live timeline: engine + store + render targets + change source.
//
// This is the production wiring the bootstrap calls. When not inside an Excel
// host it returns `null` and the bootstrap renders the fake-backed pane instead,
// so the UI still loads in a browser tab or a test.

import { InMemoryStore, TimelineEngineImpl } from '@timeline/engine';
import { getExcelRun, getIsSetSupported } from '../excel/excel-host.ts';
import { ExpectedWriteSet } from '../excel/expected-write-set.ts';
import { OfficeChangeSource } from '../excel/office-change-source.ts';
import type { ExcelRun, RequestContextLike, WorksheetLike } from '../excel/office-types.ts';
import { PreviewSheetRenderTarget, RealSheetRenderTarget } from '../excel/render-target.ts';
import { buildWorkbookSnapshot } from '../excel/workbook-snapshot.ts';
import { RealTimelineDataSource } from './real-data-source.ts';

/**
 * Build the live, Excel-backed data source, or `null` when no Excel host is
 * present. The single `as` cast bridges the real `Excel.run` to the adapters'
 * context-narrowed `ExcelRun` shim — runtime-safe because the adapters only
 * touch members every real Excel proxy provides (this is the slice-3 boundary).
 */
export async function createRealTimelineDataSource(): Promise<RealTimelineDataSource | null> {
  const excelRun = getExcelRun();
  if (!excelRun) {
    return null;
  }
  const runShim = excelRun as unknown as ExcelRun;

  const expectedWrites = new ExpectedWriteSet();
  const engine = new TimelineEngineImpl();
  const store = new InMemoryStore();
  const realTarget = new RealSheetRenderTarget({ run: runShim, expectedWrites });
  const previewTarget = new PreviewSheetRenderTarget({ run: runShim, expectedWrites });

  const { snapshot, sheetIds } = await buildWorkbookSnapshot(excelRun);

  const changeSource = new OfficeChangeSource({
    run: runShim,
    expectedWrites,
    isSetSupported: getIsSetSupported(),
  }).withWorksheetLister((ctx: RequestContextLike): WorksheetLike[] =>
    sheetIds.map((id) => ctx.workbook.worksheets.getItem(id)),
  );

  const source = new RealTimelineDataSource({
    engine,
    store,
    realTarget,
    previewTarget,
    changeSource,
    sheets: sheetIds,
  });

  await source.start(engine.attach(snapshot, null));
  return source;
}
