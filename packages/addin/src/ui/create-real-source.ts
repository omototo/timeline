// Compose the live timeline: engine + store + render targets + change source.
//
// This is the production wiring the bootstrap calls. When not inside an Excel
// host it returns `null` and the bootstrap renders the fake-backed pane instead,
// so the UI still loads in a browser tab or a test.

import {
  InMemoryStore,
  TimelineEngineImpl,
  type HistoryStore,
  type RehydrationData,
} from '@timeline/engine';
import {
  getExcelRun,
  getIsSetSupported,
  newWorkbookGuid,
  resolveWorkbookKey,
} from '../excel/excel-host.ts';
import { ExpectedWriteSet } from '../excel/expected-write-set.ts';
import { databaseNameFor, IndexedDbStore } from '../excel/indexeddb-store.ts';
import { OfficeChangeSource } from '../excel/office-change-source.ts';
import type { ExcelRun, RequestContextLike, WorksheetLike } from '../excel/office-types.ts';
import { OfficePreviewChrome } from '../excel/preview-chrome.ts';
import { PreviewSheetRenderTarget, RealSheetRenderTarget } from '../excel/render-target.ts';
import { buildWorkbookSnapshot } from '../excel/workbook-snapshot.ts';
import { showErrorBanner } from './error-banner.ts';
import { RealTimelineDataSource } from './real-data-source.ts';

/** Load the persisted timeline back into a `RehydrationData` payload. */
export async function loadRehydrationData(store: HistoryStore): Promise<RehydrationData> {
  const head = await store.getHead();
  const branches = await store.listBranches();
  const branchIds = new Set<string>(['main', ...branches.map((b) => b.id)]);
  const perBranch = [];
  for (const branchId of branchIds) {
    perBranch.push({
      branchId,
      deltas: await store.loadDeltas(branchId, 0, Number.MAX_SAFE_INTEGER),
      keyframes: await store.listKeyframes(branchId),
    });
  }
  return { head, branches, perBranch };
}

/**
 * Open the durable per-workbook IndexedDB store, falling back to an in-memory
 * store (no cross-session persistence) when IndexedDB is unavailable or fails to
 * open — the timeline still works for the session.
 */
async function openStore(): Promise<HistoryStore> {
  if (typeof globalThis.indexedDB === 'undefined') {
    return new InMemoryStore();
  }
  try {
    // Isolate per workbook. A session-unique key (never the shared default) when
    // the workbook can't be identified, so a new document never inherits another
    // document's history — it gets a fresh, non-persistent timeline instead.
    const workbookKey = (await resolveWorkbookKey()) ?? `session-${newWorkbookGuid()}`;
    const store = new IndexedDbStore(globalThis.indexedDB, databaseNameFor(workbookKey));
    await store.init();
    return store;
  } catch (error) {
    globalThis.console.error('[timeline] IndexedDB unavailable; history will not persist', error);
    showErrorBanner('History storage is unavailable; changes will not persist across reloads.');
    return new InMemoryStore();
  }
}

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
  const store = await openStore();
  const realTarget = new RealSheetRenderTarget({ run: runShim, expectedWrites });
  const previewTarget = new PreviewSheetRenderTarget({ run: runShim, expectedWrites });

  // Restore the persisted timeline BEFORE attach reseeds the Shadow State from
  // the live workbook, so prior history survives a reload / Excel restart.
  engine.rehydrate(await loadRehydrationData(store));

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
    chrome: new OfficePreviewChrome(runShim),
    sheets: sheetIds,
    onError: (error) => {
      const message = error instanceof Error ? error.message : String(error);
      globalThis.console.error('[timeline] reconcile failed', error);
      showErrorBanner(message);
    },
  });

  await source.start(engine.attach(snapshot, null));
  return source;
}
