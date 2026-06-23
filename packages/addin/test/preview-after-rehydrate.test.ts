import { describe, expect, it } from 'vitest';
import { InMemoryStore, TimelineEngineImpl } from '@timeline/engine';
import type {
  CellSlab,
  EffectEnvelope,
  HistoryStore,
  RehydrationData,
  ValueObservation,
  WorkbookSnapshot,
} from '@timeline/engine';
import { ExpectedWriteSet } from '../src/excel/expected-write-set.ts';
import { PreviewSheetRenderTarget, RealSheetRenderTarget } from '../src/excel/render-target.ts';
import { createFakeExcel, FakeWorkbook } from './excel/fake-excel.ts';

const SHEET = 'S1';

function cell(value: unknown): CellSlab {
  return {
    values: [[value]],
    formulas: [[null]],
    numberFormats: [['General']],
    valueTypes: [['number']],
  };
}

function emptySnapshot(): WorkbookSnapshot {
  return {
    workbookGuid: 'wb',
    contentHash: 'h',
    sheets: [
      {
        sheetId: SHEET,
        slab: {
          values: [[null]],
          formulas: [[null]],
          numberFormats: [['General']],
          valueTypes: [['empty']],
        },
      },
    ],
  };
}

function edit(row: number, col: number, value: number): ValueObservation {
  return {
    kind: 'value',
    triggerSource: 'unknown',
    source: 'local',
    sheetId: SHEET,
    area: [{ startRow: row, startCol: col, rowCount: 1, colCount: 1 }],
    after: cell(value),
  };
}

async function persist(store: HistoryStore, env: EffectEnvelope): Promise<void> {
  for (const op of env.persist ?? []) {
    if (op.op === 'appendDelta') await store.appendDelta(op.branchId, op.delta);
    else if (op.op === 'writeKeyframe')
      await store.writeKeyframe(op.branchId, op.stepIndex, op.state);
    else if (op.op === 'setHead') await store.setHead(op.head);
    else if (op.op === 'saveBranch') await store.saveBranch(op.meta);
    else await store.deleteBranch(op.branchId);
  }
}

async function loadRehydrationData(store: HistoryStore): Promise<RehydrationData> {
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
  return { head: await store.getHead(), branches, perBranch };
}

describe('preview after rehydrate (end-to-end against a fake workbook)', () => {
  it('previews a restored past step and returns to present without an Excel error', async () => {
    const workbook = new FakeWorkbook();
    workbook.addSheet(SHEET);
    const { run } = createFakeExcel(workbook);
    const expectedWrites = new ExpectedWriteSet();
    const realTarget = new RealSheetRenderTarget({ run, expectedWrites });
    const previewTarget = new PreviewSheetRenderTarget({ run, expectedWrites });

    const apply = async (env: EffectEnvelope): Promise<void> => {
      if (env.reconcile) {
        await (env.reconcile.target === 'previewSheet' ? previewTarget : realTarget).reconcile(
          env.reconcile,
        );
      }
    };

    // Session 1: H12 first, then a few more edits.
    const store = new InMemoryStore();
    const engine1 = new TimelineEngineImpl();
    engine1.attach(emptySnapshot(), null);
    await persist(store, engine1.ingest(edit(11, 7, 42))); // H12 = 42 (step 0)
    await persist(store, engine1.ingest(edit(14, 15, 5))); // P15 (step 1)
    await persist(store, engine1.ingest(edit(15, 15, 4))); // P16 (step 2)
    await persist(store, engine1.ingest(edit(16, 15, 5))); // P17 (step 3)

    // Session 2: reload — rehydrate from the store, attach to the live workbook.
    const engine2 = new TimelineEngineImpl();
    engine2.rehydrate(await loadRehydrationData(store));
    engine2.attach(emptySnapshot(), null);

    expect(engine2.timeline().steps).toHaveLength(4);

    // Preview a restored past step, then return — neither must throw.
    await apply(engine2.goto({ branchId: 'main', stepIndex: 1 }));
    await apply(engine2.returnToPresent());
    await apply(engine2.goto({ branchId: 'main', stepIndex: 3 }));
    await apply(engine2.returnToPresent());
  });
});
