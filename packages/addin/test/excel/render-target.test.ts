import { describe, it, expect, beforeEach } from 'vitest';
import type { CellSlab, ReconcilePlan } from '@timeline/engine';
import {
  RealSheetRenderTarget,
  PreviewSheetRenderTarget,
  rectToAddress,
} from '../../src/excel/render-target.ts';
import { ExpectedWriteSet } from '../../src/excel/expected-write-set.ts';
import { createFakeExcel, FakeWorkbook, type FakeWorksheet } from './fake-excel.ts';

/** Resolve a sheet, asserting it exists (keeps the test free of `!`). */
function requireSheet(workbook: FakeWorkbook, id: string): FakeWorksheet {
  const sheet = workbook.findSheet(id);
  expect(sheet).toBeDefined();
  if (sheet === undefined) {
    throw new Error(`test setup: sheet "${id}" missing`);
  }
  return sheet;
}

function slab(values: unknown[][], formulas: (string | null)[][]): CellSlab {
  return {
    values,
    formulas,
    numberFormats: values.map((row) => row.map(() => 'General')),
    valueTypes: values.map((row) => row.map(() => 'number')),
  };
}

describe('rectToAddress', () => {
  it('renders a single cell and a range', () => {
    expect(rectToAddress({ startRow: 0, startCol: 0, rowCount: 1, colCount: 1 })).toBe('A1');
    expect(rectToAddress({ startRow: 1, startCol: 1, rowCount: 2, colCount: 2 })).toBe('B2:C3');
  });

  it('handles multi-letter columns', () => {
    expect(rectToAddress({ startRow: 0, startCol: 26, rowCount: 1, colCount: 1 })).toBe('AA1');
  });
});

describe('RealSheetRenderTarget', () => {
  let workbook: FakeWorkbook;
  let run: ReturnType<typeof createFakeExcel>['run'];

  beforeEach(() => {
    const fake = createFakeExcel(new FakeWorkbook());
    workbook = fake.workbook;
    run = fake.run;
    workbook.addSheet('Sheet1');
  });

  it('applies setCells in formula mode and untracks', async () => {
    const target = new RealSheetRenderTarget({ run });
    const plan: ReconcilePlan = {
      target: 'realSheet',
      ops: [
        {
          op: 'setCells',
          sheetId: 'Sheet1',
          area: [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 2 }],
          slab: slab([[1, 2]], [['=A1+1', null]]),
          mode: 'formula',
        },
      ],
    };
    await target.reconcile(plan);

    const sheet = requireSheet(workbook, 'Sheet1');
    expect(sheet.cellAt(0, 0).formula).toBe('=A1+1');
    expect(sheet.cellAt(0, 1).value).toBe(2); // constant cell falls back to literal value
    expect(workbook.untrackCount).toBe(1);
  });

  it('registers writes in the expected-write set (echo cancellation)', async () => {
    const expectedWrites = new ExpectedWriteSet({ now: () => 0 });
    const target = new RealSheetRenderTarget({ run, expectedWrites });
    await target.reconcile({
      target: 'realSheet',
      ops: [
        {
          op: 'setCells',
          sheetId: 'Sheet1',
          area: [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 1 }],
          slab: slab([[5]], [[null]]),
          mode: 'formula',
        },
      ],
    });
    expect(expectedWrites.consume('Sheet1', 'A1')).toBe(true);
  });

  it('applies a structural row insert via Range.insert', async () => {
    const target = new RealSheetRenderTarget({ run });
    await target.reconcile({
      target: 'realSheet',
      ops: [
        {
          op: 'applyStructural',
          sheetId: 'Sheet1',
          changeType: 'rowInserted',
          address: { startRow: 2, startCol: 0, rowCount: 1, colCount: 1 },
          shiftDirection: 'down',
        },
      ],
    });
    expect(workbook.structuralOps).toEqual([
      { sheetId: 'Sheet1', kind: 'insert', shift: 'Down', startRow: 2, startCol: 0 },
    ]);
  });

  it('applies a structural column delete via Range.delete', async () => {
    const target = new RealSheetRenderTarget({ run });
    await target.reconcile({
      target: 'realSheet',
      ops: [
        {
          op: 'applyStructural',
          sheetId: 'Sheet1',
          changeType: 'columnDeleted',
          address: { startRow: 0, startCol: 1, rowCount: 1, colCount: 1 },
          shiftDirection: 'left',
        },
      ],
    });
    expect(workbook.structuralOps[0]).toMatchObject({ kind: 'delete', shift: 'Left' });
  });

  it('is a no-op for a plan with no write ops', async () => {
    const target = new RealSheetRenderTarget({ run });
    await target.reconcile({ target: 'realSheet', ops: [] });
    expect(workbook.untrackCount).toBe(0);
  });
});

describe('PreviewSheetRenderTarget', () => {
  let workbook: FakeWorkbook;
  let run: ReturnType<typeof createFakeExcel>['run'];

  beforeEach(() => {
    const fake = createFakeExcel(new FakeWorkbook());
    workbook = fake.workbook;
    run = fake.run;
  });

  it('creates a veryHidden preview sheet, writes values, then deletes it', async () => {
    const target = new PreviewSheetRenderTarget({ run });

    await target.reconcile({
      target: 'previewSheet',
      ops: [{ op: 'createPreviewSheet', previewSheetId: 'Preview' }],
    });
    const preview = requireSheet(workbook, 'Preview');
    expect(preview.visibility).toBe('VeryHidden');

    await target.reconcile({
      target: 'previewSheet',
      ops: [
        {
          op: 'setCells',
          sheetId: 'Preview',
          area: [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 1 }],
          slab: slab([[99]], [['=NOW()']]),
          mode: 'value',
        },
      ],
    });
    // Frozen value (ADR-0008): the computed value is written, not the formula.
    expect(preview.cellAt(0, 0).value).toBe(99);
    expect(preview.cellAt(0, 0).formula).toBeUndefined();

    await target.reconcile({
      target: 'previewSheet',
      ops: [{ op: 'deletePreviewSheet', previewSheetId: 'Preview' }],
    });
    expect(workbook.findSheet('Preview')).toBeUndefined();
  });

  it('activates a sheet by revealing it and switching the view to it', async () => {
    const target = new PreviewSheetRenderTarget({ run });
    const sheet = workbook.addSheet('Preview');
    sheet.visibility = 'VeryHidden';
    await target.reconcile({
      target: 'previewSheet',
      ops: [{ op: 'activateSheet', sheetId: 'Preview' }],
    });
    expect(sheet.visibility).toBe('Visible');
    expect(workbook.activeSheetId).toBe('Preview');
  });

  it('re-creating over a stale preview sheet replaces it (idempotent, no collision)', async () => {
    const target = new PreviewSheetRenderTarget({ run });
    const plan: ReconcilePlan = {
      target: 'previewSheet',
      ops: [{ op: 'createPreviewSheet', previewSheetId: 'Preview' }],
    };
    await target.reconcile(plan);
    // A second create (as after a branch left the surface behind) must not throw.
    await target.reconcile(plan);
    expect(workbook.sheets.filter((s) => s.name === 'Preview')).toHaveLength(1);
    expect(requireSheet(workbook, 'Preview').visibility).toBe('VeryHidden');
  });

  it('tolerates deleting an already-absent preview sheet', async () => {
    const target = new PreviewSheetRenderTarget({ run });
    await target.reconcile({
      target: 'previewSheet',
      ops: [{ op: 'deletePreviewSheet', previewSheetId: 'Ghost' }],
    });
    expect(workbook.findSheet('Ghost')).toBeUndefined();
  });

  it('folds an engine preview id (with ":" and a GUID) into a legal Excel name', async () => {
    const target = new PreviewSheetRenderTarget({ run });
    // The engine emits ids like "__preview__::{GUID}" — illegal as a sheet name.
    const previewId = '__preview__::{6F9619FF-8B86-D011-B42D-00CF4FC964FF}';
    await target.reconcile({
      target: 'previewSheet',
      ops: [{ op: 'createPreviewSheet', previewSheetId: previewId }],
    });
    // A sheet was created under a folded, legal name and is resolvable by the id.
    const created = workbook.sheets[0];
    expect(created).toBeDefined();
    expect(created?.name).not.toContain(':');
    expect((created?.name ?? '').length).toBeLessThanOrEqual(31);
    // Writing to the same engine id lands on that same sheet.
    await target.reconcile({
      target: 'previewSheet',
      ops: [
        {
          op: 'setCells',
          sheetId: previewId,
          area: [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 1 }],
          slab: slab([[42]], [[null]]),
          mode: 'value',
        },
      ],
    });
    expect(created?.cellAt(0, 0).value).toBe(42);
  });
});

describe('RealSheetRenderTarget — returnToPresent lifecycle', () => {
  it('deletes the preview surface and reactivates the real sheet', async () => {
    const fake = createFakeExcel(new FakeWorkbook());
    const { workbook, run } = fake;
    workbook.addSheet('Sheet1');
    const preview = workbook.addSheet('__tl_preview_0a0a0a0a');
    preview.visibility = 'VeryHidden';
    const target = new RealSheetRenderTarget({ run });

    // returnToPresent emits a realSheet-targeted plan carrying lifecycle ops.
    await target.reconcile({
      target: 'realSheet',
      ops: [
        { op: 'deletePreviewSheet', previewSheetId: '__tl_preview_0a0a0a0a' },
        { op: 'activateSheet', sheetId: 'Sheet1' },
      ],
    });

    expect(workbook.findSheet('__tl_preview_0a0a0a0a')).toBeUndefined();
    expect(workbook.activeSheetId).toBe('Sheet1');
  });
});
