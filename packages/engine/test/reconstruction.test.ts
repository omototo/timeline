import { describe, it, expect, beforeEach } from 'vitest';
import { TimelineEngineImpl } from '../src/index.ts';
import type {
  Area,
  CellSlab,
  CellState,
  EffectEnvelope,
  PersistOp,
  ReconcileOp,
  Rect,
  StepRef,
  ValueObservation,
} from '../src/index.ts';

/** A single-cell Rect at an absolute coordinate. */
function cellRect(row: number, col: number): Rect {
  return { startRow: row, startCol: col, rowCount: 1, colCount: 1 };
}

function state(partial: Partial<CellState> & { value: unknown }): CellState {
  return {
    value: partial.value,
    formula: partial.formula ?? null,
    valueType: partial.valueType ?? 'string',
    numberFormat: partial.numberFormat ?? 'General',
  };
}

function slabFromStates(rows: CellState[][]): CellSlab {
  return {
    values: rows.map((r) => r.map((c) => c.value)),
    formulas: rows.map((r) => r.map((c) => c.formula)),
    numberFormats: rows.map((r) => r.map((c) => c.numberFormat)),
    valueTypes: rows.map((r) => r.map((c) => c.valueType)),
  };
}

function valueObs(sheetId: string, area: Area, rows: CellState[][]): ValueObservation {
  return {
    kind: 'value',
    triggerSource: 'thisLocalAddin',
    source: 'local',
    sheetId,
    area,
    after: slabFromStates(rows),
  };
}

/** Set a single cell on Sheet1 at (row,col) to a string value via ingest. */
function setCell(engine: TimelineEngineImpl, row: number, col: number, value: string): void {
  engine.ingest(valueObs('Sheet1', [cellRect(row, col)], [[state({ value })]]));
}

function writeKeyframeOps(env: EffectEnvelope): Extract<PersistOp, { op: 'writeKeyframe' }>[] {
  return (env.persist ?? []).filter(
    (p): p is Extract<PersistOp, { op: 'writeKeyframe' }> => p.op === 'writeKeyframe',
  );
}

function setCellsOps(env: EffectEnvelope): Extract<ReconcileOp, { op: 'setCells' }>[] {
  return (env.reconcile?.ops ?? []).filter(
    (o): o is Extract<ReconcileOp, { op: 'setCells' }> => o.op === 'setCells',
  );
}

const ref = (stepIndex: number): StepRef => ({ branchId: 'main', stepIndex });

describe('Reconstruction — forward-replay (Wave 3)', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    // Tight cadence so a keyframe boundary is easy to cross deterministically.
    engine = new TimelineEngineImpl({ keyframeStepInterval: 3, keyframeByteThreshold: 1e9 });
  });

  it('reconstructs the exact state at an arbitrary earlier step', () => {
    setCell(engine, 0, 0, 'a'); // step 0
    setCell(engine, 0, 0, 'b'); // step 1
    setCell(engine, 0, 0, 'c'); // step 2

    expect(engine.readReconstructed(ref(0), 'Sheet1', 0, 0).value).toBe('a');
    expect(engine.readReconstructed(ref(1), 'Sheet1', 0, 0).value).toBe('b');
    expect(engine.readReconstructed(ref(2), 'Sheet1', 0, 0).value).toBe('c');
  });

  it('reconstructs exact state across a keyframe boundary', () => {
    // cadence=3 -> keyframe at step 2; reconstructing step 4 seeds from it and
    // replays steps 3,4 forward.
    for (let i = 0; i < 6; i++) setCell(engine, 0, 0, `v${String(i)}`);
    expect(engine.keyframeIndices()).toEqual([2, 5]);

    // A step strictly after the keyframe (replay forward from kf@2).
    expect(engine.readReconstructed(ref(4), 'Sheet1', 0, 0).value).toBe('v4');
    // Exactly on the keyframe boundary.
    expect(engine.readReconstructed(ref(2), 'Sheet1', 0, 0).value).toBe('v2');
    // Before any keyframe (kf@2 not <= 1) -> seeds from empty, replays 0,1.
    expect(engine.readReconstructed(ref(1), 'Sheet1', 0, 0).value).toBe('v1');
  });

  it('reconstructs multi-cell state independently per cell', () => {
    setCell(engine, 0, 0, 'A1');
    setCell(engine, 1, 1, 'B2');
    setCell(engine, 0, 0, 'A1-x');

    const s = ref(1);
    expect(engine.readReconstructed(s, 'Sheet1', 0, 0).value).toBe('A1');
    expect(engine.readReconstructed(s, 'Sheet1', 1, 1).value).toBe('B2');
    // At the tip both are present with the latest A1.
    expect(engine.readReconstructed(ref(2), 'Sheet1', 0, 0).value).toBe('A1-x');
    expect(engine.readReconstructed(ref(2), 'Sheet1', 1, 1).value).toBe('B2');
  });
});

describe('Adaptive keyframe cadence (Wave 3)', () => {
  it('fires on the step-count trigger', () => {
    const engine = new TimelineEngineImpl({
      keyframeStepInterval: 4,
      keyframeByteThreshold: 1e9,
    });
    // Count how many envelopes carried a writeKeyframe, and at which steps.
    const keyframeSteps: number[] = [];
    for (let i = 0; i < 8; i++) {
      const env = engine.ingest(
        valueObs('Sheet1', [cellRect(i, 0)], [[state({ value: `v${String(i)}` })]]),
      );
      if (writeKeyframeOps(env).length > 0) keyframeSteps.push(i);
    }
    // Keyframes at steps 3 and 7 (every 4 steps); the writeKeyframe op rides the
    // same envelope as the triggering Step.
    expect(keyframeSteps).toEqual([3, 7]);
    expect(engine.keyframeIndices()).toEqual([3, 7]);
  });

  it('fires on the byte-threshold trigger before the step count', () => {
    // A small step budget would never fire by step count within 3 steps, but a
    // tiny byte threshold trips on the very first (large) delta.
    const engine = new TimelineEngineImpl({
      keyframeStepInterval: 1000,
      keyframeByteThreshold: 50,
    });
    const big = 'x'.repeat(200);
    const env = engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: big })]]));
    expect(engine.keyframeIndices()).toEqual([0]);
    expect(writeKeyframeOps(env)).toHaveLength(1);
  });

  it('carries the serialized snapshot in the writeKeyframe op', () => {
    const engine = new TimelineEngineImpl({ keyframeStepInterval: 1, keyframeByteThreshold: 1e9 });
    const env = engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 'hi' })]]));
    const kf = writeKeyframeOps(env)[0];
    expect(kf?.branchId).toBe('main');
    expect(kf?.stepIndex).toBe(0);
    // State is a structurally-cloneable snapshot, not a Map.
    const snap = kf?.state as { sheets: unknown[]; sheetMeta: unknown[] };
    expect(Array.isArray(snap.sheets)).toBe(true);
  });

  it('resets both counters after a keyframe (no immediate re-fire)', () => {
    const engine = new TimelineEngineImpl({ keyframeStepInterval: 2, keyframeByteThreshold: 1e9 });
    for (let i = 0; i < 5; i++) setCell(engine, i, 0, `v${String(i)}`);
    // Keyframes at 1 and 3 (every 2 steps); step 4 has not re-fired yet.
    expect(engine.keyframeIndices()).toEqual([1, 3]);
  });
});

describe('goto — minimal value-mode preview diff (Wave 3)', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    setCell(engine, 0, 0, 'a'); // step 0
    setCell(engine, 1, 0, 'b'); // step 1
    setCell(engine, 0, 0, 'a2'); // step 2
  });

  it('creates + activates the preview sheet on first entry, in value mode', () => {
    const env = engine.goto(ref(2));
    expect(env.reconcile?.target).toBe('previewSheet');
    const ops = env.reconcile?.ops ?? [];
    // Surface created first; the user is anchored to it (activate emitted once
    // surfaces settle, so it lands on the final preview sheet).
    expect(ops[0]).toEqual({ op: 'createPreviewSheet', previewSheetId: '__preview__::Sheet1' });
    expect(ops).toContainEqual({ op: 'activateSheet', sheetId: '__preview__::Sheet1' });
    // First entry flags the shell to hide the real sheets (full-workbook rollback).
    expect(env.reconcile?.enterPreview).toBe(true);
    // Every cell op is value (Frozen Values, ADR-0008).
    for (const op of setCellsOps(env)) expect(op.mode).toBe('value');
    // HEAD flipped to preview.
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'preview', previewStepIndex: 2 });
  });

  it('writes the full target on first entry (from empty projection)', () => {
    const env = engine.goto(ref(2));
    const cells = setCellsOps(env);
    // At step 2: A1='a2', A2='b' -> two cells written.
    expect(cells).toHaveLength(2);
    const written = cells.map((c) => ({ rc: c.area[0], v: c.slab.values[0]?.[0] }));
    expect(written).toContainEqual({ rc: cellRect(0, 0), v: 'a2' });
    expect(written).toContainEqual({ rc: cellRect(1, 0), v: 'b' });
  });

  it('a subsequent goto writes only the MINIMAL diff vs the projected state', () => {
    engine.goto(ref(2)); // project A1='a2', A2='b'
    const env = engine.goto(ref(0)); // target: A1='a', A2 cleared
    // No createPreviewSheet on a re-entry.
    expect(env.reconcile?.ops.some((o) => o.op === 'createPreviewSheet')).toBe(false);

    const cells = setCellsOps(env);
    // Minimal: A1 'a2'->'a' (write), A2 'b'->empty (clear). Exactly two ops.
    expect(cells).toHaveLength(2);
    const a1 = cells.find((c) => c.area[0]?.startRow === 0);
    const a2 = cells.find((c) => c.area[0]?.startRow === 1);
    expect(a1?.slab.values[0]?.[0]).toBe('a');
    // A2 cleared to the empty cell.
    expect(a2?.slab.values[0]?.[0]).toBe('');
    expect(a2?.slab.valueTypes[0]?.[0]).toBe('empty');
  });

  it('goto to an identical projected state writes no cell ops', () => {
    engine.goto(ref(2));
    const env = engine.goto(ref(2)); // same target
    expect(setCellsOps(env)).toHaveLength(0);
  });

  it('multi-sheet: projects each logical sheet onto its OWN surface (no collisions)', () => {
    // Two populated sheets with COLLIDING coordinates: both write A1 (0,0).
    const e = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    e.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 's1a1' })]])); // step 0
    e.ingest(valueObs('Sheet2', [cellRect(0, 0)], [[state({ value: 's2a1' })]])); // step 1

    const env = e.goto(ref(1)); // target: both sheets populated at A1
    const ops = env.reconcile?.ops ?? [];

    // One createPreviewSheet per logical sheet, on distinct surfaces.
    const created = ops
      .filter(
        (o): o is Extract<ReconcileOp, { op: 'createPreviewSheet' }> =>
          o.op === 'createPreviewSheet',
      )
      .map((o) => o.previewSheetId);
    expect(created).toEqual(['__preview__::Sheet1', '__preview__::Sheet2']);
    // Exactly one activateSheet, for the first surface.
    const activated = ops.filter(
      (o): o is Extract<ReconcileOp, { op: 'activateSheet' }> => o.op === 'activateSheet',
    );
    expect(activated).toEqual([{ op: 'activateSheet', sheetId: '__preview__::Sheet1' }]);

    // Both A1 cells are written — to DISTINCT preview surfaces (no overwrite).
    const cells = setCellsOps(env);
    expect(cells).toHaveLength(2);
    const bySurface = new Map(cells.map((c) => [c.sheetId, c.slab.values[0]?.[0]]));
    expect(bySurface.get('__preview__::Sheet1')).toBe('s1a1');
    expect(bySurface.get('__preview__::Sheet2')).toBe('s2a1');
    // The two ops share the same (0,0) coordinate but never collide.
    const coords = cells.map((c) => {
      const rect = c.area[0];
      expect(rect).toBeDefined();
      return { row: rect?.startRow, col: rect?.startCol };
    });
    expect(coords).toEqual([
      { row: 0, col: 0 },
      { row: 0, col: 0 },
    ]);
    expect(new Set(cells.map((c) => c.sheetId)).size).toBe(2);
  });

  it('multi-sheet: a later goto only re-creates surfaces it has not seen yet', () => {
    const e = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    e.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 's1' })]])); // step 0
    e.ingest(valueObs('Sheet2', [cellRect(0, 0)], [[state({ value: 's2' })]])); // step 1

    e.goto(ref(0)); // only Sheet1 populated at step 0 -> creates Sheet1 surface
    const env = e.goto(ref(1)); // now Sheet2 also populated -> create Sheet2 only
    const created = (env.reconcile?.ops ?? [])
      .filter(
        (o): o is Extract<ReconcileOp, { op: 'createPreviewSheet' }> =>
          o.op === 'createPreviewSheet',
      )
      .map((o) => o.previewSheetId);
    expect(created).toEqual(['__preview__::Sheet2']);
    // A scrub re-anchors the user to their sheet's preview (Sheet1 here), and is
    // NOT a first entry, so it does not re-flag the shell to hide real sheets.
    expect(env.reconcile?.ops).toContainEqual({
      op: 'activateSheet',
      sheetId: '__preview__::Sheet1',
    });
    expect(env.reconcile?.enterPreview ?? false).toBe(false);
  });

  it('multi-sheet: scrubbing before a sheet existed deletes its preview surface', () => {
    const e = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    e.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 's1' })]])); // step 0
    e.ingest(valueObs('Sheet2', [cellRect(0, 0)], [[state({ value: 's2' })]])); // step 1

    e.goto(ref(1)); // surfaces: Sheet1 + Sheet2
    const env = e.goto(ref(0)); // Sheet2 did not exist at step 0 -> drop its surface
    const deleted = (env.reconcile?.ops ?? [])
      .filter(
        (o): o is Extract<ReconcileOp, { op: 'deletePreviewSheet' }> =>
          o.op === 'deletePreviewSheet',
      )
      .map((o) => o.previewSheetId);
    expect(deleted).toEqual(['__preview__::Sheet2']);
    // No setCells targets the now-deleted Sheet2 surface.
    expect(setCellsOps(env).every((o) => o.sheetId !== '__preview__::Sheet2')).toBe(true);
  });

  it('orders same-row cell ops left-to-right (column tie-break)', () => {
    const e = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    // Two cells on the same row, written in the order C1 then B1.
    e.ingest(valueObs('Sheet1', [cellRect(0, 2)], [[state({ value: 'c' })]])); // step 0
    e.ingest(valueObs('Sheet1', [cellRect(0, 1)], [[state({ value: 'b' })]])); // step 1

    const env = e.goto(ref(1));
    const cols = setCellsOps(env).map((o) => o.area[0]?.startCol);
    // Row-major within the row -> ascending columns regardless of write order.
    expect(cols).toEqual([1, 2]);
  });
});

describe('returnToPresent (Wave 3)', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
    setCell(engine, 0, 0, 'a');
  });

  it('deletes the preview sheet, reactivates the real sheet, HEAD -> present', () => {
    engine.goto(ref(0));
    const env = engine.returnToPresent();
    expect(env.reconcile?.target).toBe('realSheet');
    // Deletes the per-sheet preview surface and reactivates the REAL worksheet
    // that was active when Preview began (Sheet1) — NOT the branch id 'main'.
    expect(env.reconcile?.ops).toEqual([
      { op: 'deletePreviewSheet', previewSheetId: '__preview__::Sheet1' },
      { op: 'activateSheet', sheetId: 'Sheet1' },
    ]);
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'present' });
  });

  it('never reactivates a BranchId as a SheetId', () => {
    engine.goto(ref(0));
    const env = engine.returnToPresent();
    const activate = (env.reconcile?.ops ?? []).find(
      (o): o is Extract<ReconcileOp, { op: 'activateSheet' }> => o.op === 'activateSheet',
    );
    // The reactivated id is the REAL worksheet (Sheet1), never the branch 'main'.
    expect(activate?.sheetId).toBe('Sheet1');
    expect(activate?.sheetId).not.toBe('main');
  });

  it('multi-sheet: tears down every preview surface and reactivates the real sheet', () => {
    const e = new TimelineEngineImpl();
    e.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 's1' })]]));
    e.ingest(valueObs('Sheet2', [cellRect(0, 0)], [[state({ value: 's2' })]]));
    e.goto(ref(1)); // projects both sheets onto their own surfaces

    const env = e.returnToPresent();
    expect(env.reconcile?.target).toBe('realSheet');
    expect(env.reconcile?.ops).toEqual([
      { op: 'deletePreviewSheet', previewSheetId: '__preview__::Sheet1' },
      { op: 'deletePreviewSheet', previewSheetId: '__preview__::Sheet2' },
      { op: 'activateSheet', sheetId: 'Sheet1' },
    ]);
  });

  it('reactivates the registered sheet in tab order when sheet metadata exists', () => {
    const e = new TimelineEngineImpl();
    // Register two sheets (tab order Alpha, Beta) then populate the second.
    e.ingest({
      kind: 'worksheet',
      op: 'add',
      sheetId: 'Alpha',
      triggerSource: 'thisLocalAddin',
      source: 'local',
    });
    e.ingest({
      kind: 'worksheet',
      op: 'add',
      sheetId: 'Beta',
      triggerSource: 'thisLocalAddin',
      source: 'local',
    });
    e.ingest(valueObs('Beta', [cellRect(0, 0)], [[state({ value: 'x' })]])); // step 2

    e.goto(ref(2));
    const env = e.returnToPresent();
    const activate = (env.reconcile?.ops ?? []).find(
      (o): o is Extract<ReconcileOp, { op: 'activateSheet' }> => o.op === 'activateSheet',
    );
    // Registered tab order wins: Alpha is first, even though Beta holds the cells.
    expect(activate?.sheetId).toBe('Alpha');
  });

  it('omits activateSheet when no real sheet is knowable', () => {
    const e = new TimelineEngineImpl();
    // A structural op on an untouched sheet: records a Step but populates and
    // registers nothing, so there is no real sheet to reactivate.
    e.ingest({
      kind: 'structural',
      sheetId: 'Ghost',
      changeType: 'rowInserted',
      address: cellRect(0, 0),
      triggerSource: 'thisLocalAddin',
      source: 'local',
    });
    e.goto(ref(0));
    const env = e.returnToPresent();
    // No surfaces created (nothing populated) and no activateSheet op.
    expect(env.reconcile?.ops).toEqual([]);
    expect(engine.head()).toBeDefined();
  });

  it('is a no-op when not in preview', () => {
    const env = engine.returnToPresent();
    expect(env).toEqual({});
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'present' });
  });

  it('ingest after returnToPresent records normally again', () => {
    engine.goto(ref(0));
    engine.returnToPresent();
    const env = engine.ingest(valueObs('Sheet1', [cellRect(2, 2)], [[state({ value: 'z' })]]));
    expect(env.persist?.some((p) => p.op === 'appendDelta')).toBe(true);
    expect(engine.readShadow('Sheet1', 2, 2).value).toBe('z');
  });
});

describe('leaving preview via branch / switch tears down the rollback (ADR-0014)', () => {
  function inPreviewOnMain(): TimelineEngineImpl {
    const e = new TimelineEngineImpl({ keyframeStepInterval: 1000, keyframeByteThreshold: 1e9 });
    e.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 'a' })]])); // step 0
    e.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 'b' })]])); // step 1
    e.goto(ref(0)); // enter preview -> hides reals, creates Sheet1 surface
    return e;
  }

  it('branch from preview deletes surfaces, flags exitPreview, lands in Present', () => {
    const e = inPreviewOnMain();
    const env = e.branch(ref(0));
    expect(env.reconcile?.target).toBe('realSheet');
    expect(env.reconcile?.exitPreview).toBe(true);
    expect(env.reconcile?.ops).toContainEqual({
      op: 'deletePreviewSheet',
      previewSheetId: '__preview__::Sheet1',
    });
    expect(e.head().mode).toBe('present');
  });

  it('switch from preview deletes surfaces and flags exitPreview', () => {
    const e = inPreviewOnMain();
    e.branch(ref(0)); // -> on provisional branch-1, Present
    e.ingest(valueObs('Sheet1', [cellRect(1, 0)], [[state({ value: 'fork' })]])); // persist branch-1
    e.goto(ref(0)); // re-enter preview
    const env = e.switch('main');
    expect(env.reconcile?.target).toBe('realSheet');
    expect(env.reconcile?.exitPreview).toBe(true);
    expect(env.reconcile?.ops.some((o) => o.op === 'deletePreviewSheet')).toBe(true);
    expect(e.head()).toEqual({ branchId: 'main', mode: 'present' });
  });

  it('branch from Present (not preview) emits no exitPreview', () => {
    const e = new TimelineEngineImpl();
    e.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 'a' })]]));
    const env = e.branch(ref(0));
    expect(env.reconcile?.exitPreview ?? false).toBe(false);
  });
});
