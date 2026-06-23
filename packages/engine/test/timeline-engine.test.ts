import { describe, it, expect, beforeEach } from 'vitest';
import { TimelineEngineImpl } from '../src/index.ts';
import type {
  Area,
  CellSlab,
  CellState,
  EffectEnvelope,
  PersistOp,
  Rect,
  StructuralObservation,
  ValueDelta,
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

/** Pull the appendDelta op out of an envelope's persist list. */
function appendOp(env: EffectEnvelope): Extract<PersistOp, { op: 'appendDelta' }> | undefined {
  return env.persist?.find(
    (p): p is Extract<PersistOp, { op: 'appendDelta' }> => p.op === 'appendDelta',
  );
}

function setHeadOp(env: EffectEnvelope): Extract<PersistOp, { op: 'setHead' }> | undefined {
  return env.persist?.find((p): p is Extract<PersistOp, { op: 'setHead' }> => p.op === 'setHead');
}

describe('TimelineEngineImpl.ingest — value path (Present)', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('starts on main in present mode with no steps', () => {
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'present' });
    expect(engine.tipStepIndex()).toBe(-1);
    expect(engine.steps()).toEqual([]);
    expect(engine.lastDiagnostic()).toBeNull();
  });

  it('records a single-cell edit: Step, Shadow State, advanced HEAD, PersistOps', () => {
    const after = state({ value: 'hi' });
    const env = engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[after]]));

    // Shadow State updated.
    expect(engine.readShadow('Sheet1', 0, 0)).toEqual(after);
    expect(engine.shadowCellCount('Sheet1')).toBe(1);

    // A Step was appended and the tip advanced to 0.
    expect(engine.tipStepIndex()).toBe(0);
    expect(engine.steps()).toHaveLength(1);
    expect(engine.steps()[0]?.stepIndex).toBe(0);

    // Returned PersistOps: appendDelta + setHead, no reconcile.
    expect(env.reconcile).toBeUndefined();
    expect(env.persist).toHaveLength(2);
    const append = appendOp(env);
    expect(append?.branchId).toBe('main');
    const delta = append?.delta as ValueDelta;
    expect(delta.kind).toBe('value');
    expect(delta.sheetId).toBe('Sheet1');
    expect(delta.cells).toEqual([
      {
        addr: cellRect(0, 0),
        before: { value: '', formula: null, valueType: 'empty', numberFormat: 'General' },
        after,
      },
    ]);
    expect(setHeadOp(env)?.head).toEqual({ branchId: 'main', mode: 'present' });
  });

  it('records correct before/after across two sequential edits to one cell', () => {
    const v1 = state({ value: 'first' });
    engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[v1]]));

    const v2 = state({ value: 'second' });
    const env = engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[v2]]));

    const delta = appendOp(env)?.delta as ValueDelta;
    expect(delta.cells[0]?.before).toEqual(v1);
    expect(delta.cells[0]?.after).toEqual(v2);
    expect(engine.readShadow('Sheet1', 0, 0)).toEqual(v2);
    expect(engine.tipStepIndex()).toBe(1);
    expect(engine.steps()).toHaveLength(2);
  });

  it('records a multi-cell edit in a single Step', () => {
    const a = state({ value: 'a' });
    const b = state({ value: 'b' });
    const area: Area = [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 2 }];
    const env = engine.ingest(valueObs('Sheet1', area, [[a, b]]));

    const delta = appendOp(env)?.delta as ValueDelta;
    expect(delta.cells).toHaveLength(2);
    expect(engine.shadowCellCount('Sheet1')).toBe(2);
    // Still one Step.
    expect(engine.steps()).toHaveLength(1);
    expect(engine.tipStepIndex()).toBe(0);
  });

  it('records a multi-area edit (disjoint rectangles) in a single Step', () => {
    const x = state({ value: 'x' });
    const y = state({ value: 'y' });
    const z = state({ value: 'z' });
    const area: Area = [
      { startRow: 0, startCol: 0, rowCount: 1, colCount: 2 },
      { startRow: 9, startCol: 9, rowCount: 1, colCount: 1 },
    ];
    const env = engine.ingest(valueObs('Sheet1', area, [[x, y], [z]]));

    const delta = appendOp(env)?.delta as ValueDelta;
    expect(delta.cells.map((c) => c.addr)).toEqual([
      cellRect(0, 0),
      cellRect(0, 1),
      cellRect(9, 9),
    ]);
    expect(engine.shadowCellCount('Sheet1')).toBe(3);
    expect(engine.steps()).toHaveLength(1);
  });

  it('treats a no-op observation as no Step and an empty envelope', () => {
    const v = state({ value: 7, valueType: 'number' });
    engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[v]]));
    expect(engine.tipStepIndex()).toBe(0);

    // Re-observe the identical value: nothing changed.
    const env = engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[v]]));
    expect(env).toEqual({});
    expect(engine.tipStepIndex()).toBe(0);
    expect(engine.steps()).toHaveLength(1);
  });

  it('records only the changed cell within a partly-unchanged multi-cell area', () => {
    const a = state({ value: 'a' });
    const b = state({ value: 'b' });
    const area: Area = [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 2 }];
    engine.ingest(valueObs('Sheet1', area, [[a, b]]));

    // A1 unchanged, B1 -> b2.
    const b2 = state({ value: 'b2' });
    const env = engine.ingest(valueObs('Sheet1', area, [[a, b2]]));
    const delta = appendOp(env)?.delta as ValueDelta;
    expect(delta.cells).toEqual([{ addr: cellRect(0, 1), before: b, after: b2 }]);
  });

  it('isolates edits per sheet', () => {
    const v1 = state({ value: 's1' });
    const v2 = state({ value: 's2' });
    engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[v1]]));
    engine.ingest(valueObs('Sheet2', [cellRect(0, 0)], [[v2]]));

    expect(engine.readShadow('Sheet1', 0, 0)).toEqual(v1);
    expect(engine.readShadow('Sheet2', 0, 0)).toEqual(v2);
    expect(engine.shadowCellCount('Sheet1')).toBe(1);
    expect(engine.shadowCellCount('Sheet2')).toBe(1);
    // Both edits live on the same branch log.
    expect(engine.steps()).toHaveLength(2);
    expect(engine.tipStepIndex()).toBe(1);
  });

  it('advances HEAD.stepIndex monotonically across many edits', () => {
    for (let i = 0; i < 5; i++) {
      engine.ingest(valueObs('Sheet1', [cellRect(i, 0)], [[state({ value: `v${String(i)}` })]]));
    }
    expect(engine.tipStepIndex()).toBe(4);
    expect(engine.steps().map((s) => s.stepIndex)).toEqual([0, 1, 2, 3, 4]);
  });
});

describe('TimelineEngineImpl.ingest — refusals & diagnostics', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('refuses a non-value Observation kind with a diagnostic (no Step)', () => {
    const structural: StructuralObservation = {
      kind: 'structural',
      triggerSource: 'thisLocalAddin',
      source: 'local',
      sheetId: 'Sheet1',
      changeType: 'rowInserted',
      address: cellRect(0, 0),
    };
    const env = engine.ingest(structural);
    expect(env).toEqual({});
    expect(engine.steps()).toHaveLength(0);
    expect(engine.lastDiagnostic()?.code).toBe('unsupportedKind');
  });

  it('clears the last diagnostic on a subsequent successful ingest', () => {
    engine.ingest({
      kind: 'structural',
      triggerSource: 'thisLocalAddin',
      source: 'local',
      sheetId: 'Sheet1',
      changeType: 'rowInserted',
      address: cellRect(0, 0),
    });
    expect(engine.lastDiagnostic()).not.toBeNull();

    engine.ingest(valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 'ok' })]]));
    expect(engine.lastDiagnostic()).toBeNull();
  });
});

describe('TimelineEngineImpl — Wave-1-out-of-scope methods throw', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('timeline / inspectStep / lifecycle / navigation are not implemented', () => {
    const ref = { branchId: 'main', stepIndex: 0 };
    expect(() => engine.timeline()).toThrow();
    expect(() => engine.inspectStep(ref)).toThrow();
    expect(() =>
      engine.attach({ workbookGuid: 'g', contentHash: 'h', sheets: [] }, null),
    ).toThrow();
    expect(() => engine.detachToCoauthoring()).toThrow();
    expect(() => engine.goto(ref)).toThrow();
    expect(() => engine.returnToPresent()).toThrow();
    expect(() => engine.branch(ref)).toThrow();
    expect(() => engine.switch('main')).toThrow();
  });
});
