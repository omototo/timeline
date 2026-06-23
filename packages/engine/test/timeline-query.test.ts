/**
 * Wave 5 — the query surface: `timeline()` (histogram model — step magnitudes +
 * branch splits) and `inspectStep()` (per-cell formula metadata). Pure: drives
 * engine verbs and asserts on the returned views (ADR-0004, ADR-0013).
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { TimelineEngineImpl, stepMagnitude, stepFormulaCells } from '../src/index.ts';
import type {
  Area,
  CellSlab,
  CellState,
  Delta,
  Rect,
  StructuralObservation,
  ValueObservation,
  WorksheetObservation,
} from '../src/index.ts';

// --- builders ---------------------------------------------------------------

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

function setCell(
  engine: TimelineEngineImpl,
  sheet: string,
  row: number,
  col: number,
  value: string,
) {
  return engine.ingest(valueObs(sheet, [cellRect(row, col)], [[state({ value })]]));
}

/** A wide single-row paste of `width` cells starting at (0,0). */
function pasteRow(engine: TimelineEngineImpl, sheet: string, width: number) {
  const cells = Array.from({ length: width }, (_, c) => state({ value: `v${String(c)}` }));
  return engine.ingest(
    valueObs(sheet, [{ startRow: 0, startCol: 0, rowCount: 1, colCount: width }], [cells]),
  );
}

// ---------------------------------------------------------------------------

describe('TimelineEngineImpl.timeline — histogram magnitudes', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('is empty before any Step', () => {
    const view = engine.timeline();
    expect(view.steps).toEqual([]);
    expect(view.branches).toEqual([]);
  });

  it('includes a main BranchMeta after a main-only Step so step refs resolve', () => {
    // Regression (Wave 5 review): a single ingest on main must yield a coherent
    // histogram — every Step ref must resolve within view.branches. Before the
    // fix, view.steps pointed at branchId 'main' while view.branches was [].
    setCell(engine, 'Sheet1', 0, 0, 'hi');
    const view = engine.timeline();

    const main = view.branches.find((b) => b.id === 'main');
    expect(main).toBeDefined();
    expect(main).toEqual({ id: 'main', order: 0, provisional: false });

    // Every Step ref resolves to a branch in the returned graph.
    const branchIds = new Set(view.branches.map((b) => b.id));
    for (const step of view.steps) {
      expect(branchIds.has(step.ref.branchId)).toBe(true);
    }
  });

  it('does NOT persist main as a saved BranchMeta on its first Step', () => {
    // The implicit main root is registered resident but never emits a saveBranch
    // (ADR-0005: "main is never a saved BranchMeta"). The appendDelta + setHead
    // are the only persist ops for a main-only edit.
    const env = setCell(engine, 'Sheet1', 0, 0, 'hi');
    const ops = env.persist ?? [];
    expect(ops.some((o) => o.op === 'saveBranch')).toBe(false);
    expect(ops.map((o) => o.op)).toEqual(['appendDelta', 'setHead']);
    // main is resident even though it was never persisted.
    expect(engine.hasBranch('main')).toBe(true);
  });

  it("a single-cell edit's magnitude is 1", () => {
    setCell(engine, 'Sheet1', 0, 0, 'hi');
    const view = engine.timeline();
    expect(view.steps).toHaveLength(1);
    expect(view.steps[0]).toEqual({
      ref: { branchId: 'main', stepIndex: 0 },
      kind: 'value',
      magnitude: 1,
    });
  });

  it("a wide paste's magnitude equals its changed-cell count (towers over edits)", () => {
    pasteRow(engine, 'Sheet1', 250);
    setCell(engine, 'Sheet1', 5, 0, 'edit');
    const view = engine.timeline();
    expect(view.steps).toHaveLength(2);
    expect(view.steps[0]?.magnitude).toBe(250);
    expect(view.steps[1]?.magnitude).toBe(1);
    // The paste bar dwarfs the single-cell edit bar.
    expect(view.steps[0]?.magnitude).toBeGreaterThan((view.steps[1]?.magnitude ?? 0) * 100);
  });

  it('a structural Step has a small fixed magnitude (1), not a tall bar', () => {
    pasteRow(engine, 'Sheet1', 100);
    const structural: StructuralObservation = {
      kind: 'structural',
      triggerSource: 'thisLocalAddin',
      source: 'local',
      sheetId: 'Sheet1',
      changeType: 'rowInserted',
      address: cellRect(0, 0),
    };
    engine.ingest(structural);
    const view = engine.timeline();
    expect(view.steps[1]).toEqual({
      ref: { branchId: 'main', stepIndex: 1 },
      kind: 'structural',
      magnitude: 1,
    });
  });

  it('a worksheet Step has a small fixed magnitude (1)', () => {
    const worksheet: WorksheetObservation = {
      kind: 'worksheet',
      triggerSource: 'thisLocalAddin',
      source: 'local',
      op: 'add',
      sheetId: 'Sheet2',
      newName: 'Budget',
    };
    engine.ingest(worksheet);
    const view = engine.timeline();
    expect(view.steps[0]).toEqual({
      ref: { branchId: 'main', stepIndex: 0 },
      kind: 'worksheet',
      magnitude: 1,
    });
  });

  it('orders Steps by stepIndex within a branch', () => {
    setCell(engine, 'Sheet1', 0, 0, 'a');
    setCell(engine, 'Sheet1', 1, 0, 'b');
    setCell(engine, 'Sheet1', 2, 0, 'c');
    const view = engine.timeline();
    expect(view.steps.map((s) => s.ref.stepIndex)).toEqual([0, 1, 2]);
  });

  it('a reconciliation Step magnitude is its total repaired-cell count', () => {
    // Seed the mirror, then attach with drifted content -> reconciliation Step.
    setCell(engine, 'Sheet1', 0, 0, 'orig');
    const drifted: CellSlab = slabFromStates([
      [state({ value: 'changed' }), state({ value: 'new' })],
    ]);
    engine.attach(
      {
        workbookGuid: 'g',
        contentHash: 'different-hash',
        sheets: [{ sheetId: 'Sheet1', slab: drifted }],
      },
      { head: { branchId: 'main', mode: 'present' }, tipHash: 'old-hash' },
    );
    const view = engine.timeline();
    const recon = view.steps.find((s) => s.kind === 'reconciliation');
    expect(recon).toBeDefined();
    // A1 changed orig->changed, B1 empty->new: 2 cells.
    expect(recon?.magnitude).toBe(2);
  });
});

describe('TimelineEngineImpl.timeline — branch splits', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('exposes the fork graph (parent + forkedAt) once a branch is created', () => {
    setCell(engine, 'Sheet1', 0, 0, 'a');
    setCell(engine, 'Sheet1', 1, 0, 'b');
    // Fork from main#0.
    engine.branch({ branchId: 'main', stepIndex: 0 });
    setCell(engine, 'Sheet1', 2, 0, 'c'); // first Step on the fork (persists it)

    const view = engine.timeline();
    const fork = view.branches.find((b) => b.id === 'branch-1');
    expect(fork).toBeDefined();
    expect(fork?.parentBranchId).toBe('main');
    expect(fork?.forkedAt).toEqual({ branchId: 'main', stepIndex: 0 });

    // main is present in the graph too (the fork recorded it as a parent).
    expect(view.branches.some((b) => b.id === 'main')).toBe(true);
  });

  it('reports Steps across both the parent and the forked branch', () => {
    setCell(engine, 'Sheet1', 0, 0, 'a'); // main#0
    engine.branch({ branchId: 'main', stepIndex: 0 });
    setCell(engine, 'Sheet1', 1, 0, 'fork0'); // branch-1#0
    setCell(engine, 'Sheet1', 2, 0, 'fork1'); // branch-1#1

    const view = engine.timeline();
    const byBranch = new Map<string, number[]>();
    for (const s of view.steps) {
      const list = byBranch.get(s.ref.branchId) ?? [];
      list.push(s.ref.stepIndex);
      byBranch.set(s.ref.branchId, list);
    }
    expect(byBranch.get('main')).toEqual([0]);
    expect(byBranch.get('branch-1')).toEqual([0, 1]);
  });

  it('branches in the view are tab-order ascending', () => {
    setCell(engine, 'Sheet1', 0, 0, 'a');
    engine.branch({ branchId: 'main', stepIndex: 0 });
    setCell(engine, 'Sheet1', 1, 0, 'b');
    engine.switch('main');
    engine.branch({ branchId: 'main', stepIndex: 0 });
    setCell(engine, 'Sheet1', 2, 0, 'c');

    const view = engine.timeline();
    const orders = view.branches.map((b) => b.order);
    expect(orders).toEqual([...orders].sort((x, y) => x - y));
  });
});

describe('TimelineEngineImpl.timeline — query filters', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
    setCell(engine, 'Sheet1', 0, 0, 'a'); // main#0
    engine.branch({ branchId: 'main', stepIndex: 0 });
    setCell(engine, 'Sheet1', 1, 0, 'f0'); // branch-1#0
    setCell(engine, 'Sheet1', 2, 0, 'f1'); // branch-1#1
    setCell(engine, 'Sheet1', 3, 0, 'f2'); // branch-1#2
  });

  it('scopes Steps to a single branch but keeps the full branch graph', () => {
    const view = engine.timeline({ branchId: 'branch-1' });
    expect(view.steps.every((s) => s.ref.branchId === 'branch-1')).toBe(true);
    expect(view.steps).toHaveLength(3);
    // The graph still includes main (the renderer needs the parent).
    expect(view.branches.some((b) => b.id === 'main')).toBe(true);
  });

  it('clips Steps to an inclusive [fromStepIndex, toStepIndex] window', () => {
    const view = engine.timeline({ branchId: 'branch-1', fromStepIndex: 1, toStepIndex: 2 });
    expect(view.steps.map((s) => s.ref.stepIndex)).toEqual([1, 2]);
  });

  it('fromStepIndex alone clips the lower bound', () => {
    const view = engine.timeline({ branchId: 'branch-1', fromStepIndex: 2 });
    expect(view.steps.map((s) => s.ref.stepIndex)).toEqual([2]);
  });

  it('toStepIndex alone clips the upper bound', () => {
    const view = engine.timeline({ branchId: 'branch-1', toStepIndex: 0 });
    expect(view.steps.map((s) => s.ref.stepIndex)).toEqual([0]);
  });
});

describe('TimelineEngineImpl.inspectStep — formula metadata', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('reports before/after formula text for a value Step', () => {
    // First write a literal, then overwrite with a formula so before != after.
    engine.ingest(
      valueObs('Sheet1', [cellRect(0, 0)], [[state({ value: 1, valueType: 'number' })]]),
    );
    engine.ingest(
      valueObs(
        'Sheet1',
        [cellRect(0, 0)],
        [[state({ value: 3, valueType: 'number', formula: '=1+2' })]],
      ),
    );

    const detail = engine.inspectStep({ branchId: 'main', stepIndex: 1 });
    expect(detail.kind).toBe('value');
    expect(detail.cells).toEqual([
      { addr: cellRect(0, 0), beforeFormula: null, afterFormula: '=1+2' },
    ]);
  });

  it('reports formula metadata for every cell a multi-cell Step touched', () => {
    const area: Area = [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 2 }];
    engine.ingest(
      valueObs('Sheet1', area, [
        [
          state({ value: 2, valueType: 'number', formula: '=A1' }),
          state({ value: 5, valueType: 'number', formula: '=SUM(A1:B1)' }),
        ],
      ]),
    );
    const detail = engine.inspectStep({ branchId: 'main', stepIndex: 0 });
    expect(detail.cells).toEqual([
      { addr: cellRect(0, 0), beforeFormula: null, afterFormula: '=A1' },
      { addr: cellRect(0, 1), beforeFormula: null, afterFormula: '=SUM(A1:B1)' },
    ]);
  });

  it('a structural Step has empty formula metadata (no formula text changes)', () => {
    const structural: StructuralObservation = {
      kind: 'structural',
      triggerSource: 'thisLocalAddin',
      source: 'local',
      sheetId: 'Sheet1',
      changeType: 'rowInserted',
      address: cellRect(0, 0),
    };
    engine.ingest(structural);
    const detail = engine.inspectStep({ branchId: 'main', stepIndex: 0 });
    expect(detail.kind).toBe('structural');
    expect(detail.cells).toEqual([]);
  });

  it('a worksheet Step has empty formula metadata', () => {
    const worksheet: WorksheetObservation = {
      kind: 'worksheet',
      triggerSource: 'thisLocalAddin',
      source: 'local',
      op: 'rename',
      sheetId: 'Sheet1',
      newName: 'Renamed',
    };
    engine.ingest(worksheet);
    const detail = engine.inspectStep({ branchId: 'main', stepIndex: 0 });
    expect(detail.kind).toBe('worksheet');
    expect(detail.cells).toEqual([]);
  });

  it('reports formula metadata for a reconciliation Step', () => {
    engine.ingest(
      valueObs(
        'Sheet1',
        [cellRect(0, 0)],
        [[state({ value: 1, valueType: 'number', formula: '=ORIG' })]],
      ),
    );
    const drifted: CellSlab = slabFromStates([
      [state({ value: 9, valueType: 'number', formula: '=DRIFTED' })],
    ]);
    engine.attach(
      {
        workbookGuid: 'g',
        contentHash: 'drift',
        sheets: [{ sheetId: 'Sheet1', slab: drifted }],
      },
      { head: { branchId: 'main', mode: 'present' }, tipHash: 'stale' },
    );
    const reconRef = { branchId: 'main', stepIndex: 1 };
    const detail = engine.inspectStep(reconRef);
    expect(detail.kind).toBe('reconciliation');
    expect(detail.cells).toEqual([
      { addr: cellRect(0, 0), beforeFormula: '=ORIG', afterFormula: '=DRIFTED' },
    ]);
  });

  it('throws a RangeError for a ref that is not a recorded Step', () => {
    setCell(engine, 'Sheet1', 0, 0, 'a');
    expect(() => engine.inspectStep({ branchId: 'main', stepIndex: 99 })).toThrow(RangeError);
    expect(() => engine.inspectStep({ branchId: 'no-such-branch', stepIndex: 0 })).toThrow(
      RangeError,
    );
  });
});

describe('stepMagnitude / stepFormulaCells — exported pure helpers', () => {
  it('stepMagnitude covers every Delta kind', () => {
    const value: Delta = {
      kind: 'value',
      sheetId: 'S',
      cells: [
        { addr: cellRect(0, 0), before: state({ value: '' }), after: state({ value: 'x' }) },
        { addr: cellRect(0, 1), before: state({ value: '' }), after: state({ value: 'y' }) },
      ],
    };
    const structural: Delta = {
      kind: 'structural',
      sheetId: 'S',
      changeType: 'rowInserted',
      address: cellRect(0, 0),
    };
    const worksheet: Delta = { kind: 'worksheet', op: 'add', sheetId: 'S2' };
    const reconciliation: Delta = {
      kind: 'reconciliation',
      perSheet: [
        {
          sheetId: 'S',
          cells: [
            { addr: cellRect(0, 0), before: state({ value: '' }), after: state({ value: 'a' }) },
          ],
          structural: [],
        },
        {
          sheetId: 'T',
          cells: [
            { addr: cellRect(0, 0), before: state({ value: '' }), after: state({ value: 'b' }) },
            { addr: cellRect(1, 0), before: state({ value: '' }), after: state({ value: 'c' }) },
          ],
          structural: [],
        },
      ],
    };

    expect(stepMagnitude(value)).toBe(2);
    expect(stepMagnitude(structural)).toBe(1);
    expect(stepMagnitude(worksheet)).toBe(1);
    expect(stepMagnitude(reconciliation)).toBe(3);
  });

  it('stepFormulaCells flattens reconciliation per-sheet cells', () => {
    const reconciliation: Delta = {
      kind: 'reconciliation',
      perSheet: [
        {
          sheetId: 'S',
          cells: [
            {
              addr: cellRect(0, 0),
              before: state({ value: '', formula: '=OLD' }),
              after: state({ value: 'a', formula: '=NEW' }),
            },
          ],
          structural: [],
        },
      ],
    };
    expect(stepFormulaCells(reconciliation)).toEqual([
      { addr: cellRect(0, 0), beforeFormula: '=OLD', afterFormula: '=NEW' },
    ]);
  });
});
