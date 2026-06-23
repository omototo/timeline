import { describe, it, expect } from 'vitest';
import { ShadowState } from '../src/index.ts';
import type {
  Area,
  CellSlab,
  CellState,
  Rect,
  ValueDelta,
  ValueObservation,
  ValueType,
} from '../src/index.ts';

/** A single-cell Rect at an absolute coordinate. */
function cellRect(row: number, col: number): Rect {
  return { startRow: row, startCol: col, rowCount: 1, colCount: 1 };
}

/** Build a CellState with sensible defaults so tests stay terse. */
function state(partial: Partial<CellState> & { value: unknown }): CellState {
  return {
    value: partial.value,
    formula: partial.formula ?? null,
    valueType: partial.valueType ?? 'string',
    numberFormat: partial.numberFormat ?? 'General',
  };
}

/**
 * Build a dense CellSlab from a grid of CellStates laid out row-major across an
 * Area's rectangles (one row of the slab per Area row, in Area order).
 */
function slabFromStates(rows: CellState[][]): CellSlab {
  return {
    values: rows.map((r) => r.map((c) => c.value)),
    formulas: rows.map((r) => r.map((c) => c.formula)),
    numberFormats: rows.map((r) => r.map((c) => c.numberFormat)),
    valueTypes: rows.map((r) => r.map((c) => c.valueType)),
  };
}

/** Build a ValueObservation over an Area with a slab. */
function obs(sheetId: string, area: Area, slab: CellSlab): ValueObservation {
  return {
    kind: 'value',
    triggerSource: 'thisLocalAddin',
    source: 'local',
    sheetId,
    area,
    after: slab,
  };
}

describe('ShadowState.read', () => {
  it('reads an unwritten cell as the canonical empty cell', () => {
    const ss = new ShadowState();
    expect(ss.read('Sheet1', 5, 7)).toEqual<CellState>({
      value: '',
      formula: null,
      valueType: 'empty',
      numberFormat: 'General',
    });
  });

  it('reads back a written cell losslessly', () => {
    const ss = new ShadowState();
    const after = state({ value: 42, valueType: 'number', numberFormat: '0.00', formula: '=6*7' });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after }],
    });
    expect(ss.read('Sheet1', 0, 0)).toEqual(after);
  });
});

describe('ShadowState.diff', () => {
  it('returns a single changed cell with before/after', () => {
    const ss = new ShadowState();
    const after = state({ value: 'hello' });
    const changed = ss.diff(obs('Sheet1', [cellRect(2, 3)], slabFromStates([[after]])));
    expect(changed).toEqual([
      {
        addr: cellRect(2, 3),
        before: { value: '', formula: null, valueType: 'empty', numberFormat: 'General' },
        after,
      },
    ]);
  });

  it('returns only the cells that actually changed (sparse)', () => {
    const ss = new ShadowState();
    const a = state({ value: 'a' });
    const b = state({ value: 'b' });
    // Seed A1=a, B1=b.
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [
        { addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after: a },
        { addr: cellRect(0, 1), before: ss.read('Sheet1', 0, 1), after: b },
      ],
    });
    // Observe A1 unchanged, B1 -> b2 over a 1x2 area.
    const b2 = state({ value: 'b2' });
    const area: Area = [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 2 }];
    const changed = ss.diff(obs('Sheet1', area, slabFromStates([[a, b2]])));
    expect(changed).toEqual([{ addr: cellRect(0, 1), before: b, after: b2 }]);
  });

  it('returns an empty list for a no-op observation', () => {
    const ss = new ShadowState();
    const v = state({ value: 7, valueType: 'number' });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(1, 1), before: ss.read('Sheet1', 1, 1), after: v }],
    });
    const changed = ss.diff(obs('Sheet1', [cellRect(1, 1)], slabFromStates([[v]])));
    expect(changed).toEqual([]);
  });

  it('walks multiple disjoint rectangles in Area order, row-major', () => {
    const ss = new ShadowState();
    const x = state({ value: 'x' });
    const y = state({ value: 'y' });
    const z = state({ value: 'z' });
    // Two disjoint rects: a 1x2 at (0,0) and a 1x1 at (5,5).
    const area: Area = [
      { startRow: 0, startCol: 0, rowCount: 1, colCount: 2 },
      { startRow: 5, startCol: 5, rowCount: 1, colCount: 1 },
    ];
    const changed = ss.diff(obs('Sheet1', area, slabFromStates([[x, y], [z]])));
    expect(changed).toEqual([
      { addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after: x },
      { addr: cellRect(0, 1), before: ss.read('Sheet1', 0, 1), after: y },
      { addr: cellRect(5, 5), before: ss.read('Sheet1', 5, 5), after: z },
    ]);
  });

  it('detects a number-format-only change as a change (lossless)', () => {
    const ss = new ShadowState();
    const v = state({ value: 1, valueType: 'number', numberFormat: 'General' });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after: v }],
    });
    const reformatted = state({ value: 1, valueType: 'number', numberFormat: '0.00%' });
    const changed = ss.diff(obs('Sheet1', [cellRect(0, 0)], slabFromStates([[reformatted]])));
    expect(changed).toEqual([{ addr: cellRect(0, 0), before: v, after: reformatted }]);
  });

  it('uses ragged-row fallbacks for a short slab row (empty defaults)', () => {
    const ss = new ShadowState();
    // Area claims 1x2 but slab row has a single cell; missing cell -> empty,
    // which equals the unwritten before, so only the present cell is a change.
    const present = state({ value: 'p' });
    const area: Area = [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 2 }];
    const slab: CellSlab = {
      values: [['p']],
      formulas: [[null]],
      numberFormats: [['General']],
      valueTypes: [['string'] as ValueType[]],
    };
    const changed = ss.diff(obs('Sheet1', area, slab));
    expect(changed).toEqual([
      { addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after: present },
    ]);
  });
});

describe('ShadowState.apply', () => {
  it('applies a delta forward, then reads the new state', () => {
    const ss = new ShadowState();
    const after = state({ value: 'v2' });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(3, 4), before: ss.read('Sheet1', 3, 4), after }],
    });
    expect(ss.read('Sheet1', 3, 4)).toEqual(after);
    expect(ss.cellCount('Sheet1')).toBe(1);
  });

  it('removes a cell that returns to empty (sparse store stays clean)', () => {
    const ss = new ShadowState();
    const filled = state({ value: 'fill' });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after: filled }],
    });
    expect(ss.cellCount('Sheet1')).toBe(1);
    const cleared: CellState = {
      value: '',
      formula: null,
      valueType: 'empty',
      numberFormat: 'General',
    };
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(0, 0), before: filled, after: cleared }],
    });
    expect(ss.cellCount('Sheet1')).toBe(0);
    expect(ss.read('Sheet1', 0, 0)).toEqual(cleared);
  });
});

describe('ShadowState sheet isolation', () => {
  it('keeps per-sheet state independent', () => {
    const ss = new ShadowState();
    const onSheet1 = state({ value: 's1' });
    const onSheet2 = state({ value: 's2' });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after: onSheet1 }],
    });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet2',
      cells: [{ addr: cellRect(0, 0), before: ss.read('Sheet2', 0, 0), after: onSheet2 }],
    });
    expect(ss.read('Sheet1', 0, 0)).toEqual(onSheet1);
    expect(ss.read('Sheet2', 0, 0)).toEqual(onSheet2);
    expect(ss.cellCount('Sheet1')).toBe(1);
    expect(ss.cellCount('Sheet2')).toBe(1);
    expect(ss.cellCount('SheetUnknown')).toBe(0);
  });

  it('diffs an observation against the right sheet only', () => {
    const ss = new ShadowState();
    const v = state({ value: 'only-on-1' });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after: v }],
    });
    // Same coordinate+value on Sheet2 is a change (Sheet2's before is empty).
    const changed = ss.diff(obs('Sheet2', [cellRect(0, 0)], slabFromStates([[v]])));
    expect(changed).toHaveLength(1);
    expect(changed[0]?.before.valueType).toBe('empty');
  });
});

describe('ShadowState.slab', () => {
  it('materializes the current mirror over an area, row-major', () => {
    const ss = new ShadowState();
    const a = state({ value: 'a' });
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [{ addr: cellRect(0, 0), before: ss.read('Sheet1', 0, 0), after: a }],
    });
    const area: Area = [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 2 }];
    const out = ss.slab('Sheet1', area);
    expect(out.values).toEqual([['a', '']]);
    expect(out.valueTypes).toEqual([['string', 'empty']]);
    expect(out.formulas).toEqual([[null, null]]);
    expect(out.numberFormats).toEqual([['General', 'General']]);
  });

  it('round-trips diff -> apply -> slab', () => {
    const ss = new ShadowState();
    const after = state({ value: 9, valueType: 'number', formula: '=3*3', numberFormat: '0' });
    const o = obs('Sheet1', [cellRect(2, 2)], slabFromStates([[after]]));
    const changed = ss.diff(o);
    const delta: ValueDelta = { kind: 'value', sheetId: 'Sheet1', cells: changed };
    ss.apply(delta);
    expect(ss.slab('Sheet1', [cellRect(2, 2)])).toEqual(slabFromStates([[after]]));
  });
});
