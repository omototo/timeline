import { describe, it, expect } from 'vitest';
import { ShadowState, applyDelta, reconstruct } from '../src/index.ts';
import type { CellState, Delta, Rect } from '../src/index.ts';

function cellRect(row: number, col: number): Rect {
  return { startRow: row, startCol: col, rowCount: 1, colCount: 1 };
}

function state(value: unknown): CellState {
  return { value, formula: null, valueType: 'string', numberFormat: 'General' };
}

describe('applyDelta — forward dispatch over every Delta kind', () => {
  it('applies a value delta', () => {
    const s = new ShadowState();
    applyDelta(s, {
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(0, 0), before: state(''), after: state('x') }],
    });
    expect(s.read('S', 0, 0).value).toBe('x');
  });

  it('applies a structural (row insert) delta as a coordinate remap', () => {
    const s = new ShadowState();
    applyDelta(s, {
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(0, 0), before: state(''), after: state('top') }],
    });
    applyDelta(s, {
      kind: 'structural',
      sheetId: 'S',
      changeType: 'rowInserted',
      address: cellRect(0, 0),
    });
    // The cell shifted down one row.
    expect(s.read('S', 1, 0).value).toBe('top');
  });

  it('applies a worksheet (add) delta to the sheet-meta map', () => {
    const s = new ShadowState();
    applyDelta(s, { kind: 'worksheet', op: 'add', sheetId: 'New', newName: 'New' });
    expect(s.sheetMeta('New')?.name).toBe('New');
  });

  it('applies a reconciliation delta: structural ops then cell writes per sheet', () => {
    const s = new ShadowState();
    // Seed a cell so the structural remap has something to move.
    applyDelta(s, {
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(0, 0), before: state(''), after: state('orig') }],
    });
    const recon: Delta = {
      kind: 'reconciliation',
      perSheet: [
        {
          sheetId: 'S',
          structural: [{ changeType: 'rowInserted', address: cellRect(0, 0) }],
          cells: [{ addr: cellRect(0, 0), before: state(''), after: state('reconciled') }],
        },
      ],
    };
    applyDelta(s, recon);
    // 'orig' shifted down to row 1; the reconciled write landed at row 0.
    expect(s.read('S', 1, 0).value).toBe('orig');
    expect(s.read('S', 0, 0).value).toBe('reconciled');
  });

  it('applies a reconciliation delta with structural ops carrying a shift direction', () => {
    const s = new ShadowState();
    applyDelta(s, {
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(0, 1), before: state(''), after: state('right') }],
    });
    applyDelta(s, {
      kind: 'reconciliation',
      perSheet: [
        {
          sheetId: 'S',
          structural: [
            { changeType: 'cellInserted', address: cellRect(0, 0), shiftDirection: 'right' },
          ],
          cells: [],
        },
      ],
    });
    // The cell at (0,1) shifted right to (0,2).
    expect(s.read('S', 0, 2).value).toBe('right');
  });
});

describe('reconstruct — seed + forward replay', () => {
  it('seeds from null (empty) and replays deltas forward', () => {
    const deltas: Delta[] = [
      {
        kind: 'value',
        sheetId: 'S',
        cells: [{ addr: cellRect(0, 0), before: state(''), after: state('a') }],
      },
      {
        kind: 'value',
        sheetId: 'S',
        cells: [{ addr: cellRect(0, 0), before: state('a'), after: state('b') }],
      },
    ];
    const s = reconstruct(null, deltas);
    expect(s.read('S', 0, 0).value).toBe('b');
  });

  it('seeds from a snapshot and replays the window forward', () => {
    const seed = new ShadowState();
    applyDelta(seed, {
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(0, 0), before: state(''), after: state('seed') }],
    });
    const snap = seed.snapshot();

    const s = reconstruct(snap, [
      {
        kind: 'value',
        sheetId: 'S',
        cells: [{ addr: cellRect(1, 0), before: state(''), after: state('next') }],
      },
    ]);
    // Both the seeded cell and the replayed cell are present.
    expect(s.read('S', 0, 0).value).toBe('seed');
    expect(s.read('S', 1, 0).value).toBe('next');
  });
});

describe('ShadowState snapshot/fromSnapshot round-trip', () => {
  it('round-trips cells and sheet metadata losslessly', () => {
    const s = new ShadowState();
    applyDelta(s, {
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(2, 3), before: state(''), after: state('keep') }],
    });
    applyDelta(s, { kind: 'worksheet', op: 'add', sheetId: 'S', newName: 'Sheet S' });

    const restored = ShadowState.fromSnapshot(s.snapshot());
    expect(restored.read('S', 2, 3).value).toBe('keep');
    expect(restored.sheetMeta('S')?.name).toBe('Sheet S');
    // It is an independent copy — mutating the source does not touch the restore.
    applyDelta(s, {
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(2, 3), before: state('keep'), after: state('changed') }],
    });
    expect(restored.read('S', 2, 3).value).toBe('keep');
  });
});
