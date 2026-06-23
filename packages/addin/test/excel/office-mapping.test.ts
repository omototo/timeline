import { describe, it, expect } from 'vitest';
import {
  parseAddress,
  toStructuralChangeType,
  toShiftDirection,
  toInsertShift,
  toDeleteShift,
  toValueType,
} from '../../src/excel/office-mapping.ts';
import { ExpectedWriteSet } from '../../src/excel/expected-write-set.ts';

describe('parseAddress', () => {
  it('parses a single cell', () => {
    expect(parseAddress('B2')).toEqual({ startRow: 1, startCol: 1, rowCount: 1, colCount: 1 });
  });

  it('parses a range', () => {
    expect(parseAddress('B2:D5')).toEqual({ startRow: 1, startCol: 1, rowCount: 4, colCount: 3 });
  });

  it('strips a sheet qualifier', () => {
    expect(parseAddress('Sheet1!A1:A2')).toEqual({
      startRow: 0,
      startCol: 0,
      rowCount: 2,
      colCount: 1,
    });
  });

  it('parses multi-letter columns', () => {
    expect(parseAddress('AA1')).toEqual({ startRow: 0, startCol: 26, rowCount: 1, colCount: 1 });
  });

  it('normalizes a reversed range', () => {
    expect(parseAddress('D5:B2')).toEqual({ startRow: 1, startCol: 1, rowCount: 4, colCount: 3 });
  });

  it('throws on a malformed address', () => {
    expect(() => parseAddress('not-a-cell')).toThrow('cannot parse');
    expect(() => parseAddress('A1:zzz')).toThrow('cannot parse');
  });
});

describe('toStructuralChangeType', () => {
  it('maps structural change types', () => {
    expect(toStructuralChangeType('RowInserted')).toBe('rowInserted');
    expect(toStructuralChangeType('RowDeleted')).toBe('rowDeleted');
    expect(toStructuralChangeType('ColumnInserted')).toBe('columnInserted');
    expect(toStructuralChangeType('ColumnDeleted')).toBe('columnDeleted');
    expect(toStructuralChangeType('CellInserted')).toBe('cellInserted');
    expect(toStructuralChangeType('CellDeleted')).toBe('cellDeleted');
  });

  it('returns null for value edits', () => {
    expect(toStructuralChangeType('RangeEdited')).toBeNull();
    expect(toStructuralChangeType('Unknown')).toBeNull();
  });
});

describe('toShiftDirection', () => {
  it('maps insert/delete shift state', () => {
    expect(toShiftDirection({ insertShiftDirection: 'Down' })).toBe('down');
    expect(toShiftDirection({ insertShiftDirection: 'Right' })).toBe('right');
    expect(toShiftDirection({ deleteShiftDirection: 'Up' })).toBe('up');
    expect(toShiftDirection({ deleteShiftDirection: 'Left' })).toBe('left');
  });

  it('returns undefined when absent or empty', () => {
    expect(toShiftDirection(undefined)).toBeUndefined();
    expect(toShiftDirection({})).toBeUndefined();
  });
});

describe('shift mapping back to Office enums', () => {
  it('toInsertShift', () => {
    expect(toInsertShift('right')).toBe('Right');
    expect(toInsertShift('down')).toBe('Down');
    expect(toInsertShift(undefined)).toBe('Down');
  });

  it('toDeleteShift', () => {
    expect(toDeleteShift('left')).toBe('Left');
    expect(toDeleteShift('up')).toBe('Up');
    expect(toDeleteShift(undefined)).toBe('Up');
  });
});

describe('toValueType', () => {
  it('maps known Excel value types', () => {
    expect(toValueType('Empty')).toBe('empty');
    expect(toValueType('String')).toBe('string');
    expect(toValueType('Double')).toBe('number');
    expect(toValueType('Integer')).toBe('number');
    expect(toValueType('Boolean')).toBe('boolean');
    expect(toValueType('Error')).toBe('error');
  });

  it('flattens rich/entity types to richValue', () => {
    expect(toValueType('Entity')).toBe('richValue');
    expect(toValueType('LinkedEntity')).toBe('richValue');
  });
});

describe('ExpectedWriteSet', () => {
  it('registers and consumes a write within the window', () => {
    const set = new ExpectedWriteSet({ windowMs: 100, now: () => 0 });
    set.register('S1', 'A1');
    expect(set.size).toBe(1);
    expect(set.consume('S1', 'a1')).toBe(true); // case-insensitive
    expect(set.size).toBe(0);
  });

  it('reports a miss for an unregistered region', () => {
    const set = new ExpectedWriteSet();
    expect(set.consume('S1', 'A1')).toBe(false);
  });

  it('treats an expired registration as not matching', () => {
    let clock = 0;
    const set = new ExpectedWriteSet({ windowMs: 100, now: () => clock });
    set.register('S1', 'A1');
    clock = 200;
    expect(set.consume('S1', 'A1')).toBe(false);
  });

  it('prunes expired registrations', () => {
    let clock = 0;
    const set = new ExpectedWriteSet({ windowMs: 100, now: () => clock });
    set.register('S1', 'A1');
    set.register('S1', 'B2');
    clock = 200;
    set.prune();
    expect(set.size).toBe(0);
  });

  it('defaults the window and clock when unconfigured', () => {
    const set = new ExpectedWriteSet();
    set.register('S1', 'A1');
    expect(set.consume('S1', 'A1')).toBe(true);
  });
});
