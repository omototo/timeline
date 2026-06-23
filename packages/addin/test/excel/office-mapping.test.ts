import { describe, it, expect } from 'vitest';
import {
  parseAddress,
  toStructuralChangeType,
  toShiftDirection,
  toInsertShift,
  toDeleteShift,
  toExcelSheetName,
  toValueType,
  type ParsedRect,
} from '../../src/excel/office-mapping.ts';
import { previewSheetIdFor } from '@timeline/engine';
import { ExpectedWriteSet } from '../../src/excel/expected-write-set.ts';

describe('toExcelSheetName', () => {
  it('passes a real sheet id through unchanged', () => {
    expect(toExcelSheetName('{6F9619FF-8B86-D011-B42D-00CF4FC964FF}')).toBe(
      '{6F9619FF-8B86-D011-B42D-00CF4FC964FF}',
    );
  });

  it('folds an engine preview id into a legal, <=31-char Excel name', () => {
    const guid = '{6F9619FF-8B86-D011-B42D-00CF4FC964FF}';
    const name = toExcelSheetName(previewSheetIdFor(guid));
    expect(name.length).toBeLessThanOrEqual(31);
    expect(name).not.toMatch(/[\\/?*[\]:]/);
  });

  it('is deterministic and distinct per logical sheet', () => {
    const a = toExcelSheetName(previewSheetIdFor('Sheet1'));
    const b = toExcelSheetName(previewSheetIdFor('Sheet1'));
    const c = toExcelSheetName(previewSheetIdFor('Sheet2'));
    expect(a).toBe(b);
    expect(a).not.toBe(c);
  });
});

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

  it('parses a normal range (B2:D4)', () => {
    expect(parseAddress('B2:D4')).toEqual({ startRow: 1, startCol: 1, rowCount: 3, colCount: 3 });
  });

  it('strips absolute-reference $ anchors ($B$2)', () => {
    expect(parseAddress('$B$2')).toEqual({ startRow: 1, startCol: 1, rowCount: 1, colCount: 1 });
    expect(parseAddress('$B$2:$D$4')).toEqual({
      startRow: 1,
      startCol: 1,
      rowCount: 3,
      colCount: 3,
    });
  });

  it('strips a sheet qualifier on a single cell (Sheet1!A1)', () => {
    expect(parseAddress('Sheet1!A1')).toEqual({
      startRow: 0,
      startCol: 0,
      rowCount: 1,
      colCount: 1,
    });
  });

  it('parses a whole-row range (3:3) spanning every column', () => {
    expect(parseAddress('3:3')).toEqual({
      startRow: 2,
      startCol: 0,
      rowCount: 1,
      colCount: 16_384,
    });
    expect(parseAddress('5:7')).toEqual({
      startRow: 4,
      startCol: 0,
      rowCount: 3,
      colCount: 16_384,
    });
  });

  it('parses a whole-column range (C:C) spanning every row', () => {
    expect(parseAddress('C:C')).toEqual({
      startRow: 0,
      startCol: 2,
      rowCount: 1_048_576,
      colCount: 1,
    });
    expect(parseAddress('B:D')).toEqual({
      startRow: 0,
      startCol: 1,
      rowCount: 1_048_576,
      colCount: 3,
    });
  });

  it('does not throw inside an awaited onChanged-style handler for structural forms', async () => {
    // parseAddress runs inside an awaited onChanged handler; the structural
    // forms Excel emits must resolve, never reject.
    const handler = async (address: string): Promise<ParsedRect> => {
      await Promise.resolve();
      return parseAddress(address);
    };
    for (const addr of ['3:3', 'C:C', '$B$2', 'Sheet1!A1', 'B2:D4']) {
      await expect(handler(addr)).resolves.toBeDefined();
    }
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
