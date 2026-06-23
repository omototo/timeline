/**
 * Wave 2 — structural (coordinate remap) and worksheet (sheet-map) paths.
 *
 * Asserts: insert/delete shift coordinates correctly with NO false value
 * deltas; worksheet add/delete/rename/reorder; sheet isolation. Drives the
 * engine's `ingest` for the `structural` and `worksheet` kinds and the
 * ShadowState remap directly.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { ShadowState, TimelineEngineImpl } from '../src/index.ts';
import type {
  CellState,
  EffectEnvelope,
  PersistOp,
  Rect,
  ShiftDirection,
  StructuralChangeType,
  StructuralDelta,
  StructuralObservation,
  WorksheetDelta,
  WorksheetObservation,
  WorksheetOp,
} from '../src/index.ts';

function cellRect(row: number, col: number): Rect {
  return { startRow: row, startCol: col, rowCount: 1, colCount: 1 };
}

function rect(startRow: number, startCol: number, rowCount: number, colCount: number): Rect {
  return { startRow, startCol, rowCount, colCount };
}

function state(value: unknown): CellState {
  return { value, formula: null, valueType: 'string', numberFormat: 'General' };
}

const EMPTY: CellState = { value: '', formula: null, valueType: 'empty', numberFormat: 'General' };

/** Seed a single cell into the Shadow State via a value observation. */
function seed(engine: TimelineEngineImpl, sheet: string, row: number, col: number, value: unknown) {
  engine.ingest({
    kind: 'value',
    triggerSource: 'thisLocalAddin',
    source: 'local',
    sheetId: sheet,
    area: [cellRect(row, col)],
    after: {
      values: [[value]],
      formulas: [[null]],
      numberFormats: [['General']],
      valueTypes: [['string']],
    },
  });
}

function structuralObs(
  sheetId: string,
  changeType: StructuralChangeType,
  address: Rect,
  shiftDirection?: ShiftDirection,
): StructuralObservation {
  return {
    kind: 'structural',
    triggerSource: 'thisLocalAddin',
    source: 'local',
    sheetId,
    changeType,
    address,
    ...(shiftDirection !== undefined ? { shiftDirection } : {}),
  };
}

function worksheetObs(
  op: WorksheetOp,
  sheetId: string,
  extra?: { newName?: string; newPosition?: number },
): WorksheetObservation {
  return {
    kind: 'worksheet',
    triggerSource: 'thisLocalAddin',
    source: 'local',
    op,
    sheetId,
    ...(extra?.newName !== undefined ? { newName: extra.newName } : {}),
    ...(extra?.newPosition !== undefined ? { newPosition: extra.newPosition } : {}),
  };
}

function appendOp(env: EffectEnvelope): Extract<PersistOp, { op: 'appendDelta' }> | undefined {
  return env.persist?.find(
    (p): p is Extract<PersistOp, { op: 'appendDelta' }> => p.op === 'appendDelta',
  );
}

// ---------------------------------------------------------------------------
// ShadowState.applyStructural — row/column inserts & deletes
// ---------------------------------------------------------------------------

describe('ShadowState.applyStructural — rows', () => {
  let ss: ShadowState;

  beforeEach(() => {
    ss = new ShadowState();
    // A1='a' (row 0), A3='c' (row 2) on Sheet1.
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [
        { addr: cellRect(0, 0), before: EMPTY, after: state('a') },
        { addr: cellRect(2, 0), before: EMPTY, after: state('c') },
      ],
    });
  });

  it('rowInserted opens blank space and shifts rows at/below down', () => {
    const delta: StructuralDelta = {
      kind: 'structural',
      sheetId: 'Sheet1',
      changeType: 'rowInserted',
      address: rect(1, 0, 1, 1),
    };
    ss.applyStructural(delta);

    expect(ss.read('Sheet1', 0, 0)).toEqual(state('a')); // above insert: unmoved
    expect(ss.read('Sheet1', 1, 0)).toEqual(EMPTY); // newly-opened blank row
    expect(ss.read('Sheet1', 3, 0)).toEqual(state('c')); // shifted down by 1
    expect(ss.cellCount('Sheet1')).toBe(2); // no false value delta
  });

  it('rowDeleted removes the spanned row and shifts rows below up', () => {
    const delta: StructuralDelta = {
      kind: 'structural',
      sheetId: 'Sheet1',
      changeType: 'rowDeleted',
      address: rect(0, 0, 1, 1),
    };
    ss.applyStructural(delta);

    expect(ss.read('Sheet1', 1, 0)).toEqual(state('c')); // row 2 shifted up by 1 to row 1
    expect(ss.read('Sheet1', 0, 0)).toEqual(EMPTY); // 'a' was deleted with row 0
    expect(ss.read('Sheet1', 2, 0)).toEqual(EMPTY); // old position now empty
    expect(ss.cellCount('Sheet1')).toBe(1); // 'a' was deleted with its row
  });

  it('multi-row insert shifts by rowCount', () => {
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'Sheet1',
      changeType: 'rowInserted',
      address: rect(0, 0, 3, 1),
    });
    expect(ss.read('Sheet1', 3, 0)).toEqual(state('a'));
    expect(ss.read('Sheet1', 5, 0)).toEqual(state('c'));
  });
});

describe('ShadowState.applyStructural — columns', () => {
  let ss: ShadowState;

  beforeEach(() => {
    ss = new ShadowState();
    // A1='a' (col 0), C1='c' (col 2).
    ss.apply({
      kind: 'value',
      sheetId: 'Sheet1',
      cells: [
        { addr: cellRect(0, 0), before: EMPTY, after: state('a') },
        { addr: cellRect(0, 2), before: EMPTY, after: state('c') },
      ],
    });
  });

  it('columnInserted shifts columns at/right of the insert right', () => {
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'Sheet1',
      changeType: 'columnInserted',
      address: rect(0, 1, 1, 1),
    });
    expect(ss.read('Sheet1', 0, 0)).toEqual(state('a'));
    expect(ss.read('Sheet1', 0, 1)).toEqual(EMPTY);
    expect(ss.read('Sheet1', 0, 3)).toEqual(state('c'));
  });

  it('columnDeleted removes the spanned column and shifts right cols left', () => {
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'Sheet1',
      changeType: 'columnDeleted',
      address: rect(0, 0, 1, 1),
    });
    expect(ss.read('Sheet1', 0, 1)).toEqual(state('c')); // col 2 shifted left to col 1
    expect(ss.read('Sheet1', 0, 0)).toEqual(EMPTY); // 'a' deleted with col 0
    expect(ss.cellCount('Sheet1')).toBe(1); // only 'c' remains
  });

  it('does nothing on an untouched sheet', () => {
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'Untouched',
      changeType: 'rowInserted',
      address: rect(0, 0, 1, 1),
    });
    expect(ss.cellCount('Untouched')).toBe(0);
  });
});

describe('ShadowState.applyStructural — cell insert/delete with shiftDirection', () => {
  let ss: ShadowState;

  beforeEach(() => {
    ss = new ShadowState();
    // A grid: B2='b2'(1,1), B3='b3'(2,1), C2='c2'(1,2)
    ss.apply({
      kind: 'value',
      sheetId: 'S',
      cells: [
        { addr: cellRect(1, 1), before: EMPTY, after: state('b2') },
        { addr: cellRect(2, 1), before: EMPTY, after: state('b3') },
        { addr: cellRect(1, 2), before: EMPTY, after: state('c2') },
      ],
    });
  });

  it('cellInserted shift down moves only the affected columns', () => {
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'S',
      changeType: 'cellInserted',
      address: rect(1, 1, 1, 1),
      shiftDirection: 'down',
    });
    // Column 1 cells at/below row 1 shift down; column 2 untouched.
    expect(ss.read('S', 2, 1)).toEqual(state('b2'));
    expect(ss.read('S', 3, 1)).toEqual(state('b3'));
    expect(ss.read('S', 1, 1)).toEqual(EMPTY);
    expect(ss.read('S', 1, 2)).toEqual(state('c2')); // unaffected column
  });

  it('cellInserted shift right moves only the affected rows', () => {
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'S',
      changeType: 'cellInserted',
      address: rect(1, 1, 1, 1),
      shiftDirection: 'right',
    });
    expect(ss.read('S', 1, 2)).toEqual(state('b2')); // B2 -> C2
    expect(ss.read('S', 1, 1)).toEqual(EMPTY);
    expect(ss.read('S', 2, 1)).toEqual(state('b3')); // row 2 unaffected
  });

  it('cellDeleted shift up removes the cell and shifts the column up', () => {
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'S',
      changeType: 'cellDeleted',
      address: rect(1, 1, 1, 1),
      shiftDirection: 'up',
    });
    expect(ss.read('S', 1, 1)).toEqual(state('b3')); // b3 shifted up
    expect(ss.read('S', 2, 1)).toEqual(EMPTY);
    expect(ss.read('S', 1, 2)).toEqual(state('c2')); // unaffected column
  });

  it('cellDeleted shift up leaves cells above the deletion in place', () => {
    // A2='a2'(1,0) sits above a deletion at (2,0); it must not move.
    ss.apply({
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(1, 0), before: EMPTY, after: state('a2') }],
    });
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'S',
      changeType: 'cellDeleted',
      address: rect(2, 0, 1, 1),
      shiftDirection: 'up',
    });
    expect(ss.read('S', 1, 0)).toEqual(state('a2')); // above deletion: unmoved
  });

  it('cellDeleted shift left leaves cells left of the deletion in place', () => {
    // A2='a2'(1,0) sits left of a deletion at (1,2); it must not move.
    ss.apply({
      kind: 'value',
      sheetId: 'S',
      cells: [{ addr: cellRect(1, 0), before: EMPTY, after: state('a2') }],
    });
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'S',
      changeType: 'cellDeleted',
      address: rect(1, 2, 1, 1),
      shiftDirection: 'left',
    });
    expect(ss.read('S', 1, 0)).toEqual(state('a2')); // left of deletion: unmoved
  });

  it('cellDeleted shift left removes the cell and shifts the row left', () => {
    ss.applyStructural({
      kind: 'structural',
      sheetId: 'S',
      changeType: 'cellDeleted',
      address: rect(1, 1, 1, 1),
      shiftDirection: 'left',
    });
    expect(ss.read('S', 1, 1)).toEqual(state('c2')); // c2 shifted left into B2
    expect(ss.read('S', 1, 2)).toEqual(EMPTY);
    expect(ss.read('S', 2, 1)).toEqual(state('b3')); // unaffected row
  });
});

// ---------------------------------------------------------------------------
// ShadowState.applyWorksheet — sheet map
// ---------------------------------------------------------------------------

describe('ShadowState.applyWorksheet', () => {
  let ss: ShadowState;

  beforeEach(() => {
    ss = new ShadowState();
  });

  function ws(
    op: WorksheetOp,
    sheetId: string,
    extra?: { newName?: string; newPosition?: number },
  ): WorksheetDelta {
    return {
      kind: 'worksheet',
      op,
      sheetId,
      ...(extra?.newName !== undefined ? { newName: extra.newName } : {}),
      ...(extra?.newPosition !== undefined ? { newPosition: extra.newPosition } : {}),
    };
  }

  it('add registers a sheet with a name and order', () => {
    ss.applyWorksheet(ws('add', 'S1', { newName: 'Sheet One' }));
    ss.applyWorksheet(ws('add', 'S2', { newName: 'Sheet Two' }));
    expect(ss.sheetMeta('S1')).toEqual({ sheetId: 'S1', name: 'Sheet One', order: 0 });
    expect(ss.sheetMeta('S2')).toEqual({ sheetId: 'S2', name: 'Sheet Two', order: 1 });
    expect(ss.sheets().map((s) => s.sheetId)).toEqual(['S1', 'S2']);
  });

  it('delete drops the sheet metadata and its cells, repacking order', () => {
    ss.applyWorksheet(ws('add', 'S1'));
    ss.applyWorksheet(ws('add', 'S2'));
    ss.apply({
      kind: 'value',
      sheetId: 'S1',
      cells: [{ addr: cellRect(0, 0), before: EMPTY, after: state('x') }],
    });
    ss.applyWorksheet(ws('delete', 'S1'));
    expect(ss.sheetMeta('S1')).toBeUndefined();
    expect(ss.cellCount('S1')).toBe(0);
    expect(ss.sheetMeta('S2')?.order).toBe(0); // repacked to front
  });

  it('rename changes the name but keeps the stable id and order', () => {
    ss.applyWorksheet(ws('add', 'S1', { newName: 'Old' }));
    ss.applyWorksheet(ws('rename', 'S1', { newName: 'New' }));
    expect(ss.sheetMeta('S1')).toEqual({ sheetId: 'S1', name: 'New', order: 0 });
  });

  it('reorder moves a tab to newPosition and repacks order densely', () => {
    ss.applyWorksheet(ws('add', 'A'));
    ss.applyWorksheet(ws('add', 'B'));
    ss.applyWorksheet(ws('add', 'C'));
    // Move C to the front.
    ss.applyWorksheet(ws('reorder', 'C', { newPosition: 0 }));
    expect(ss.sheets().map((s) => s.sheetId)).toEqual(['C', 'A', 'B']);
    expect(ss.sheets().map((s) => s.order)).toEqual([0, 1, 2]);
  });

  it('reorder past the end clamps to the last position', () => {
    ss.applyWorksheet(ws('add', 'A'));
    ss.applyWorksheet(ws('add', 'B'));
    ss.applyWorksheet(ws('reorder', 'A', { newPosition: 99 }));
    expect(ss.sheets().map((s) => s.sheetId)).toEqual(['B', 'A']);
  });

  it('reorder of an unknown sheet is a no-op', () => {
    ss.applyWorksheet(ws('add', 'A'));
    ss.applyWorksheet(ws('reorder', 'Z', { newPosition: 0 }));
    expect(ss.sheets().map((s) => s.sheetId)).toEqual(['A']);
  });

  it('add honours an explicit newPosition', () => {
    ss.applyWorksheet(ws('add', 'A'));
    ss.applyWorksheet(ws('add', 'B', { newPosition: 0 }));
    expect(ss.sheets().map((s) => s.sheetId)).toEqual(['B', 'A']);
  });

  it('rename of an unknown sheet creates it with the given name', () => {
    ss.applyWorksheet(ws('rename', 'New', { newName: 'Fresh' }));
    expect(ss.sheetMeta('New')).toEqual({ sheetId: 'New', name: 'Fresh', order: 0 });
  });
});

// ---------------------------------------------------------------------------
// Engine.ingest — structural & worksheet recording paths
// ---------------------------------------------------------------------------

describe('TimelineEngineImpl.ingest — structural path', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('records a StructuralDelta Step and remaps the Shadow State, no value delta', () => {
    seed(engine, 'Sheet1', 2, 0, 'c'); // A3='c'
    const env = engine.ingest(structuralObs('Sheet1', 'rowInserted', rect(0, 0, 1, 1)));

    const delta = appendOp(env)?.delta as StructuralDelta;
    expect(delta.kind).toBe('structural');
    expect(delta.changeType).toBe('rowInserted');
    expect(delta.address).toEqual(rect(0, 0, 1, 1));

    // Coordinate remapped: 'c' shifted from row 2 to row 3.
    expect(engine.readShadow('Sheet1', 3, 0)).toEqual(state('c'));
    expect(engine.readShadow('Sheet1', 2, 0)).toEqual(EMPTY);

    // Two steps total (seed value + structural); both on main.
    expect(engine.tipStepIndex()).toBe(1);
    expect(env.reconcile).toBeUndefined();
  });

  it('carries shiftDirection through into the delta', () => {
    const env = engine.ingest(structuralObs('S', 'cellInserted', rect(1, 1, 1, 1), 'right'));
    const delta = appendOp(env)?.delta as StructuralDelta;
    expect(delta.shiftDirection).toBe('right');
  });

  it('omits shiftDirection when the observation has none', () => {
    const env = engine.ingest(structuralObs('S', 'rowInserted', rect(0, 0, 1, 1)));
    const delta = appendOp(env)?.delta as StructuralDelta;
    expect('shiftDirection' in delta).toBe(false);
  });

  it('records a structural Step even when it moves no populated cell', () => {
    const env = engine.ingest(structuralObs('Empty', 'rowInserted', rect(0, 0, 1, 1)));
    expect(env.persist).toHaveLength(2);
    expect(engine.steps()).toHaveLength(1);
  });

  it('isolates structural remaps per sheet', () => {
    seed(engine, 'Sheet1', 0, 0, 's1');
    seed(engine, 'Sheet2', 0, 0, 's2');
    engine.ingest(structuralObs('Sheet1', 'rowInserted', rect(0, 0, 1, 1)));

    expect(engine.readShadow('Sheet1', 1, 0)).toEqual(state('s1')); // shifted
    expect(engine.readShadow('Sheet2', 0, 0)).toEqual(state('s2')); // untouched
  });
});

describe('TimelineEngineImpl.ingest — worksheet path', () => {
  let engine: TimelineEngineImpl;

  beforeEach(() => {
    engine = new TimelineEngineImpl();
  });

  it('records add/delete/rename/reorder as worksheet Steps', () => {
    const addEnv = engine.ingest(worksheetObs('add', 'S1', { newName: 'One' }));
    const addDelta = appendOp(addEnv)?.delta as WorksheetDelta;
    expect(addDelta.kind).toBe('worksheet');
    expect(addDelta.op).toBe('add');
    expect(engine.sheetMeta('S1')).toEqual({ sheetId: 'S1', name: 'One', order: 0 });

    engine.ingest(worksheetObs('add', 'S2', { newName: 'Two' }));
    engine.ingest(worksheetObs('rename', 'S1', { newName: 'Renamed' }));
    expect(engine.sheetMeta('S1')?.name).toBe('Renamed');

    engine.ingest(worksheetObs('reorder', 'S2', { newPosition: 0 }));
    expect(engine.shadowSheets().map((s) => s.sheetId)).toEqual(['S2', 'S1']);

    engine.ingest(worksheetObs('delete', 'S1'));
    expect(engine.sheetMeta('S1')).toBeUndefined();

    // Five worksheet steps recorded.
    expect(engine.steps()).toHaveLength(5);
    expect(engine.steps().every((s) => s.delta.kind === 'worksheet')).toBe(true);
  });

  it('omits optional fields when the observation has none', () => {
    const env = engine.ingest(worksheetObs('add', 'Plain'));
    const delta = appendOp(env)?.delta as WorksheetDelta;
    expect('newName' in delta).toBe(false);
    expect('newPosition' in delta).toBe(false);
  });
});
