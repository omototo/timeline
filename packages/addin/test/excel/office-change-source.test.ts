import { describe, it, expect, vi, beforeEach } from 'vitest';
import type {
  Observation,
  StructuralObservation,
  ValueObservation,
  WorksheetObservation,
} from '@timeline/engine';
import {
  OfficeChangeSource,
  type OfficeChangeSourceOptions,
} from '../../src/excel/office-change-source.ts';
import { ExpectedWriteSet } from '../../src/excel/expected-write-set.ts';
import { createFakeExcel, changedEvent, FakeWorkbook } from './fake-excel.ts';

/** A controllable timer so debounce is deterministic. */
function manualTimer() {
  let pending: (() => void) | null = null;
  return {
    setTimer: (cb: () => void) => {
      pending = cb;
      return 1 as unknown as ReturnType<typeof setTimeout>;
    },
    clearTimer: () => {
      pending = null;
    },
    flush: () => {
      const cb = pending;
      pending = null;
      cb?.();
    },
    get armed() {
      return pending !== null;
    },
  };
}

describe('OfficeChangeSource', () => {
  let workbook: FakeWorkbook;
  let run: ReturnType<typeof createFakeExcel>['run'];
  let observed: Observation[];

  beforeEach(() => {
    const fake = createFakeExcel(new FakeWorkbook());
    workbook = fake.workbook;
    run = fake.run;
    observed = [];
  });

  /**
   * Build a source that auto-attaches per-sheet handlers to every worksheet
   * currently in the workbook (the lister is invoked inside `start`).
   */
  function makeSource(opts: Partial<OfficeChangeSourceOptions> = {}) {
    return new OfficeChangeSource({ run, ...opts }).withWorksheetLister(() => [...workbook.sheets]);
  }

  it('registers worksheet-collection handlers on start', async () => {
    const sheet = workbook.addSheet('Sheet1');
    const src = makeSource();
    await src.start((o) => observed.push(o));
    expect(workbook.worksheets.onAdded.handlerCount).toBe(1);
    expect(workbook.worksheets.onDeleted.handlerCount).toBe(1);
    expect(workbook.worksheets.onNameChanged.handlerCount).toBe(1);
    expect(sheet.onChanged.handlerCount).toBe(1);
    expect(sheet.onFormatChanged.handlerCount).toBe(1);
    await src.stop();
    expect(workbook.worksheets.onAdded.handlerCount).toBe(0);
    expect(sheet.onChanged.handlerCount).toBe(0);
  });

  it('maps a single-cell value edit to a ValueObservation with a read-back slab', async () => {
    const sheet = workbook.addSheet('Sheet1');
    sheet.setCell(1, 1, { value: 42, formula: '=6*7', numberFormat: '0.00', valueType: 'Double' });
    const timer = manualTimer();
    const src = makeSource({ setTimer: timer.setTimer, clearTimer: timer.clearTimer });
    await src.start((o) => observed.push(o));

    await sheet.onChanged.fire(
      changedEvent({ worksheetId: sheet.id, address: 'B2', details: null }),
    );
    timer.flush();

    expect(observed).toHaveLength(1);
    const obs = observed[0] as ValueObservation;
    expect(obs.kind).toBe('value');
    expect(obs.sheetId).toBe(sheet.id);
    expect(obs.area).toEqual([{ startRow: 1, startCol: 1, rowCount: 1, colCount: 1 }]);
    expect(obs.after.values).toEqual([[42]]);
    expect(obs.after.formulas).toEqual([['=6*7']]);
    expect(obs.after.numberFormats).toEqual([['0.00']]);
    expect(obs.after.valueTypes).toEqual([['number']]);
    expect(obs.source).toBe('local');
  });

  it('reads back a multi-cell change into a slab matching the address', async () => {
    const sheet = workbook.addSheet('Sheet1');
    sheet.setCell(0, 0, { value: 1, valueType: 'Double' });
    sheet.setCell(0, 1, { value: 2, valueType: 'Double' });
    sheet.setCell(1, 0, { value: 3, valueType: 'Double' });
    sheet.setCell(1, 1, { value: 4, valueType: 'Double' });
    const timer = manualTimer();
    const src = makeSource({ setTimer: timer.setTimer, clearTimer: timer.clearTimer });
    await src.start((o) => observed.push(o));

    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A1:B2' }));
    timer.flush();

    const obs = observed[0] as ValueObservation;
    expect(obs.after.values).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  it('maps a structural row insert (with changeDirectionState) to a StructuralObservation', async () => {
    const sheet = workbook.addSheet('Sheet1');
    const timer = manualTimer();
    const src = makeSource({ setTimer: timer.setTimer, clearTimer: timer.clearTimer });
    await src.start((o) => observed.push(o));

    await sheet.onChanged.fire(
      changedEvent({
        worksheetId: sheet.id,
        address: 'A3:A3',
        changeType: 'RowInserted',
        changeDirectionState: { insertShiftDirection: 'Down' },
      }),
    );
    timer.flush();

    const obs = observed[0] as StructuralObservation;
    expect(obs.kind).toBe('structural');
    expect(obs.changeType).toBe('rowInserted');
    expect(obs.shiftDirection).toBe('down');
    expect(obs.address).toEqual({ startRow: 2, startCol: 0, rowCount: 1, colCount: 1 });
  });

  it('debounces a burst of events into one batch', async () => {
    const sheet = workbook.addSheet('Sheet1');
    const timer = manualTimer();
    const src = makeSource({ setTimer: timer.setTimer, clearTimer: timer.clearTimer });
    await src.start((o) => observed.push(o));

    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A1' }));
    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A2' }));
    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A3' }));
    expect(observed).toHaveLength(0); // nothing emitted until the window closes
    timer.flush();
    expect(observed).toHaveLength(3); // all three coalesced into one flush
  });

  it('drops echo events where triggerSource === ThisLocalAddin (1.14)', async () => {
    const sheet = workbook.addSheet('Sheet1');
    const timer = manualTimer();
    const src = makeSource({ setTimer: timer.setTimer, clearTimer: timer.clearTimer });
    await src.start((o) => observed.push(o));

    await sheet.onChanged.fire(
      changedEvent({ worksheetId: sheet.id, address: 'A1', triggerSource: 'ThisLocalAddin' }),
    );
    expect(timer.armed).toBe(false); // never enqueued
    timer.flush();
    expect(observed).toHaveLength(0);
  });

  it('falls back to the expected-write set for echo cancellation below 1.14', async () => {
    const sheet = workbook.addSheet('Sheet1');
    const expectedWrites = new ExpectedWriteSet({ now: () => 0 });
    expectedWrites.register(sheet.id, 'A1');
    const timer = manualTimer();
    const src = makeSource({
      isSetSupported: () => false, // pre-1.14
      expectedWrites,
      setTimer: timer.setTimer,
      clearTimer: timer.clearTimer,
    });
    await src.start((o) => observed.push(o));

    // First A1 event is our own write echo → dropped.
    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A1' }));
    timer.flush();
    expect(observed).toHaveLength(0);

    // A later genuine A1 edit is NOT swallowed (registration was consumed).
    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A1' }));
    timer.flush();
    expect(observed).toHaveLength(1);
  });

  it('chunks the read-back when the region exceeds the cell limit', async () => {
    const sheet = workbook.addSheet('Sheet1');
    for (let r = 0; r < 4; r++) {
      sheet.setCell(r, 0, { value: r, valueType: 'Double' });
    }
    const timer = manualTimer();
    const src = makeSource({
      maxCellsPerRead: 1, // force a tile per row
      setTimer: timer.setTimer,
      clearTimer: timer.clearTimer,
    });
    await src.start((o) => observed.push(o));

    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A1:A4' }));
    timer.flush();

    const obs = observed[0] as ValueObservation;
    expect(obs.after.values).toEqual([[0], [1], [2], [3]]);
    // Four single-cell tiles were untracked (one untrack per tile read).
    expect(workbook.untrackCount).toBe(4);
  });

  it('defensively sub-tiles a wide region whose single-row tile still exceeds the limit', async () => {
    const sheet = workbook.addSheet('Sheet1');
    // 2 rows x 3 cols. maxCellsPerRead = 2 → rowsPerTile = floor(2/3) clamped to
    // 1, so each whole-row tile is 3 cells > 2 → the defensive row-halving path
    // (then single-row read-as-is) is exercised.
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 3; c++) {
        sheet.setCell(r, c, { value: r * 3 + c, valueType: 'Double' });
      }
    }
    const timer = manualTimer();
    const src = makeSource({
      maxCellsPerRead: 2,
      setTimer: timer.setTimer,
      clearTimer: timer.clearTimer,
    });
    await src.start((o) => observed.push(o));

    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A1:C2' }));
    timer.flush();

    const obs = observed[0] as ValueObservation;
    expect(obs.after.values).toEqual([
      [0, 1, 2],
      [3, 4, 5],
    ]);
  });

  it('registers no onMoved handler when the host lacks the collection event', async () => {
    const sheet = workbook.addSheet('Sheet1');
    // Simulate a host without the reorder event.
    (workbook.worksheets as { onMoved?: unknown }).onMoved = undefined;
    const src = makeSource();
    await src.start((o) => observed.push(o));
    expect(sheet.onChanged.handlerCount).toBe(1);
    await src.stop();
  });

  it('maps sheet add / delete / rename / reorder to WorksheetObservations', async () => {
    const timer = manualTimer();
    const src = makeSource({ setTimer: timer.setTimer, clearTimer: timer.clearTimer });
    await src.start((o) => observed.push(o));

    await workbook.worksheets.onAdded.fire({ worksheetId: 'S2', source: 'Local' });
    await workbook.worksheets.onNameChanged.fire({
      worksheetId: 'S2',
      nameAfter: 'Renamed',
      source: 'Local',
    });
    await workbook.worksheets.onMoved.fire({
      worksheetId: 'S2',
      positionAfter: 0,
      source: 'Local',
    });
    await workbook.worksheets.onDeleted.fire({ worksheetId: 'S2', source: 'Local' });
    timer.flush();

    const ops = (observed as WorksheetObservation[]).map((o) => o.op);
    expect(ops).toEqual(['add', 'rename', 'reorder', 'delete']);
    const rename = observed[1] as WorksheetObservation;
    expect(rename.newName).toBe('Renamed');
    const reorder = observed[2] as WorksheetObservation;
    expect(reorder.newPosition).toBe(0);
  });

  it('invokes onRemoteChange for a co-authoring (source: remote) event', async () => {
    const sheet = workbook.addSheet('Sheet1');
    const onRemoteChange = vi.fn();
    const timer = manualTimer();
    const src = makeSource({
      onRemoteChange,
      setTimer: timer.setTimer,
      clearTimer: timer.clearTimer,
    });
    await src.start((o) => observed.push(o));

    await sheet.onChanged.fire(
      changedEvent({ worksheetId: sheet.id, address: 'A1', source: 'Remote' }),
    );
    expect(onRemoteChange).toHaveBeenCalledTimes(1);
    const arg = onRemoteChange.mock.calls[0]?.[0] as Observation;
    expect(arg.source).toBe('remote');
  });

  it('ignores events fired after stop()', async () => {
    const sheet = workbook.addSheet('Sheet1');
    const timer = manualTimer();
    const src = makeSource({ setTimer: timer.setTimer, clearTimer: timer.clearTimer });
    await src.start((o) => observed.push(o));
    await src.stop();
    await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A1' }));
    expect(timer.armed).toBe(false);
    expect(observed).toHaveLength(0);
  });

  it('uses the real timer path by default (no injected timer)', async () => {
    vi.useFakeTimers();
    try {
      const sheet = workbook.addSheet('Sheet1');
      const src = makeSource({ debounceMs: 50 });
      await src.start((o) => observed.push(o));
      await sheet.onChanged.fire(changedEvent({ worksheetId: sheet.id, address: 'A1' }));
      expect(observed).toHaveLength(0);
      await vi.advanceTimersByTimeAsync(50);
      expect(observed).toHaveLength(1);
      await src.stop();
    } finally {
      vi.useRealTimers();
    }
  });
});
