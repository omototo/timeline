import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { CellSlab, ValueObservation, WorkbookSnapshot } from '@timeline/engine';
import { InMemoryStore, TimelineEngineImpl } from '@timeline/engine';
import { FakeChangeSource, RecordingRenderTarget } from '../src/excel/fakes.ts';
import { RealTimelineDataSource } from '../src/ui/real-data-source.ts';

function cell(value: unknown, valueType: 'empty' | 'number'): CellSlab {
  return {
    values: [[value]],
    formulas: [[null]],
    numberFormats: [['General']],
    valueTypes: [[valueType]],
  };
}

const emptySheet1: WorkbookSnapshot = {
  workbookGuid: 'wb-1',
  contentHash: 'h0',
  sheets: [{ sheetId: 'Sheet1', slab: cell(null, 'empty') }],
};

/** A single-cell edit at A1 on Sheet1 (non-echo, local). */
function edit(value: number): ValueObservation {
  return {
    kind: 'value',
    triggerSource: 'unknown',
    source: 'local',
    sheetId: 'Sheet1',
    area: [{ startRow: 0, startCol: 0, rowCount: 1, colCount: 1 }],
    after: cell(value, 'number'),
  };
}

interface Harness {
  source: RealTimelineDataSource;
  changeSource: FakeChangeSource;
  realTarget: RecordingRenderTarget;
  previewTarget: RecordingRenderTarget;
}

async function harness(): Promise<Harness> {
  const engine = new TimelineEngineImpl();
  const store = new InMemoryStore();
  const changeSource = new FakeChangeSource();
  const realTarget = new RecordingRenderTarget();
  const previewTarget = new RecordingRenderTarget();
  const source = new RealTimelineDataSource({
    engine,
    store,
    realTarget,
    previewTarget,
    changeSource,
    sheets: ['Sheet1'],
  });
  const attach = engine.attach(emptySheet1, null);
  await source.start(attach);
  return { source, changeSource, realTarget, previewTarget };
}

describe('RealTimelineDataSource', () => {
  let h: Harness;
  beforeEach(async () => {
    h = await harness();
  });

  it('turns a debounced edit into a Step in the view', () => {
    expect(h.source.getView().branches.flatMap((b) => b.steps)).toHaveLength(0);
    h.changeSource.emit(edit(42));
    const steps = h.source.getView().branches.flatMap((b) => b.steps);
    expect(steps).toHaveLength(1);
    expect(steps[0]?.kind).toBe('value');
    expect(steps[0]?.magnitude).toBe(1);
  });

  it('notifies subscribers when an edit lands', () => {
    const listener = vi.fn();
    const unsubscribe = h.source.subscribe(listener);
    h.changeSource.emit(edit(7));
    expect(listener).toHaveBeenCalled();
    unsubscribe();
    listener.mockClear();
    h.changeSource.emit(edit(8));
    expect(listener).not.toHaveBeenCalled();
  });

  it('previews a past Step on the preview surface, then restores the real sheet', async () => {
    h.changeSource.emit(edit(42));
    h.changeSource.emit(edit(99));

    h.source.dispatch({ type: 'goto', ref: { branchId: 'main', stepIndex: 0 } });
    expect(h.source.getView().head).toMatchObject({ mode: 'preview', previewStepIndex: 0 });
    await vi.waitFor(() => {
      expect(h.previewTarget.lastPlan?.target).toBe('previewSheet');
    });

    h.source.dispatch({ type: 'returnToPresent' });
    expect(h.source.getView().head.mode).toBe('present');
    await vi.waitFor(() => {
      expect(h.realTarget.lastPlan?.target).toBe('realSheet');
    });
  });

  it('returns a referentially stable view between changes (no useSyncExternalStore loop)', () => {
    // Same object identity when nothing changed...
    expect(h.source.getView()).toBe(h.source.getView());
    const before = h.source.getView();
    // ...a fresh identity after a real change (so React re-renders exactly once).
    h.changeSource.emit(edit(5));
    const after = h.source.getView();
    expect(after).not.toBe(before);
    expect(h.source.getView()).toBe(after);
  });

  it('routes rename/delete to a no-op (no engine op yet)', () => {
    h.changeSource.emit(edit(1));
    const before = h.source.getView();
    h.source.dispatch({ type: 'renameBranch', branchId: 'main', name: 'baseline' });
    h.source.dispatch({ type: 'deleteBranch', branchId: 'main' });
    expect(h.source.getView().branches.map((b) => b.id)).toEqual(before.branches.map((b) => b.id));
  });
});
