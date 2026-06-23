import { describe, expect, it } from 'vitest';
import { InMemoryStore, TimelineEngineImpl } from '@timeline/engine';
import type { CellSlab, ReconcilePlan, ValueObservation, WorkbookSnapshot } from '@timeline/engine';
import { FakeChangeSource, RecordingRenderTarget } from '../src/excel/fakes.ts';
import type { RenderTarget } from '../src/excel/seams.ts';
import { RealTimelineDataSource } from '../src/ui/real-data-source.ts';

function cell(value: unknown, valueType: 'empty' | 'number'): CellSlab {
  return {
    values: [[value]],
    formulas: [[null]],
    numberFormats: [['General']],
    valueTypes: [[valueType]],
  };
}

const SNAPSHOT: WorkbookSnapshot = {
  workbookGuid: 'wb',
  contentHash: 'h',
  sheets: [{ sheetId: 'Sheet1', slab: cell(null, 'empty') }],
};

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

/** A render target whose reconciles block until explicitly released, recording order. */
class GatedRenderTarget implements RenderTarget {
  readonly events: string[] = [];
  readonly #releases: (() => void)[] = [];

  reconcile(_plan: ReconcilePlan): Promise<void> {
    this.events.push('start');
    return new Promise<void>((resolve) => {
      this.#releases.push(() => {
        this.events.push('end');
        resolve();
      });
    });
  }

  /** True once at least one reconcile is waiting to be released. */
  get pending(): number {
    return this.#releases.length;
  }

  releaseOne(): void {
    this.#releases.shift()?.();
  }
}

describe('RealTimelineDataSource reconcile serialization', () => {
  it('applies preview reconciles strictly in order — the next waits for the previous', async () => {
    const engine = new TimelineEngineImpl();
    const store = new InMemoryStore();
    const changeSource = new FakeChangeSource();
    const previewTarget = new GatedRenderTarget();
    const source = new RealTimelineDataSource({
      engine,
      store,
      realTarget: new RecordingRenderTarget(),
      previewTarget,
      changeSource,
    });
    await source.start(engine.attach(SNAPSHOT, null));
    changeSource.emit(edit(10)); // step 0
    changeSource.emit(edit(20)); // step 1
    changeSource.emit(edit(30)); // step 2

    // Drain microtasks + macrotasks up to the gate (reconcile blocks unreleased).
    const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

    // A fast scrub: two gotos dispatched back-to-back (no await between).
    source.dispatch({ type: 'goto', ref: { branchId: 'main', stepIndex: 0 } });
    source.dispatch({ type: 'goto', ref: { branchId: 'main', stepIndex: 1 } });

    // Only the FIRST reconcile may have started; the second waits behind it.
    await tick();
    expect(previewTarget.events).toEqual(['start']);
    expect(previewTarget.pending).toBe(1);

    // Release the first → the second now starts (and only then).
    previewTarget.releaseOne();
    await tick();
    expect(previewTarget.events).toEqual(['start', 'end', 'start']);

    previewTarget.releaseOne();
    await tick();
    expect(previewTarget.events).toEqual(['start', 'end', 'start', 'end']);
  });
});
