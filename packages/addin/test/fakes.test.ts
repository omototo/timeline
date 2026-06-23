import { describe, it, expect } from 'vitest';
import { FakeChangeSource, RecordingRenderTarget } from '../src/excel/fakes.ts';
import type { Observation, ReconcilePlan } from '@timeline/engine';

function valueObservation(n: number): Observation {
  return {
    kind: 'value',
    triggerSource: 'unknown',
    source: 'local',
    sheetId: 'Sheet1',
    area: [{ startRow: n, startCol: 0, rowCount: 1, colCount: 1 }],
    after: {
      values: [[n]],
      formulas: [[null]],
      numberFormats: [['General']],
      valueTypes: [['number']],
    },
  };
}

function plan(target: ReconcilePlan['target']): ReconcilePlan {
  return { target, ops: [] };
}

describe('FakeChangeSource', () => {
  it('is not started before start()', () => {
    const source = new FakeChangeSource();
    expect(source.started).toBe(false);
  });

  it('delivers emitted observations to the registered handler', async () => {
    const source = new FakeChangeSource();
    const received: Observation[] = [];
    await source.start((obs) => {
      received.push(obs);
    });
    expect(source.started).toBe(true);
    source.emit(valueObservation(0));
    source.emit(valueObservation(1));
    expect(received).toEqual([valueObservation(0), valueObservation(1)]);
  });

  it('throws when emit() is called before start()', () => {
    const source = new FakeChangeSource();
    expect(() => {
      source.emit(valueObservation(0));
    }).toThrow('before start()');
  });

  it('stops delivering after stop()', async () => {
    const source = new FakeChangeSource();
    await source.start(() => {
      throw new Error('handler should not fire after stop');
    });
    await source.stop();
    expect(source.started).toBe(false);
    expect(() => {
      source.emit(valueObservation(0));
    }).toThrow('before start()');
  });
});

describe('RecordingRenderTarget', () => {
  it('starts empty', () => {
    const target = new RecordingRenderTarget();
    expect(target.plans).toEqual([]);
    expect(target.lastPlan).toBeNull();
  });

  it('records every plan it receives, in order', async () => {
    const target = new RecordingRenderTarget();
    const a = plan('realSheet');
    const b = plan('previewSheet');
    await target.reconcile(a);
    await target.reconcile(b);
    expect(target.plans).toEqual([a, b]);
    expect(target.lastPlan).toBe(b);
  });
});
