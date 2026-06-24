import { describe, expect, it } from 'vitest';
import { InMemoryStore } from '../src/in-memory-store.ts';
import { TimelineEngineImpl } from '../src/timeline-engine.ts';
import type {
  CellSlab,
  EffectEnvelope,
  RehydrationData,
  ValueObservation,
  WorkbookSnapshot,
} from '../src/types.ts';
import type { HistoryStore } from '../src/ports.ts';
import type { TimelineEngine } from '../src/engine.ts';

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

/** A single-cell edit (non-echo, local). */
function edit(row: number, col: number, value: number): ValueObservation {
  return {
    kind: 'value',
    triggerSource: 'unknown',
    source: 'local',
    sheetId: 'Sheet1',
    area: [{ startRow: row, startCol: col, rowCount: 1, colCount: 1 }],
    after: cell(value, 'number'),
  };
}

/** Persist every op an envelope produced, in order — mirrors the shell. */
async function persist(store: HistoryStore, envelope: EffectEnvelope): Promise<void> {
  for (const op of envelope.persist ?? []) {
    switch (op.op) {
      case 'appendDelta':
        await store.appendDelta(op.branchId, op.delta);
        break;
      case 'writeKeyframe':
        await store.writeKeyframe(op.branchId, op.stepIndex, op.state);
        break;
      case 'setHead':
        await store.setHead(op.head);
        break;
      case 'saveBranch':
        await store.saveBranch(op.meta);
        break;
      case 'deleteBranch':
        await store.deleteBranch(op.branchId);
        break;
    }
  }
}

/** Load the persisted timeline back into a RehydrationData payload (shell side). */
async function loadRehydrationData(store: HistoryStore): Promise<RehydrationData> {
  const head = await store.getHead();
  const branches = await store.listBranches();
  const branchIds = new Set<string>(['main', ...branches.map((b) => b.id)]);
  const perBranch = [];
  for (const branchId of branchIds) {
    perBranch.push({
      branchId,
      deltas: await store.loadDeltas(branchId, 0, Number.MAX_SAFE_INTEGER),
      keyframes: await store.listKeyframes(branchId),
    });
  }
  return { head, branches, perBranch };
}

/** The value written to A1 by a preview reconcile plan (asserts reconstruction). */
function a1PreviewValue(envelope: EffectEnvelope): unknown {
  for (const op of envelope.reconcile?.ops ?? []) {
    if (op.op === 'setCells' && op.area.at(0)?.startRow === 0 && op.area.at(0)?.startCol === 0) {
      return op.slab.values.at(0)?.at(0);
    }
  }
  return undefined;
}

async function buildSeededEngineAndStore(): Promise<{
  engine: TimelineEngine;
  store: InMemoryStore;
}> {
  const engine = new TimelineEngineImpl();
  const store = new InMemoryStore();
  engine.attach(SNAPSHOT, null);
  await persist(store, engine.ingest(edit(0, 0, 10))); // main step 0: A1=10
  await persist(store, engine.ingest(edit(0, 0, 20))); // main step 1: A1=20
  await persist(store, engine.ingest(edit(0, 1, 99))); // main step 2: B1=99
  await persist(store, engine.branch({ branchId: 'main', stepIndex: 1 })); // fork at main#1
  await persist(store, engine.ingest(edit(0, 0, 777))); // branch-1 step 0: A1=777
  return { engine, store };
}

describe('rehydrate', () => {
  it('restores the timeline (branches, steps, head) from a store round-trip', async () => {
    const { engine: original, store } = await buildSeededEngineAndStore();
    const before = original.timeline();
    const head = original.head();

    const restored = new TimelineEngineImpl();
    restored.rehydrate(await loadRehydrationData(store));
    restored.attach(SNAPSHOT, null);

    const after = restored.timeline();
    expect(after.branches).toEqual(before.branches);
    expect(after.steps).toEqual(before.steps);
    expect(restored.head()).toEqual(head); // head was on the fork, in Present
    expect(head.branchId).toBe('branch-1');
  });

  it('reconstructs a fork after rehydration (recomputed base keyframe)', async () => {
    const { store } = await buildSeededEngineAndStore();
    const restored = new TimelineEngineImpl();
    restored.rehydrate(await loadRehydrationData(store));
    restored.attach(SNAPSHOT, null);

    // main#0 previews A1=10.
    expect(a1PreviewValue(restored.goto({ branchId: 'main', stepIndex: 0 }))).toBe(10);
    restored.returnToPresent();
    // branch-1#0 = fork base (main#1: A1=20) + the branch's own edit (A1=777).
    // Correct only if the fork's base keyframe was recomputed from the parent.
    expect(a1PreviewValue(restored.goto({ branchId: 'branch-1', stepIndex: 0 }))).toBe(777);
  });

  it('continues recording correctly after rehydration (branch ids do not collide)', async () => {
    const { store } = await buildSeededEngineAndStore();
    const restored = new TimelineEngineImpl();
    restored.rehydrate(await loadRehydrationData(store));
    restored.attach(SNAPSHOT, null);

    // A new fork must mint a fresh id, not reuse branch-1.
    restored.branch({ branchId: 'main', stepIndex: 0 });
    const ids = restored.timeline().branches.map((b) => b.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids).toContain('branch-2');
  });

  it('normalizes a persisted preview head to Present (projection is not restorable)', async () => {
    const store = new InMemoryStore();
    await store.setHead({ branchId: 'main', mode: 'preview', previewStepIndex: 3 });
    const engine = new TimelineEngineImpl();
    engine.rehydrate(await loadRehydrationData(store));
    // A reload cannot resume a preview (the projection was in-memory) — land in
    // Present so the user is not stuck in a phantom preview.
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'present' });
  });

  it('is a no-op restore for an empty store (fresh workbook)', async () => {
    const store = new InMemoryStore();
    const engine = new TimelineEngineImpl();
    engine.rehydrate(await loadRehydrationData(store));
    engine.attach(SNAPSHOT, null);
    expect(engine.timeline().steps).toHaveLength(0);
    expect(engine.head()).toEqual({ branchId: 'main', mode: 'present' });
  });
});
