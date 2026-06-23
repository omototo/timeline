// `fake-indexeddb/auto` installs the structured-clone-backed IndexedDB globals
// (incl. `IDBKeyRange`) onto the jsdom environment, which omits IndexedDB.
// We still hand each store its own `new IDBFactory()` for per-test isolation.
import 'fake-indexeddb/auto';
import { describe, it, expect, beforeEach } from 'vitest';
import { IDBFactory } from 'fake-indexeddb';
import { IndexedDbStore } from '../src/excel/indexeddb-store.ts';
import type { BranchMeta, Delta, Head, StructuralDelta, ValueDelta } from '@timeline/engine';

const BRANCH = 'main';

/** A trivially-distinguishable structural delta carrying its ordinal in startRow. */
function structuralDelta(n: number): StructuralDelta {
  return {
    kind: 'structural',
    sheetId: 'Sheet1',
    changeType: 'rowInserted',
    address: { startRow: n, startCol: 0, rowCount: 1, colCount: 1 },
  };
}

/** A fresh, initialised store backed by an isolated fake-indexeddb instance. */
async function freshStore(): Promise<IndexedDbStore> {
  const store = new IndexedDbStore(new IDBFactory());
  await store.init();
  return store;
}

describe('IndexedDbStore', () => {
  let store: IndexedDbStore;

  beforeEach(async () => {
    store = await freshStore();
  });

  describe('init', () => {
    it('is idempotent (a second call is a no-op)', async () => {
      await expect(store.init()).resolves.toBeUndefined();
    });

    it('throws when used before init()', async () => {
      const uninit = new IndexedDbStore(new IDBFactory());
      await expect(uninit.getHead()).rejects.toThrow('not initialised');
    });

    it('best-effort requests persistent storage when available', async () => {
      const calls: number[] = [];
      const original = globalThis.navigator;
      Object.defineProperty(globalThis, 'navigator', {
        value: {
          storage: {
            persist: (): Promise<boolean> => {
              calls.push(1);
              return Promise.resolve(true);
            },
          },
        },
        configurable: true,
      });
      try {
        const s = new IndexedDbStore(new IDBFactory());
        await s.init();
        expect(calls).toHaveLength(1);
      } finally {
        Object.defineProperty(globalThis, 'navigator', { value: original, configurable: true });
      }
    });

    it('survives a rejecting persist() (persistence is best-effort)', async () => {
      const original = globalThis.navigator;
      Object.defineProperty(globalThis, 'navigator', {
        value: { storage: { persist: (): Promise<boolean> => Promise.reject(new Error('nope')) } },
        configurable: true,
      });
      try {
        const s = new IndexedDbStore(new IDBFactory());
        await expect(s.init()).resolves.toBeUndefined();
      } finally {
        Object.defineProperty(globalThis, 'navigator', { value: original, configurable: true });
      }
    });
  });

  describe('appendDelta / loadDeltas', () => {
    it('round-trips appended deltas in order', async () => {
      const deltas = [structuralDelta(0), structuralDelta(1), structuralDelta(2)];
      for (const d of deltas) {
        await store.appendDelta(BRANCH, d);
      }
      expect(await store.loadDeltas(BRANCH, 0, 2)).toEqual(deltas);
    });

    it('assigns strictly increasing indices to sequential appends, ordered by loadDeltas', async () => {
      // Distinguish each delta by its ordinal (carried in startRow) so we can
      // assert both the order AND that each landed at a distinct, increasing key.
      const appended = Array.from({ length: 6 }, (_, i) => structuralDelta(i));
      for (const d of appended) {
        await store.appendDelta(BRANCH, d);
      }
      const loaded = await store.loadDeltas(BRANCH, 0, appended.length - 1);
      // loadDeltas reads the [branchId, stepIndex] range in key order, so this
      // round-trip proves the appended stepIndices are strictly increasing and
      // contiguous (a duplicate or out-of-order key would corrupt this).
      expect(loaded).toEqual(appended);
      const ordinals = loaded.map((d) => (d as StructuralDelta).address.startRow);
      expect(ordinals).toEqual([0, 1, 2, 3, 4, 5]);
    });

    it('returns the inclusive [from, to] range', async () => {
      for (let i = 0; i < 5; i++) {
        await store.appendDelta(BRANCH, structuralDelta(i));
      }
      expect(await store.loadDeltas(BRANCH, 1, 3)).toEqual([
        structuralDelta(1),
        structuralDelta(2),
        structuralDelta(3),
      ]);
    });

    it('honours the lower range boundary (from is included)', async () => {
      for (let i = 0; i < 3; i++) {
        await store.appendDelta(BRANCH, structuralDelta(i));
      }
      expect(await store.loadDeltas(BRANCH, 0, 0)).toEqual([structuralDelta(0)]);
    });

    it('honours the upper range boundary (to is included, clamps past the end)', async () => {
      for (let i = 0; i < 3; i++) {
        await store.appendDelta(BRANCH, structuralDelta(i));
      }
      expect(await store.loadDeltas(BRANCH, 2, 99)).toEqual([structuralDelta(2)]);
    });

    it('clamps a negative `from` to 0', async () => {
      for (let i = 0; i < 3; i++) {
        await store.appendDelta(BRANCH, structuralDelta(i));
      }
      expect(await store.loadDeltas(BRANCH, -5, 1)).toEqual([
        structuralDelta(0),
        structuralDelta(1),
      ]);
    });

    it('returns empty when `to` is below `from`', async () => {
      await store.appendDelta(BRANCH, structuralDelta(0));
      expect(await store.loadDeltas(BRANCH, 2, 1)).toEqual([]);
    });

    it('returns an empty array for an unknown branch', async () => {
      expect(await store.loadDeltas('nope', 0, 10)).toEqual([]);
    });

    it('keeps branches isolated', async () => {
      await store.appendDelta('a', structuralDelta(0));
      await store.appendDelta('b', structuralDelta(1));
      expect(await store.loadDeltas('a', 0, 10)).toEqual([structuralDelta(0)]);
      expect(await store.loadDeltas('b', 0, 10)).toEqual([structuralDelta(1)]);
    });

    it('accepts a ValueDelta shape', async () => {
      const vd: ValueDelta = {
        kind: 'value',
        sheetId: 'Sheet1',
        cells: [
          {
            addr: { startRow: 0, startCol: 0, rowCount: 1, colCount: 1 },
            before: { value: 1, formula: null, valueType: 'number', numberFormat: 'General' },
            after: { value: 2, formula: null, valueType: 'number', numberFormat: 'General' },
          },
        ],
      };
      await store.appendDelta(BRANCH, vd);
      const loaded: Delta[] = await store.loadDeltas(BRANCH, 0, 0);
      expect(loaded[0]).toEqual(vd);
    });
  });

  describe('writeKeyframe / loadKeyframeAtOrBefore', () => {
    it('returns null when no keyframe exists', async () => {
      expect(await store.loadKeyframeAtOrBefore(BRANCH, 100)).toBeNull();
    });

    it('returns an exact-match keyframe', async () => {
      await store.writeKeyframe(BRANCH, 10, { snapshot: 10 });
      expect(await store.loadKeyframeAtOrBefore(BRANCH, 10)).toEqual({
        stepIndex: 10,
        state: { snapshot: 10 },
      });
    });

    it('returns the highest keyframe at or before stepIndex', async () => {
      await store.writeKeyframe(BRANCH, 0, { snapshot: 0 });
      await store.writeKeyframe(BRANCH, 100, { snapshot: 100 });
      await store.writeKeyframe(BRANCH, 200, { snapshot: 200 });
      expect(await store.loadKeyframeAtOrBefore(BRANCH, 150)).toEqual({
        stepIndex: 100,
        state: { snapshot: 100 },
      });
    });

    it('returns null when the only keyframes are after stepIndex', async () => {
      await store.writeKeyframe(BRANCH, 50, { snapshot: 50 });
      expect(await store.loadKeyframeAtOrBefore(BRANCH, 49)).toBeNull();
    });

    it('overwrites a keyframe at the same stepIndex', async () => {
      await store.writeKeyframe(BRANCH, 5, { snapshot: 'old' });
      await store.writeKeyframe(BRANCH, 5, { snapshot: 'new' });
      expect(await store.loadKeyframeAtOrBefore(BRANCH, 5)).toEqual({
        stepIndex: 5,
        state: { snapshot: 'new' },
      });
    });

    it('keeps keyframes isolated per branch', async () => {
      await store.writeKeyframe('a', 10, { snapshot: 'a' });
      expect(await store.loadKeyframeAtOrBefore('b', 10)).toBeNull();
    });
  });

  describe('head get / set', () => {
    it('starts null', async () => {
      expect(await store.getHead()).toBeNull();
    });

    it('round-trips a HEAD', async () => {
      const head: Head = { branchId: BRANCH, mode: 'present' };
      await store.setHead(head);
      expect(await store.getHead()).toEqual(head);
    });

    it('overwrites the previous HEAD', async () => {
      await store.setHead({ branchId: BRANCH, mode: 'present' });
      const next: Head = { branchId: BRANCH, mode: 'preview', previewStepIndex: 3 };
      await store.setHead(next);
      expect(await store.getHead()).toEqual(next);
    });
  });

  describe('branch save / list / get / delete', () => {
    const a: BranchMeta = { id: 'a', order: 1, provisional: false };
    const b: BranchMeta = { id: 'b', order: 0, provisional: true };

    it('returns null for an unknown branch', async () => {
      expect(await store.getBranch('missing')).toBeNull();
    });

    it('saves and gets a branch', async () => {
      await store.saveBranch(a);
      expect(await store.getBranch('a')).toEqual(a);
    });

    it('lists branches sorted by order', async () => {
      await store.saveBranch(a);
      await store.saveBranch(b);
      expect(await store.listBranches()).toEqual([b, a]);
    });

    it('updates an existing branch on re-save', async () => {
      await store.saveBranch(a);
      const updated: BranchMeta = { ...a, name: 'renamed' };
      await store.saveBranch(updated);
      expect(await store.getBranch('a')).toEqual(updated);
      expect(await store.listBranches()).toHaveLength(1);
    });

    it('deletes a branch and its deltas/keyframes', async () => {
      await store.saveBranch(a);
      await store.appendDelta('a', structuralDelta(0));
      await store.writeKeyframe('a', 0, { snapshot: 0 });
      await store.deleteBranch('a');
      expect(await store.getBranch('a')).toBeNull();
      expect(await store.listBranches()).toEqual([]);
      expect(await store.loadDeltas('a', 0, 10)).toEqual([]);
      expect(await store.loadKeyframeAtOrBefore('a', 10)).toBeNull();
    });

    it('leaves a sibling branch untouched when deleting', async () => {
      await store.saveBranch(a);
      await store.saveBranch(b);
      await store.appendDelta('a', structuralDelta(0));
      await store.appendDelta('b', structuralDelta(9));
      await store.deleteBranch('a');
      expect(await store.getBranch('b')).toEqual(b);
      expect(await store.loadDeltas('b', 0, 10)).toEqual([structuralDelta(9)]);
    });

    it('is a no-op deleting an unknown branch', async () => {
      await expect(store.deleteBranch('ghost')).resolves.toBeUndefined();
    });
  });

  describe('durability across reopen (same factory)', () => {
    it('persists records across a fresh store instance on the same DB', async () => {
      const factory = new IDBFactory();
      const first = new IndexedDbStore(factory);
      await first.init();
      await first.appendDelta(BRANCH, structuralDelta(0));
      await first.setHead({ branchId: BRANCH, mode: 'present' });

      const second = new IndexedDbStore(factory);
      await second.init();
      expect(await second.loadDeltas(BRANCH, 0, 0)).toEqual([structuralDelta(0)]);
      expect(await second.getHead()).toEqual({ branchId: BRANCH, mode: 'present' });
    });
  });
});
