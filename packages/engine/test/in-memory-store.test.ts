import { describe, it, expect, beforeEach } from 'vitest';
import { InMemoryStore } from '../src/index.ts';
import type { BranchMeta, Delta, Head, StructuralDelta, ValueDelta } from '../src/index.ts';

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

describe('InMemoryStore', () => {
  let store: InMemoryStore;

  beforeEach(() => {
    store = new InMemoryStore();
  });

  describe('appendDelta / loadDeltas', () => {
    it('round-trips appended deltas in order', async () => {
      const deltas = [structuralDelta(0), structuralDelta(1), structuralDelta(2)];
      for (const d of deltas) {
        await store.appendDelta(BRANCH, d);
      }
      expect(await store.loadDeltas(BRANCH, 0, 2)).toEqual(deltas);
    });

    it('returns the inclusive [from, to] range', async () => {
      for (let i = 0; i < 5; i++) {
        await store.appendDelta(BRANCH, structuralDelta(i));
      }
      const slice = await store.loadDeltas(BRANCH, 1, 3);
      expect(slice).toEqual([structuralDelta(1), structuralDelta(2), structuralDelta(3)]);
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
      // to beyond the end returns through the last element, not an error.
      expect(await store.loadDeltas(BRANCH, 2, 99)).toEqual([structuralDelta(2)]);
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

    it('is a no-op deleting an unknown branch', async () => {
      await expect(store.deleteBranch('ghost')).resolves.toBeUndefined();
    });
  });
});
