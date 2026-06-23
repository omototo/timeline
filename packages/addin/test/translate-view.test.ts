import { describe, expect, it } from 'vitest';
import type { BranchMeta, Head, TimelineView as EngineTimelineView } from '@timeline/engine';
import { stepLabel, translateView } from '../src/ui/translate-view.ts';

const main: BranchMeta = { id: 'main', order: 0, provisional: false, name: 'main' };
const fork: BranchMeta = {
  id: 'fork-1',
  order: 1,
  provisional: true,
  parentBranchId: 'main',
  forkedAt: { branchId: 'main', stepIndex: 2 },
};

function engineView(steps: EngineTimelineView['steps']): EngineTimelineView {
  return { branches: [fork, main], steps };
}

describe('translateView', () => {
  it('nests Steps under their branch, ordered by step index, in tab order', () => {
    const view = translateView(
      engineView([
        { ref: { branchId: 'main', stepIndex: 1 }, kind: 'value', magnitude: 3 },
        { ref: { branchId: 'main', stepIndex: 0 }, kind: 'value', magnitude: 1 },
        { ref: { branchId: 'fork-1', stepIndex: 0 }, kind: 'structural', magnitude: 1 },
      ]),
      { branchId: 'main', mode: 'present' },
    );

    // Tab order is engine `order`: main (0) before fork-1 (1).
    expect(view.branches.map((b) => b.id)).toEqual(['main', 'fork-1']);
    const mainBranch = view.branches[0];
    expect(mainBranch?.steps.map((s) => s.index)).toEqual([0, 1]);
    expect(mainBranch?.name).toBe('main');
    expect(view.branches[1]?.provisional).toBe(true);
    expect(view.branches[1]?.parentBranchId).toBe('main');
    expect(view.branches[1]?.forkedAt).toEqual({ branchId: 'main', stepIndex: 2 });
  });

  it('carries magnitude and a synthesized label/kind onto each Step', () => {
    const view = translateView(
      engineView([{ ref: { branchId: 'main', stepIndex: 0 }, kind: 'value', magnitude: 1000 }]),
      { branchId: 'main', mode: 'present' },
    );
    const step = view.branches[0]?.steps[0];
    expect(step?.magnitude).toBe(1000);
    expect(step?.kind).toBe('value');
    expect(step?.label).toBe('1000 cells');
  });

  it('maps the head, preserving preview index only in preview', () => {
    const present: Head = { branchId: 'main', mode: 'present' };
    expect(translateView(engineView([]), present).head).toEqual(present);

    const preview: Head = { branchId: 'main', mode: 'preview', previewStepIndex: 4 };
    expect(translateView(engineView([]), preview).head).toEqual(preview);
  });

  it('passes the worksheet list through (empty by default)', () => {
    expect(translateView(engineView([]), { branchId: 'main', mode: 'present' }).sheets).toEqual([]);
    expect(
      translateView(engineView([]), { branchId: 'main', mode: 'present' }, ['Sheet1', 'Model'])
        .sheets,
    ).toEqual(['Sheet1', 'Model']);
  });
});

describe('stepLabel', () => {
  it('reads naturally per kind, singular/plural aware', () => {
    expect(stepLabel('value', 1)).toBe('1 cell');
    expect(stepLabel('value', 42)).toBe('42 cells');
    expect(stepLabel('reconciliation', 7)).toBe('reconciled 7 cells');
    expect(stepLabel('structural', 1)).toBe('structural change');
    expect(stepLabel('worksheet', 1)).toBe('worksheet change');
  });
});
