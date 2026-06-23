// Translate the engine's histogram model onto the UI contract.
//
// The engine returns a flat `{ branches: BranchMeta[], steps: { ref, kind,
// magnitude }[] }` and a separate `Head`; the UI wants `TimelineBranch[]` with
// each branch's Steps nested, plus a `TimelineHead` and the worksheet list.
// This is the one place engine vocabulary meets UI vocabulary — keeping it pure
// (no engine instance, no Office.js) means it is exhaustively unit-testable
// without a live workbook.

import type { BranchMeta, Head, TimelineView as EngineTimelineView } from '@timeline/engine';
import type {
  SheetId,
  StepKind,
  TimelineBranch,
  TimelineHead,
  TimelineStep,
  TimelineView,
} from './contract.ts';

/** A short, human label for a Step, synthesized from its kind + magnitude. */
export function stepLabel(kind: StepKind, magnitude: number): string {
  const cells = `${String(magnitude)} cell${magnitude === 1 ? '' : 's'}`;
  switch (kind) {
    case 'value':
      return cells;
    case 'reconciliation':
      return `reconciled ${cells}`;
    case 'structural':
      return 'structural change';
    case 'worksheet':
      return 'worksheet change';
  }
}

function toBranch(meta: BranchMeta, steps: TimelineStep[]): TimelineBranch {
  return {
    id: meta.id,
    provisional: meta.provisional,
    steps,
    ...(meta.name !== undefined ? { name: meta.name } : {}),
    ...(meta.parentBranchId !== undefined ? { parentBranchId: meta.parentBranchId } : {}),
    ...(meta.forkedAt !== undefined ? { forkedAt: meta.forkedAt } : {}),
  };
}

function toHead(head: Head): TimelineHead {
  return {
    branchId: head.branchId,
    mode: head.mode,
    ...(head.previewStepIndex !== undefined ? { previewStepIndex: head.previewStepIndex } : {}),
  };
}

/**
 * Map an engine `TimelineView` + `Head` onto the UI `TimelineView`.
 *
 * `sheets` is supplied by the caller (the real workbook's worksheet list); the
 * engine's flat Step model does not carry a per-Step sheet id, so the worksheet
 * drill-down stays at "Whole workbook" on real data until the engine surfaces
 * one (the UI tolerates an empty `sheets` list). Branches keep the engine's tab
 * order; Steps within a branch are ordered by step index.
 */
export function translateView(
  engineView: EngineTimelineView,
  head: Head,
  sheets: SheetId[] = [],
): TimelineView {
  const stepsByBranch = new Map<string, TimelineStep[]>();
  for (const step of engineView.steps) {
    const kind: StepKind = step.kind;
    const uiStep: TimelineStep = {
      index: step.ref.stepIndex,
      kind,
      magnitude: step.magnitude,
      sheetId: '',
      label: stepLabel(kind, step.magnitude),
    };
    const bucket = stepsByBranch.get(step.ref.branchId);
    if (bucket) {
      bucket.push(uiStep);
    } else {
      stepsByBranch.set(step.ref.branchId, [uiStep]);
    }
  }

  const branches = [...engineView.branches]
    .sort((a, b) => a.order - b.order)
    .map((meta) => {
      const steps = (stepsByBranch.get(meta.id) ?? []).sort((a, b) => a.index - b.index);
      return toBranch(meta, steps);
    });

  return { branches, head: toHead(head), sheets };
}
