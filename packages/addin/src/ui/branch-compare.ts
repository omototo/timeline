import type { BranchId, StepRef, TimelineBranch, TimelineStep, TimelineView } from './contract.ts';

export interface BranchCompareStats {
  readonly stepCount: number;
  readonly totalMagnitude: number;
}

export interface BranchComparison {
  readonly left: TimelineBranch;
  readonly right: TimelineBranch;
  readonly fork: StepRef | null;
  readonly firstDifferentStepIndex: number | null;
  readonly leftStats: BranchCompareStats;
  readonly rightStats: BranchCompareStats;
}

function totalMagnitude(steps: TimelineStep[]): number {
  return steps.reduce((total, step) => total + step.magnitude, 0);
}

function stats(branch: TimelineBranch): BranchCompareStats {
  return {
    stepCount: branch.steps.length,
    totalMagnitude: totalMagnitude(branch.steps),
  };
}

function sameStep(left: TimelineStep | undefined, right: TimelineStep | undefined): boolean {
  if (!left || !right) return false;
  return (
    left.kind === right.kind &&
    left.magnitude === right.magnitude &&
    left.sheetId === right.sheetId &&
    left.label === right.label
  );
}

function firstDifferentStepIndex(left: TimelineBranch, right: TimelineBranch): number | null {
  const max = Math.max(left.steps.length, right.steps.length);
  for (let index = 0; index < max; index += 1) {
    if (!sameStep(left.steps[index], right.steps[index])) return index;
  }
  return null;
}

function forkBetween(left: TimelineBranch, right: TimelineBranch): StepRef | null {
  if (right.parentBranchId === left.id && right.forkedAt) return right.forkedAt;
  if (left.parentBranchId === right.id && left.forkedAt) return left.forkedAt;
  return right.forkedAt ?? left.forkedAt ?? null;
}

export function compareBranches(
  view: TimelineView,
  leftBranchId: BranchId,
  rightBranchId: BranchId,
): BranchComparison | null {
  const left = view.branches.find((branch) => branch.id === leftBranchId);
  const right = view.branches.find((branch) => branch.id === rightBranchId);
  if (!left || !right) return null;

  return {
    left,
    right,
    fork: forkBetween(left, right),
    firstDifferentStepIndex: firstDifferentStepIndex(left, right),
    leftStats: stats(left),
    rightStats: stats(right),
  };
}
