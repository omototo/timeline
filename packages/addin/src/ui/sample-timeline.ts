import type { SheetId, StepKind, TimelineStep, TimelineView } from './contract';

const SHEETS: SheetId[] = ['Sheet1', 'Assumptions', 'Summary', 'Model'];
const FALLBACK_SHEET: SheetId = 'Sheet1';

function kindFor(index: number): StepKind {
  if (index === 32) return 'reconciliation';
  if (index % 17 === 0) return 'worksheet';
  if (index % 9 === 0) return 'structural';
  return 'value';
}

function magnitudeFor(index: number): number {
  if (index === 1) return 1000;
  if (index % 13 === 0) return 650 + index * 3;
  if (index % 9 === 0) return 24;
  return 1 + ((index * 7) % 55);
}

function labelFor(index: number, kind: StepKind): string {
  if (index === 0) return 'A1 = 10';
  if (index === 1) return 'paste 1,000 rows';
  if (kind === 'structural') return `reshape model area ${String(index)}`;
  if (kind === 'worksheet') return `worksheet topology change ${String(index)}`;
  if (kind === 'reconciliation') return 'reattach reconciliation';
  return `value edit batch ${String(index)}`;
}

function step(index: number, sheetOffset = 0): TimelineStep {
  const kind = kindFor(index);
  const sheetId = SHEETS[(index + sheetOffset) % SHEETS.length] ?? FALLBACK_SHEET;
  return {
    index,
    kind,
    magnitude: magnitudeFor(index),
    sheetId,
    label: labelFor(index, kind),
  };
}

function steps(count: number, sheetOffset = 0): TimelineStep[] {
  return Array.from({ length: count }, (_, index) => step(index, sheetOffset));
}

export const sampleTimeline: TimelineView = {
  sheets: SHEETS,
  head: { branchId: 'main', mode: 'present' },
  branches: [
    {
      id: 'main',
      name: 'main',
      provisional: false,
      steps: steps(72),
    },
    {
      id: 'branch-b',
      name: 'what-if',
      parentBranchId: 'main',
      forkedAt: { branchId: 'main', stepIndex: 14 },
      provisional: false,
      steps: steps(28, 1).map((candidate) => ({
        ...candidate,
        label:
          candidate.index === 0 ? 'alt scenario' : `what-if adjustment ${String(candidate.index)}`,
      })),
    },
    {
      id: 'branch-c',
      name: 'audit cleanup',
      parentBranchId: 'main',
      forkedAt: { branchId: 'main', stepIndex: 32 },
      provisional: true,
      steps: steps(10, 2).map((candidate) => ({
        ...candidate,
        label: `audit cleanup ${String(candidate.index)}`,
      })),
    },
  ],
};
