import type { TimelineView } from './contract';

// Realistic fixture: a main line with a structural op and a big paste, plus a
// "what-if" branch forked at step 2. Use this to develop and test the UI before
// the real engine is wired in.
export const sampleTimeline: TimelineView = {
  sheets: ['Sheet1', 'Assumptions', 'Summary'],
  head: { branchId: 'main', mode: 'present' },
  branches: [
    {
      id: 'main',
      provisional: false,
      steps: [
        { index: 0, kind: 'value', magnitude: 1, sheetId: 'Sheet1', label: 'A1 = 10' },
        { index: 1, kind: 'value', magnitude: 1000, sheetId: 'Sheet1', label: 'paste 1,000 rows' },
        { index: 2, kind: 'structural', magnitude: 3, sheetId: 'Sheet1', label: 'insert column B' },
        {
          index: 3,
          kind: 'value',
          magnitude: 12,
          sheetId: 'Assumptions',
          label: 'edit assumptions',
        },
        {
          index: 4,
          kind: 'worksheet',
          magnitude: 1,
          sheetId: 'Summary',
          label: 'add sheet Summary',
        },
      ],
    },
    {
      id: 'branch-b',
      name: 'what-if',
      parentBranchId: 'main',
      forkedAt: { branchId: 'main', stepIndex: 2 },
      provisional: false,
      steps: [
        { index: 0, kind: 'value', magnitude: 8, sheetId: 'Sheet1', label: 'alt scenario' },
        { index: 1, kind: 'value', magnitude: 40, sheetId: 'Sheet1', label: 'bulk edit' },
      ],
    },
  ],
};
