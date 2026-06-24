// UI-side contract for the timeline task pane.
//
// The UI depends ONLY on this file — never on Office.js or the engine package.
// At integration, a thin layer maps the engine's `timeline()` output onto
// `TimelineView` and routes each `TimelineCommand` to the engine + adapters.
// Keeping this decoupled lets the UI develop against a fake while the engine and
// adapters land on their own branches.

export type SheetId = string;
export type BranchId = string;
export type TimelineTheme = 'light' | 'dark';

export type StepKind = 'value' | 'structural' | 'worksheet' | 'reconciliation';

/** The specific operation a Step represents — drives its timeline icon. */
export type TimelineOp =
  | 'edit'
  | 'formula'
  | 'paste'
  | 'clear'
  | 'insert-row'
  | 'delete-row'
  | 'insert-col'
  | 'delete-col'
  | 'insert-cells'
  | 'delete-cells'
  | 'sheet-add'
  | 'sheet-delete'
  | 'sheet-rename'
  | 'sheet-reorder'
  | 'reconcile';

export interface StepRef {
  branchId: BranchId;
  stepIndex: number;
}

export interface TimelineStep {
  index: number;
  kind: StepKind;
  /** The specific operation — drives the step's timeline icon. */
  op: TimelineOp;
  /** Bar height in the histogram: cells changed for value steps, op-weight otherwise. */
  magnitude: number;
  sheetId: SheetId;
  label?: string;
}

export interface TimelineBranch {
  id: BranchId;
  name?: string;
  parentBranchId?: BranchId;
  forkedAt?: StepRef;
  provisional: boolean;
  steps: TimelineStep[];
}

export interface TimelineHead {
  branchId: BranchId;
  mode: 'present' | 'preview';
  previewStepIndex?: number;
}

export interface TimelineView {
  branches: TimelineBranch[];
  head: TimelineHead;
  /** All known sheets — drives the worksheet drill-down control. */
  sheets: SheetId[];
}

export type TimelineCommand =
  | { type: 'goto'; ref: StepRef }
  | { type: 'returnToPresent' }
  | { type: 'branch'; from: StepRef }
  | { type: 'switch'; branchId: BranchId }
  | { type: 'renameBranch'; branchId: BranchId; name: string }
  | { type: 'deleteBranch'; branchId: BranchId };

export interface TimelinePaneProps {
  view: TimelineView;
  dispatch: (cmd: TimelineCommand) => void;
  theme?: TimelineTheme;
}
