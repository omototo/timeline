import type {
  BranchId,
  TimelineBranch,
  TimelineCommand,
  TimelineHead,
  TimelineView,
} from './contract.ts';
import { sampleTimeline } from './sample-timeline.ts';

/**
 * The seam between the UI and its data. The pane renders whatever `getView()`
 * returns, re-renders when a `subscribe`d listener fires, and sends user actions
 * out via `dispatch`. A fake backs development and tests; at integration the
 * real implementation wraps the engine's `timeline()` query and routes commands
 * to the engine + adapters — a one-line provider swap, no UI changes.
 */
export interface TimelineDataSource {
  getView(): TimelineView;
  subscribe(listener: () => void): () => void;
  dispatch(command: TimelineCommand): void;
}

/** A stateful in-memory fake so the UI feels live without the real engine. */
export class FakeTimelineDataSource implements TimelineDataSource {
  private view: TimelineView;
  private readonly listeners = new Set<() => void>();
  private branchSeq: number;

  constructor(initial: TimelineView = structuredClone(sampleTimeline)) {
    this.view = initial;
    this.branchSeq = initial.branches.length;
  }

  getView(): TimelineView {
    return this.view;
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  dispatch(command: TimelineCommand): void {
    switch (command.type) {
      case 'goto':
        this.setHead({
          branchId: command.ref.branchId,
          mode: 'preview',
          previewStepIndex: command.ref.stepIndex,
        });
        return;
      case 'returnToPresent':
        this.setHead({ branchId: this.view.head.branchId, mode: 'present' });
        return;
      case 'switch':
        this.setHead({ branchId: command.branchId, mode: 'present' });
        return;
      case 'branch':
        this.forkBranch(command.from.branchId, command.from.stepIndex);
        return;
      case 'renameBranch':
        this.renameBranch(command.branchId, command.name);
        return;
      case 'deleteBranch':
        this.deleteBranch(command.branchId);
        return;
    }
  }

  private setHead(head: TimelineHead): void {
    this.view = { ...this.view, head };
    this.emit();
  }

  private forkBranch(parentBranchId: BranchId, stepIndex: number): void {
    this.branchSeq += 1;
    const id = `fork-${String(this.branchSeq)}`;
    const branch: TimelineBranch = {
      id,
      name: `branch ${String(this.branchSeq)}`,
      parentBranchId,
      forkedAt: { branchId: parentBranchId, stepIndex },
      provisional: true,
      steps: [],
    };
    this.view = {
      ...this.view,
      branches: [...this.view.branches, branch],
      head: { branchId: id, mode: 'present' },
    };
    this.emit();
  }

  private renameBranch(branchId: BranchId, name: string): void {
    const trimmed = name.trim();
    if (trimmed.length === 0) return;

    const current = this.view.branches.find((branch) => branch.id === branchId);
    if (!current || current.name === trimmed) return;

    const branches = this.view.branches.map((branch) => {
      if (branch.id !== branchId || branch.name === trimmed) return branch;
      return { ...branch, name: trimmed };
    });

    this.view = { ...this.view, branches };
    this.emit();
  }

  private deleteBranch(branchId: BranchId): void {
    if (branchId === 'main') return;

    const deleted = this.view.branches.find((branch) => branch.id === branchId);
    if (!deleted) return;

    const branches = this.view.branches.filter((branch) => branch.id !== branchId);
    const fallbackBranchId =
      deleted.parentBranchId && branches.some((branch) => branch.id === deleted.parentBranchId)
        ? deleted.parentBranchId
        : branches[0]?.id;

    if (!fallbackBranchId) return;

    this.view = {
      ...this.view,
      branches,
      head:
        this.view.head.branchId === branchId
          ? { branchId: fallbackBranchId, mode: 'present' }
          : this.view.head,
    };
    this.emit();
  }

  private emit(): void {
    for (const listener of this.listeners) listener();
  }
}
