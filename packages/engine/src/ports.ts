/**
 * Timeline Engine — persistence seams (Q5).
 *
 * Two shell-side ports. The engine never calls them; it emits `PersistOp`s and
 * consumes loaded data passed back into `attach`/`goto`/`switch`. Both ports
 * are async (the IndexedDB adapter per ADR-0007 is the real target;
 * `InMemoryStore` is implemented first).
 */
import type { BranchId, BranchMeta, Delta, Head } from './types.ts';

/**
 * History persistence (async; IndexedDB per ADR-0007, `InMemoryStore` first).
 */
export interface HistoryStore {
  appendDelta(branchId: BranchId, delta: Delta): Promise<void>;
  writeKeyframe(branchId: BranchId, stepIndex: number, state: unknown): Promise<void>;
  /** Returns the highest keyframe whose stepIndex is <= the given stepIndex. */
  loadKeyframeAtOrBefore(
    branchId: BranchId,
    stepIndex: number,
  ): Promise<{ stepIndex: number; state: unknown } | null>;
  /** Returns every keyframe for a branch, step-ascending (for rehydration on launch). */
  listKeyframes(branchId: BranchId): Promise<{ stepIndex: number; state: unknown }[]>;
  /** Returns the inclusive [from, to] range of deltas for a branch. */
  loadDeltas(branchId: BranchId, from: number, to: number): Promise<Delta[]>;
  getHead(): Promise<Head | null>;
  setHead(head: Head): Promise<void>;
  saveBranch(meta: BranchMeta): Promise<void>;
  listBranches(): Promise<BranchMeta[]>;
  getBranch(id: BranchId): Promise<BranchMeta | null>;
  deleteBranch(id: BranchId): Promise<void>;
}

/** The tiny stamp written into in-file `workbook.settings` (ADR-0006). */
export interface WorkbookStampData {
  workbookGuid: string;
  tipHash: string;
}

/**
 * Workbook stamp (in-file `workbook.settings` per ADR-0006) — tiny, travels
 * with the `.xlsx`.
 */
export interface WorkbookStamp {
  read(): Promise<WorkbookStampData | null>;
  write(data: WorkbookStampData): Promise<void>;
}
