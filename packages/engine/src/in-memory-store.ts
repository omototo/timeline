/**
 * `InMemoryStore` — a real, working in-memory `HistoryStore` (ADR-0013).
 *
 * Implemented first so the engine is testable for branch/keyframe scenarios and
 * can drive the headless benchmark, ahead of the IndexedDB adapter (ADR-0007)
 * behind the same seam.
 *
 * Storage model:
 * - deltas: per-branch append-only arrays, indexed by stepIndex (0-based).
 * - keyframes: per-branch maps from stepIndex -> serialized state.
 * - head: a single nullable HEAD.
 * - branches: a map of BranchMeta keyed by branch id.
 */
import type { HistoryStore } from './ports.ts';
import type { BranchId, BranchMeta, Delta, Head } from './types.ts';

export class InMemoryStore implements HistoryStore {
  readonly #deltas = new Map<BranchId, Delta[]>();
  readonly #keyframes = new Map<BranchId, Map<number, unknown>>();
  readonly #branches = new Map<BranchId, BranchMeta>();
  #head: Head | null = null;

  appendDelta(branchId: BranchId, delta: Delta): Promise<void> {
    const log = this.#deltas.get(branchId);
    if (log === undefined) {
      this.#deltas.set(branchId, [delta]);
    } else {
      log.push(delta);
    }
    return Promise.resolve();
  }

  writeKeyframe(branchId: BranchId, stepIndex: number, state: unknown): Promise<void> {
    let frames = this.#keyframes.get(branchId);
    if (frames === undefined) {
      frames = new Map<number, unknown>();
      this.#keyframes.set(branchId, frames);
    }
    frames.set(stepIndex, state);
    return Promise.resolve();
  }

  loadKeyframeAtOrBefore(
    branchId: BranchId,
    stepIndex: number,
  ): Promise<{ stepIndex: number; state: unknown } | null> {
    const frames = this.#keyframes.get(branchId);
    if (frames === undefined) {
      return Promise.resolve(null);
    }
    let bestIndex = -1;
    for (const candidate of frames.keys()) {
      if (candidate <= stepIndex && candidate > bestIndex) {
        bestIndex = candidate;
      }
    }
    if (bestIndex === -1) {
      return Promise.resolve(null);
    }
    // bestIndex was taken from frames.keys(), so the lookup is present.
    const state = frames.get(bestIndex);
    return Promise.resolve({ stepIndex: bestIndex, state });
  }

  listKeyframes(branchId: BranchId): Promise<{ stepIndex: number; state: unknown }[]> {
    const frames = this.#keyframes.get(branchId);
    if (frames === undefined) {
      return Promise.resolve([]);
    }
    const all = [...frames.entries()]
      .map(([stepIndex, state]) => ({ stepIndex, state }))
      .sort((a, b) => a.stepIndex - b.stepIndex);
    return Promise.resolve(all);
  }

  loadDeltas(branchId: BranchId, from: number, to: number): Promise<Delta[]> {
    const log = this.#deltas.get(branchId);
    if (log === undefined) {
      return Promise.resolve([]);
    }
    // Inclusive [from, to]; slice's end is exclusive. Clamp negative `from`.
    const start = Math.max(0, from);
    return Promise.resolve(log.slice(start, to + 1));
  }

  getHead(): Promise<Head | null> {
    return Promise.resolve(this.#head);
  }

  setHead(head: Head): Promise<void> {
    this.#head = head;
    return Promise.resolve();
  }

  saveBranch(meta: BranchMeta): Promise<void> {
    this.#branches.set(meta.id, meta);
    return Promise.resolve();
  }

  listBranches(): Promise<BranchMeta[]> {
    return Promise.resolve([...this.#branches.values()].sort((a, b) => a.order - b.order));
  }

  getBranch(id: BranchId): Promise<BranchMeta | null> {
    return Promise.resolve(this.#branches.get(id) ?? null);
  }

  deleteBranch(id: BranchId): Promise<void> {
    this.#branches.delete(id);
    this.#deltas.delete(id);
    this.#keyframes.delete(id);
    return Promise.resolve();
  }
}
