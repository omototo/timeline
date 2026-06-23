/**
 * Headless benchmark harness (ADR-0004) — the engine-compute half.
 *
 * ADR-0004 gates the engine-language decision (TypeScript-in-a-Worker vs
 * Rust/WASM) on a benchmark. That benchmark has two halves:
 *
 * - **Headless (here):** `ReplayChangeSource` emits synthetic Observations →
 *   {@link TimelineEngineImpl} → a fake {@link RenderTarget} records the
 *   `ReconcileOp`s → {@link InMemoryStore} drains the `PersistOp`s. This measures
 *   pure ENGINE compute: ingest latency for big pastes, replay latency for
 *   100/1k/10k steps, and keyframe serialize+compress via the Node global
 *   `CompressionStream`. It answers "is the engine fast enough".
 * - **On-host (NOT here):** stages 1 + 5 (`getValues` I/O floor; end-to-end
 *   paste → Step committed) need a live Excel host and are out of scope. They
 *   answer "do we hit the 200 ms capture budget".
 *
 * Pure: no Office.js, DOM, or React. `performance` and `CompressionStream` are
 * Node/Web platform globals, not host APIs.
 */
import { InMemoryStore } from '@timeline/engine';
import type {
  CellSlab,
  CellState,
  EffectEnvelope,
  PersistOp,
  ReconcileOp,
  ReconcilePlan,
  ValueObservation,
} from '@timeline/engine';

// ---------------------------------------------------------------------------
// Fake RenderTarget — records ReconcileOps (the dumb shell render seam)
// ---------------------------------------------------------------------------

/**
 * A fake RenderTarget: the headless stand-in for the Office.js write seam. The
 * engine emits a {@link ReconcilePlan}; the real shell would apply it to Excel.
 * Here we just count the ops, so the harness measures engine compute without an
 * Excel round-trip (the I/O floor is the on-host half of ADR-0004).
 */
export class RecordingRenderTarget {
  opCount = 0;
  planCount = 0;

  render(plan: ReconcilePlan): void {
    this.planCount++;
    this.opCount += plan.ops.length;
  }

  /** Forget recorded ops (between bench phases). */
  reset(): void {
    this.opCount = 0;
    this.planCount = 0;
  }
}

// ---------------------------------------------------------------------------
// Synthetic Observation generators
// ---------------------------------------------------------------------------

/**
 * Build a dense rectangular paste slab of `rows × cols` distinct number cells.
 * Distinct values force every cell through the diff (a worst-case paste — no
 * cell is a no-op against the empty mirror).
 */
function pasteSlab(rows: number, cols: number): CellSlab {
  const values: CellSlab['values'] = [];
  const formulas: CellSlab['formulas'] = [];
  const numberFormats: CellSlab['numberFormats'] = [];
  const valueTypes: CellSlab['valueTypes'] = [];
  for (let r = 0; r < rows; r++) {
    const vRow: unknown[] = [];
    const fRow: (string | null)[] = [];
    const nRow: string[] = [];
    const tRow: CellState['valueType'][] = [];
    for (let c = 0; c < cols; c++) {
      vRow.push(r * cols + c);
      fRow.push(null);
      nRow.push('General');
      tRow.push('number');
    }
    values.push(vRow);
    formulas.push(fRow);
    numberFormats.push(nRow);
    valueTypes.push(tRow);
  }
  return { values, formulas, numberFormats, valueTypes };
}

/** A single-rectangle paste Observation of `rows × cols` cells at A1. */
export function pasteObservation(sheetId: string, rows: number, cols: number): ValueObservation {
  return {
    kind: 'value',
    triggerSource: 'thisLocalAddin',
    source: 'local',
    sheetId,
    area: [{ startRow: 0, startCol: 0, rowCount: rows, colCount: cols }],
    after: pasteSlab(rows, cols),
  };
}

/** A single-cell edit Observation at `(row, col)` carrying a number. */
export function editObservation(
  sheetId: string,
  row: number,
  col: number,
  value: number,
): ValueObservation {
  return {
    kind: 'value',
    triggerSource: 'thisLocalAddin',
    source: 'local',
    sheetId,
    area: [{ startRow: row, startCol: col, rowCount: 1, colCount: 1 }],
    after: {
      values: [[value]],
      formulas: [[null]],
      numberFormats: [['General']],
      valueTypes: [['number']],
    },
  };
}

// ---------------------------------------------------------------------------
// ReplayChangeSource — drives the synthetic Observation stream
// ---------------------------------------------------------------------------

/**
 * The synthetic change source for the headless bench. Emits a scripted stream of
 * Observations — a big paste, then a tail of single-cell edits — that the
 * harness pumps through the engine. Stands in for the Office.js change-event
 * adapter (the real `ChangeSource`).
 */
export class ReplayChangeSource {
  readonly #observations: ValueObservation[];

  constructor(observations: ValueObservation[]) {
    this.#observations = observations;
  }

  /** A big paste (`rows × cols`) followed by `editCount` single-cell edits. */
  static pasteThenEdits(
    sheetId: string,
    rows: number,
    cols: number,
    editCount: number,
  ): ReplayChangeSource {
    const obs: ValueObservation[] = [pasteObservation(sheetId, rows, cols)];
    for (let i = 0; i < editCount; i++) {
      // Edit cells well below the paste so each is a real, distinct change.
      obs.push(editObservation(sheetId, rows + i, 0, 1_000_000 + i));
    }
    return new ReplayChangeSource(obs);
  }

  /** `count` single-cell edits down column 0 (for replay-depth benches). */
  static edits(sheetId: string, count: number): ReplayChangeSource {
    const obs: ValueObservation[] = [];
    for (let i = 0; i < count; i++) {
      obs.push(editObservation(sheetId, i, 0, i));
    }
    return new ReplayChangeSource(obs);
  }

  get observations(): readonly ValueObservation[] {
    return this.#observations;
  }
}

// ---------------------------------------------------------------------------
// Effect drain — push an EffectEnvelope into the fake shell
// ---------------------------------------------------------------------------

/**
 * Drain an {@link EffectEnvelope} into the fake shell: record its reconcile plan
 * on the {@link RecordingRenderTarget} and apply its {@link PersistOp}s to the
 * {@link InMemoryStore}. This is the headless imperative shell (ADR-0013) — the
 * engine returned effects; here we execute them.
 */
export async function drainEffects(
  env: EffectEnvelope,
  render: RecordingRenderTarget,
  store: InMemoryStore,
): Promise<void> {
  if (env.reconcile !== undefined) render.render(env.reconcile);
  for (const op of env.persist ?? []) {
    await applyPersistOp(op, store);
  }
}

function applyPersistOp(op: PersistOp, store: InMemoryStore): Promise<void> {
  switch (op.op) {
    case 'appendDelta':
      return store.appendDelta(op.branchId, op.delta);
    case 'writeKeyframe':
      return store.writeKeyframe(op.branchId, op.stepIndex, op.state);
    case 'setHead':
      return store.setHead(op.head);
    case 'saveBranch':
      return store.saveBranch(op.meta);
    case 'deleteBranch':
      return store.deleteBranch(op.branchId);
  }
}

/** Count `setCells` ops in a plan (the headless render seam's signal). */
export function countSetCells(ops: readonly ReconcileOp[]): number {
  return ops.filter((o) => o.op === 'setCells').length;
}
