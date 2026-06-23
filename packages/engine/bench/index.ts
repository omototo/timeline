/**
 * Headless engine benchmark (ADR-0004) — runnable entry point.
 *
 * Run with `bun run bench` (from `packages/engine`) or `bun bench/index.ts`.
 * Drives synthetic Observations through the engine via the {@link harness}
 * stand-ins (ReplayChangeSource → Engine → fake RenderTarget → InMemoryStore),
 * measures, and prints a timing table.
 *
 * Stages measured (the headless half of ADR-0004):
 * - **ingest** latency for a 50k- and a 100k-cell paste (worst-case dense diff);
 * - **replay** latency reconstructing the tip after 100 / 1k / 10k single-cell
 *   Steps (forward-replay from the nearest keyframe);
 * - **keyframe serialize + compress** via the Node global `CompressionStream`
 *   (gzip) — bytes in/out + compression ratio + wall time.
 *
 * The on-host half (stages 1 + 5: `getValues` I/O floor + end-to-end
 * paste → Step committed against a live Excel host) is OUT OF SCOPE here.
 *
 * Pure: no Office.js, DOM, or React.
 */
import { TimelineEngineImpl, InMemoryStore } from '@timeline/engine';
import type { ShadowSnapshot } from '@timeline/engine';
import {
  RecordingRenderTarget,
  ReplayChangeSource,
  drainEffects,
  pasteObservation,
} from './harness.ts';

/** Wall-clock a synchronous closure in milliseconds (high-resolution). */
function timeSync<T>(fn: () => T): { ms: number; value: T } {
  const start = performance.now();
  const value = fn();
  return { ms: performance.now() - start, value };
}

/** One row of the printed timing table. */
interface Row {
  stage: string;
  detail: string;
  ms: number;
  notes: string;
}

const rows: Row[] = [];

// ---------------------------------------------------------------------------
// Ingest latency — big pastes (worst-case dense diff)
// ---------------------------------------------------------------------------

/**
 * Measure ingest latency for a single dense paste of `cells` cells. Fresh engine
 * + store each run so the diff is against an empty mirror (every cell changes).
 */
async function benchPaste(label: string, rows_: number, cols: number): Promise<void> {
  const engine = new TimelineEngineImpl();
  const store = new InMemoryStore();
  const render = new RecordingRenderTarget();
  const obs = pasteObservation('Sheet1', rows_, cols);

  const { ms, value: env } = timeSync(() => engine.ingest(obs));
  await drainEffects(env, render, store);

  const cells = rows_ * cols;
  rows.push({
    stage: 'ingest',
    detail: `${label} (${cells.toLocaleString()} cells)`,
    ms,
    notes: `${engine.shadowCellCount('Sheet1').toLocaleString()} cells in mirror, ${String(
      (env.persist ?? []).length,
    )} persist ops`,
  });
}

// ---------------------------------------------------------------------------
// Replay latency — reconstruct the tip after N Steps
// ---------------------------------------------------------------------------

/**
 * Build an engine with `steps` single-cell Steps, then measure reconstructing
 * the tip by forward-replay (from the nearest keyframe). Uses the engine's
 * additive `readReconstructed` query, which runs the same reconstruction path
 * `goto` uses.
 */
async function benchReplay(steps: number): Promise<void> {
  const engine = new TimelineEngineImpl();
  const store = new InMemoryStore();
  const render = new RecordingRenderTarget();
  const source = ReplayChangeSource.edits('Sheet1', steps);

  for (const obs of source.observations) {
    await drainEffects(engine.ingest(obs), render, store);
  }

  const tip = engine.tipStepIndex();
  const { ms } = timeSync(() =>
    engine.readReconstructed({ branchId: 'main', stepIndex: tip }, 'Sheet1', tip, 0),
  );

  const keyframes = engine.keyframeIndices().length;
  rows.push({
    stage: 'replay',
    detail: `${steps.toLocaleString()} steps`,
    ms,
    notes: `${String(keyframes)} keyframes, replay from kf ≤ ${String(tip)}`,
  });
}

// ---------------------------------------------------------------------------
// Keyframe serialize + compress (Node CompressionStream — ADR-0004 stage 3)
// ---------------------------------------------------------------------------

/** gzip a byte buffer through the Web/Node `CompressionStream` global. */
async function gzip(bytes: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([bytes as BlobPart]).stream().pipeThrough(new CompressionStream('gzip'));
  const buf = await new Response(stream).arrayBuffer();
  return new Uint8Array(buf);
}

/**
 * Serialize a keyframe snapshot (a paste-sized Shadow State) to JSON bytes and
 * gzip it via `CompressionStream`, reporting bytes in/out, ratio, and wall time.
 * This is ADR-0004 stage 3 — native compression, no third-party dep.
 */
async function benchKeyframeCompress(label: string, rows_: number, cols: number): Promise<void> {
  const engine = new TimelineEngineImpl();
  const store = new InMemoryStore();
  const render = new RecordingRenderTarget();

  // A dense paste of this size crosses the byte threshold, so the adaptive
  // cadence writes a keyframe whose payload is the full Shadow State snapshot.
  await drainEffects(engine.ingest(pasteObservation('Sheet1', rows_, cols)), render, store);

  // Serialize the loaded keyframe state from the store (the real payload).
  const kf = await store.loadKeyframeAtOrBefore('main', engine.tipStepIndex());
  const state: ShadowSnapshot = (kf?.state as ShadowSnapshot | undefined) ?? {
    sheets: [],
    sheetMeta: [],
  };

  const json = JSON.stringify(state);
  const inBytes = new TextEncoder().encode(json);

  const start = performance.now();
  const out = await gzip(inBytes);
  const ms = performance.now() - start;

  const ratio = inBytes.length === 0 ? 0 : out.length / inBytes.length;
  rows.push({
    stage: 'keyframe',
    detail: `${label} serialize+gzip`,
    ms,
    notes: `${inBytes.length.toLocaleString()} B -> ${out.length.toLocaleString()} B (${(
      ratio * 100
    ).toFixed(1)}% of original)`,
  });
}

// ---------------------------------------------------------------------------
// Print
// ---------------------------------------------------------------------------

function printTable(): void {
  const stageW = Math.max(5, ...rows.map((r) => r.stage.length));
  const detailW = Math.max(6, ...rows.map((r) => r.detail.length));
  const msW = 10;
  const header = `${'stage'.padEnd(stageW)}  ${'detail'.padEnd(detailW)}  ${'ms'.padStart(
    msW,
  )}  notes`;
  console.log('\nHeadless engine benchmark (ADR-0004 — engine-compute half)\n');
  console.log(header);
  console.log('-'.repeat(header.length + 20));
  for (const r of rows) {
    console.log(
      `${r.stage.padEnd(stageW)}  ${r.detail.padEnd(detailW)}  ${r.ms
        .toFixed(2)
        .padStart(msW)}  ${r.notes}`,
    );
  }
  console.log(
    '\nNote: the on-host half (getValues I/O floor + end-to-end paste->Step) needs real Excel and is out of scope.\n',
  );
}

async function main(): Promise<void> {
  // Ingest: 50k and 100k cell pastes.
  await benchPaste('50k paste', 500, 100);
  await benchPaste('100k paste', 1000, 100);

  // Replay depth: 100 / 1k / 10k steps.
  await benchReplay(100);
  await benchReplay(1_000);
  await benchReplay(10_000);

  // Keyframe serialize + compress (CompressionStream).
  await benchKeyframeCompress('50k', 500, 100);
  await benchKeyframeCompress('100k', 1000, 100);

  printTable();
}

await main();
