// The coordinator that makes the timeline real.
//
// `FakeTimelineDataSource` reduces UI commands against an in-memory sample so
// the pane can develop in isolation. This is its production counterpart: it
// owns the live engine, the persistence store, the two render targets, and the
// Office change source, and it is the single place where a `TimelineCommand`
// becomes an engine call whose `EffectEnvelope` is fanned out to Excel (the
// reconcile plan) and the store (the persist ops). The UI never sees any of
// that — it still talks the `TimelineDataSource` contract.

import type {
  EffectEnvelope,
  HistoryStore,
  Observation,
  PersistOp,
  TimelineEngine,
} from '@timeline/engine';
import { NOOP_PREVIEW_CHROME, type PreviewChrome } from '../excel/preview-chrome.ts';
import type { ChangeSource, RenderTarget } from '../excel/seams.ts';
import type { SheetId, TimelineCommand, TimelineView } from './contract.ts';
import type { TimelineDataSource } from './data-source.ts';
import { translateView } from './translate-view.ts';

/** Apply one persist op from an `EffectEnvelope` to the history store. */
export async function applyPersistOp(store: HistoryStore, op: PersistOp): Promise<void> {
  switch (op.op) {
    case 'appendDelta':
      await store.appendDelta(op.branchId, op.delta);
      return;
    case 'writeKeyframe':
      await store.writeKeyframe(op.branchId, op.stepIndex, op.state);
      return;
    case 'setHead':
      await store.setHead(op.head);
      return;
    case 'saveBranch':
      await store.saveBranch(op.meta);
      return;
    case 'deleteBranch':
      await store.deleteBranch(op.branchId);
      return;
  }
}

export interface RealTimelineDataSourceOptions {
  readonly engine: TimelineEngine;
  readonly store: HistoryStore;
  readonly realTarget: RenderTarget;
  readonly previewTarget: RenderTarget;
  readonly changeSource: ChangeSource;
  /** Hides/restores real sheets around Preview (full-workbook rollback). */
  readonly chrome?: PreviewChrome;
  /** The real workbook's worksheet ids, for the drill-down (empty = whole-workbook only). */
  readonly sheets?: SheetId[];
  /** Surfaced so the shell can log Excel write failures; defaults to a no-op. */
  readonly onError?: (error: unknown) => void;
}

export class RealTimelineDataSource implements TimelineDataSource {
  readonly #engine: TimelineEngine;
  readonly #store: HistoryStore;
  readonly #realTarget: RenderTarget;
  readonly #previewTarget: RenderTarget;
  readonly #changeSource: ChangeSource;
  readonly #chrome: PreviewChrome;
  readonly #sheets: SheetId[];
  readonly #onError: (error: unknown) => void;
  readonly #listeners = new Set<() => void>();
  // `useSyncExternalStore` requires getView() to be referentially stable between
  // changes — recomputing a fresh view every call drives an infinite render
  // loop. Cache the snapshot and invalidate it only when state actually changes.
  #view: TimelineView | null = null;
  // Effects (Excel reconciles) MUST apply in dispatch order: a fast scrub fires
  // many gotos, and goto N's createPreviewSheet must finish before goto N+1's
  // setCells reads it — otherwise Excel throws "the requested resource does not
  // exist". Serialize every #apply through this tail promise.
  #applyTail: Promise<void> = Promise.resolve();

  constructor(options: RealTimelineDataSourceOptions) {
    this.#engine = options.engine;
    this.#store = options.store;
    this.#realTarget = options.realTarget;
    this.#previewTarget = options.previewTarget;
    this.#changeSource = options.changeSource;
    this.#chrome = options.chrome ?? NOOP_PREVIEW_CHROME;
    this.#sheets = options.sheets ?? [];
    this.#onError = options.onError ?? (() => undefined);
  }

  /**
   * Apply the `attach` envelope (drift reconciliation, restored head) then begin
   * listening for live edits. Each debounced Observation is ingested and its
   * effects applied, mirroring `dispatch`.
   */
  async start(attachEnvelope?: EffectEnvelope): Promise<void> {
    if (attachEnvelope) {
      await this.#apply(attachEnvelope);
    }
    await this.#changeSource.start((obs: Observation) => {
      const envelope = this.#engine.ingest(obs);
      this.#emit();
      this.#enqueueApply(envelope);
    });
    this.#emit();
  }

  async stop(): Promise<void> {
    await this.#changeSource.stop();
  }

  getView(): TimelineView {
    this.#view ??= translateView(this.#engine.timeline(), this.#engine.head(), this.#sheets);
    return this.#view;
  }

  subscribe(listener: () => void): () => void {
    this.#listeners.add(listener);
    return () => {
      this.#listeners.delete(listener);
    };
  }

  dispatch(command: TimelineCommand): void {
    const envelope = this.#route(command);
    // Engine state mutates synchronously, so the view is already current — notify
    // the UI now; the Excel write (reconcile) settles asynchronously after.
    this.#emit();
    if (envelope) {
      this.#enqueueApply(envelope);
    }
  }

  /** Chain an effect onto the serial apply queue so reconciles never overlap. */
  #enqueueApply(envelope: EffectEnvelope): void {
    this.#applyTail = this.#applyTail.then(() => this.#apply(envelope));
  }

  #route(command: TimelineCommand): EffectEnvelope | null {
    switch (command.type) {
      case 'goto':
        return this.#engine.goto(command.ref);
      case 'returnToPresent':
        return this.#engine.returnToPresent();
      case 'branch':
        return this.#engine.branch(command.from);
      case 'switch':
        return this.#engine.switch(command.branchId);
      case 'renameBranch':
      case 'deleteBranch':
        // The engine exposes no rename/delete branch op yet; these stay UI-only
        // on real data until it does. No-op rather than a misleading write.
        return null;
    }
  }

  async #apply(envelope: EffectEnvelope): Promise<void> {
    try {
      for (const op of envelope.persist ?? []) {
        await applyPersistOp(this.#store, op);
      }
      const plan = envelope.reconcile;
      if (plan) {
        // Hide the real sheets BEFORE the first preview plan creates surfaces, so
        // only the historical view is visible; restore them AFTER the teardown.
        if (plan.enterPreview) {
          await this.#chrome.enter();
        }
        const target = plan.target === 'previewSheet' ? this.#previewTarget : this.#realTarget;
        await target.reconcile(plan);
        if (plan.exitPreview) {
          await this.#chrome.exit();
        }
      }
    } catch (error) {
      this.#onError(error);
    }
  }

  #emit(): void {
    // Invalidate the cached snapshot first so the next getView() recomputes.
    this.#view = null;
    for (const listener of this.#listeners) {
      listener();
    }
  }
}
