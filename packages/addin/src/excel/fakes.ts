/**
 * Test doubles for the Stream B seams (`seams.ts`).
 *
 * These are the headless stand-ins the engine/UI integration and adapter tests
 * use in place of the real Office.js-backed `ChangeSource` / `RenderTarget`:
 *
 * - `FakeChangeSource` lets a test push `Observation`s to the registered
 *   handler, simulating debounced user actions without Office.js events.
 * - `RecordingRenderTarget` records every `ReconcilePlan` it receives so a test
 *   can assert on what the engine asked the surface to render.
 */
import type { Observation, ReconcilePlan } from '@timeline/engine';
import type { ChangeSource, RenderTarget } from './seams.ts';

/**
 * A `ChangeSource` driven by the test. `start` registers the handler; `emit`
 * pushes an `Observation` to it as if one debounced user action occurred.
 */
export class FakeChangeSource implements ChangeSource {
  #handler: ((obs: Observation) => void) | null = null;

  start(onObservation: (obs: Observation) => void): Promise<void> {
    this.#handler = onObservation;
    return Promise.resolve();
  }

  stop(): Promise<void> {
    this.#handler = null;
    return Promise.resolve();
  }

  /** True between `start` and `stop`. */
  get started(): boolean {
    return this.#handler !== null;
  }

  /** Pushes one Observation to the registered handler. Throws if not started. */
  emit(obs: Observation): void {
    if (this.#handler === null) {
      throw new Error('FakeChangeSource: emit() called before start().');
    }
    this.#handler(obs);
  }
}

/** A `RenderTarget` that records every `ReconcilePlan` it is handed, in order. */
export class RecordingRenderTarget implements RenderTarget {
  readonly #plans: ReconcilePlan[] = [];

  reconcile(plan: ReconcilePlan): Promise<void> {
    this.#plans.push(plan);
    return Promise.resolve();
  }

  /** The plans received so far, oldest first. */
  get plans(): readonly ReconcilePlan[] {
    return this.#plans;
  }

  /** The most recently received plan, or `null` if none. */
  get lastPlan(): ReconcilePlan | null {
    return this.#plans.at(-1) ?? null;
  }
}
