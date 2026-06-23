import type { Observation, ReconcilePlan } from '@timeline/engine';

// Shell-side seams for Stream B (Office.js adapters). Concrete Office.js
// implementations land alongside this file (office-change-source.ts,
// render-target.ts); a fake implements the same interfaces for headless tests.
// No Office.js types cross these seams — the engine stays host-agnostic.

/** Produces engine-neutral Observations from a host (Office.js) or a fake. */
export interface ChangeSource {
  /** Begin observing. `onObservation` fires once per debounced user action. */
  start(onObservation: (obs: Observation) => void): Promise<void>;
  /** Stop observing and release any host event handlers. */
  stop(): Promise<void>;
}

/** Applies a ReconcilePlan to a render surface (the real sheet or a preview sheet). */
export interface RenderTarget {
  reconcile(plan: ReconcilePlan): Promise<void>;
}
