# Engine seams: the Observation boundary, the store seam, and the lifecycle interface

## Context

A review of the keystone architecture (the host-agnostic Timeline Engine + its seams) found three soft spots in the original sketch:

1. The `ChangeSource` was drawn as emitting a finalized `Step`. That contradicts ADR-0001: the **engine** owns the Shadow State and produces the Delta/Step by diffing read-back against it. If the capture adapter emitted a Step it would need its own Shadow State — duplicating the engine and stealing its responsibility.
2. `HistoryStore` was rated "worth exploring." But branch/tip/keyframe semantics (ADR-0006/0007) and the engine's own multi-step/branch operations cannot be exercised without a store — it is load-bearing, not optional.
3. The `record/goto/branch/switch/head` interface omitted the entire attach/lifecycle surface from ADR-0006.

## Decision

0. **Execution model — functional core, imperative shell.** The engine is a synchronous, stateful, in-memory instance that holds the Shadow State and HEAD and **returns effect descriptions** (`{ reconcile, persist }`) — it never performs or awaits I/O. All async (Office.js read-back, IndexedDB persistence) lives in the addin shell. The interface is the test surface: drive a call sequence, assert on returned effects, no fakes. (Rejected: an engine that holds async ports and orchestrates I/O — testing orchestration timing instead of logic.)

1. **`ChangeSource` emits `Observation`s, not Steps.** An Observation is raw host facts about one (debounced) change: `{ address, changeType, changeDirectionState, triggerSource, source, read-back values/formulas/formats slab }`. The **adapter** performs the Office.js read-back (host I/O the engine cannot do) and the burst debounce; the **engine** diffs the Observation against the Shadow State to produce the Delta and decide Step boundaries. Capture never owns history state. (`ReplayChangeSource` emits synthetic Observations — the second adapter that makes the seam real.)
2. **The `HistoryStore` seam is part of the keystone.** `InMemoryStore` is implemented first (so the engine is testable for branch/keyframe scenarios and drives the headless benchmark); the IndexedDB adapter (ADR-0007) is deferred behind the same seam.
3. **The engine interface includes a lifecycle surface**, beyond `record/goto/branch/switch/head`: `attach(observedWorkbookState)` → drift reconciliation → a Reconciliation Step (ADR-0006); resume from persisted HEAD; workbook stamping; and tracking-disable on co-authoring (`source: Remote`). The full interface is grilled before it is frozen.

## Consequences

- The Observation boundary keeps the single-source-of-truth invariant: only the engine ever computes a Delta.
- Debounce placement (adapter coalesces the event burst; engine decides Step boundaries) is now explicit and is itself a grilling topic.
- The benchmark must drive both the headless path (`ReplayChangeSource → Engine → fake RenderTarget → InMemoryStore`) and the on-host path — see ADR-0004 update.
