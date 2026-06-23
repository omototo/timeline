# "Track everything" becomes tiered fidelity with honest gap-marking

## Context

"Track everything" is unbounded; Office.js will not honor it. The cell grid has change events (fine-grained Deltas come free), but charts, pivots, and shapes have **no change events** — their mutations can only be detected by snapshotting an object inventory at Step boundaries. Reconstruction is also API-capped: a basic pivot round-trips, a data-model pivot with calculated fields does not. Silently restoring a degraded object is worse than not tracking it, because the user believes the past is faithful.

## Decision

Replace "track everything" with three Fidelity Tiers, documented from the start in a living capability map (docs/capability-map.md):

- **Tier 1 — full fidelity, event-driven (v1):** the cell grid (values, formulas, number formats, font/fill/border/alignment) and structural ops captured via change events.
- **Tier 2 — config fidelity, snapshot-driven, best-effort (v1):** charts, basic pivots, images/shapes, view settings — captured as config snapshots at Step boundaries; restored by delete-and-rebuild.
- **Tier 3 — existence-tracked only, restore not guaranteed (out of v1):** data-model/OLAP pivots, calculated fields, slicers, sparklines, form controls, VBA, external queries — tracked only so a rollback raises a **Fidelity Caveat**.

Governing principle: lossless where the API permits, graceful degradation where it doesn't, and a Fidelity Caveat on the Step whenever fidelity isn't guaranteed. The system never pretends to a fidelity it can't deliver.

## Consequences

- Capture mechanism is not uniform: event-driven for the grid, snapshot-driven for embedded objects. The engine must run both paths.
- Several event-availability assumptions (merge, row/column resize, conditional-format edits) need empirical verification in the same spike as ADR-0001/0004 — flagged in the capability map.
