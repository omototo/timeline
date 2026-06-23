# Shadow State as source of truth; events as triggers; two delta classes

## Context

The product must reconstruct any past sheet state. The intuitive design is to treat the Office.js change-event stream as the source of truth — diff the event payloads. This does not work: `Worksheet.onChanged` populates `details.valueBefore/valueAfter` only for single-cell edits; for any multi-cell change (e.g. a 1,000-row paste) `details` is null, and the *before* state is already gone by the time the handler runs. Structural changes (`rowInserted`, `columnInserted`, etc.) report only an address and a type — they describe a coordinate shift, not a value change, and Office.js names none of the cells that moved.

## Decision

1. **The engine maintains a Shadow State** — a complete in-memory mirror of the live sheet — as its own source of truth. Office.js change events are treated as a *trigger and a bounding-box* only. On each event the JS layer reads the current state of the changed address via Office.js and hands it to the engine, which diffs against the Shadow State, records a Delta, and updates the Shadow State.

2. **Deltas come in two classes.** A *Value Delta* is a sparse `(address, before, after)` list replayed by writing values. A *Structural Delta* is a coordinate-remapping transform replayed by *applying the operation* and broadcasting it to every coordinate-keyed store (Shadow State, format map, chart/pivot anchors, formula engine). When a change event's `changeType` is structural, the value-diff path is suppressed so a single column insert records as one small Step, not a false 50,000-cell delta.

## Consequences

- Reading state back on every change adds an Office.js round-trip per Step, so per-keystroke granularity is off the table (see Step debounce). 
- The engine, not Excel, owns history; correctness depends entirely on the Shadow State staying faithful to the live sheet, so any external mutation we fail to observe is a divergence bug.
- The exact event fan-out of a single user action (and whether structural events arrive reliably) needs an empirical spike before build.
- **Read-back is bounded by two distinct Office.js limits (verified against the API reference).** The **5 MB** request/response payload cap applies to **Excel on the web only**; the **5,000,000-cell** read cap applies to **all platforms**. For wide lossless captures the 5 MB web cap bites well before 5M cells. The read-back that feeds the diff must therefore **chunk** large changed regions (validate `range.cellCount` before `context.sync()`; tile or use `RangeAreas` when exceeded) — chunking is not optional for the headline 50K–100K paste. The capture spike must measure chunked read-back, and proxy objects must be `untrack()`ed in the read/write loops to avoid unbounded memory growth.
- **Structural Delta capture is directly supported by the event (verified).** `changeType` (`DataChangeType` enum, ExcelApi 1.7) classifies the op (`rowInserted`/`columnDeleted`/…), and `changeDirectionState` (ExcelApi 1.14) reports the shift direction on the event itself — so a Structural Delta is capturable from the event on 1.14+ without inferring the shift. Below 1.14 the direction must be inferred from `changeType` + `address`. Single-cell `details` (ExcelApi 1.9) is a fast path for 1-cell edits only; multi-cell always requires read-back + Shadow State diff (docs: *"if triggered on multiple cells, this property cannot be retrieved"*).
