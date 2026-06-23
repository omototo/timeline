# Excel is the only calculation authority; no second formula engine

## Context

The PRD specified HyperFormula as a core layer to recalculate formula references after structural changes (e.g. inserting a column shifts references in dependent formulas). But Excel is itself a formula engine: reconciling a state onto a Render Target triggers native recalculation, and replaying a Structural Delta by performing the actual insert/delete through Office.js makes Excel shift formula references natively — the exact job assigned to HyperFormula.

## Decision

Cut HyperFormula. Excel is the single calculation authority. The engine stores **formula text** (via `getFormulas`) in Value Deltas and **evaluated values** for Value Freezing; replay writes formulas back and lets Excel recalc; structural reference-shifting relies on native Excel behaviour during the applied insert/delete.

## Consequences

- Eliminates an entire divergence-prone layer: no second engine whose evaluated results must agree with Excel's cell-for-cell, and no exposure to dialect gaps (dynamic-array spill, `LAMBDA`/`LET`, newer functions, locale/volatile semantics).
- We give up **headless evaluation** — computing a formula's result without rendering onto an Excel surface (e.g. validating a branch in the background). Nothing in the current feature set needs this; the histogram is driven by Delta size, not computed values. If a concrete feature later requires headless eval, that is the trigger to revisit a second engine — not before.
