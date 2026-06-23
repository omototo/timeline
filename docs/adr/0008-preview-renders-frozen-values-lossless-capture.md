# Preview renders Frozen Values; capture is lossless (formula + value + type + format)

## Context

Writing a historical Step's formula text back to a sheet makes Excel recalculate it against the *present* context — `=TODAY()` shows today, `=Assumptions!B4` reads the current Assumptions value, `=RAND()` rerolls. So "show the past by replaying its formulas" actually shows the past's formulas evaluated in the present — a lie. The PRD treated this as a narrow volatile patch ("Value Freezing"); it is in fact every formula.

## Decision

Preview always renders the **Frozen Values** captured at each Step, never live formulas. Because the Preview Sheet is read-only, live recalculation there is unnecessary; formula text is surfaced as inert, inspectable metadata to satisfy "check formulas during preview." Live formulas exist only in the Present, where Excel is the calculation authority (ADR-0003).

Capture is **lossless**: each Step stores formula text, evaluated value, value type, and number format — the full superset, never a lossy subset.

## Consequences

- Volatiles, cross-sheet references, and general recalc drift are all handled by one rule (everything frozen), not special cases.
- Roughly doubles per-formula-cell storage (text + value). Accepted: it is the difference between a faithful time machine and a plausible fake, and it is the foundation for chosen extensibility.
- Lossless capture keeps future features (live recalc-in-preview, what-if branches promoted from a Preview, headless evaluation) reachable without touching the capture layer.
