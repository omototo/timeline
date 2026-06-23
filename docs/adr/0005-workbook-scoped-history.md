# History is workbook-scoped (multi-sheet) from v1

## Context

The PRD used "sheet" and "workbook" interchangeably. Single-worksheet scope is far simpler (one coordinate space, per-sheet keyframes) but leaks for the target audience of financial analysts: cross-sheet references like `=Assumptions!B4` would be computed by Excel against the *current* values of untracked sheets, so previewing a historical state of one sheet would show today's other sheets feeding yesterday's sheet — an internally inconsistent past. Sheet add/delete/rename/reorder are also real mutations for these users.

## Decision

History is scoped to the whole workbook from v1. The Shadow State is a map of all worksheets; every Delta and keyframe carries a sheet dimension; sheet lifecycle operations (add/delete/rename/reorder) are first-class **Worksheet Deltas** within the Structural Delta class. The Timeline is zoomable on two axes — temporal (scrub across Steps) and structural (whole-workbook view drilling down into a single worksheet).

## Consequences

- No single-sheet shortcut for the MVP: the benchmark (ADR-0004) must size against 50K cells × N sheets, and keyframes snapshot all tracked sheets.
- Preview is internally consistent because the engine owns every sheet's history; cross-sheet references resolve against coherent same-Step state.
- Branching is workbook-wide — a branch forks the entire workbook, not one sheet.
