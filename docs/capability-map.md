# Capability Map — what the timeline can and cannot track

Living document. Maps each Excel feature to its **Fidelity Tier**, **capture mechanism**, **restore fidelity**, and **v1 status**. See [ADR-0009](./adr/0009-tiered-fidelity-honest-gaps.md) for the governing decision.

- **Capture mechanism**: `event` = observed via Office.js change events (fine-grained); `snapshot` = diffed from an object inventory at Step boundaries (coarse); `existence` = only presence is recorded.
- `(verify)` = the event/API behaviour is assumed but must be confirmed in the capture spike (ADR-0001/0004).

## Tier 1 — full fidelity, event-driven (v1)

| Feature | Capture | Restore fidelity | Notes |
|---|---|---|---|
| Cell values | event (`onChanged`) | full | `details` null for multi-cell — read back via `getValues` |
| Formulas (text) | event (`onChanged`) | full | stored alongside Frozen Value (lossless capture) |
| Number formats | event (`onChanged`/`onFormatChanged`) | full | |
| Font (name/size/color/bold/italic/underline) | event (`onFormatChanged`) | full | |
| Fill / background color | event (`onFormatChanged`) | full | |
| Borders | event (`onFormatChanged`) | full | |
| Alignment (h/v, wrap, indent) | event (`onFormatChanged`) | full | |
| Insert/delete row or column | event (`onChanged` structural type) | full | Structural Delta; native reference-shift |
| Worksheet add/delete/rename/reorder | event (worksheet collection + `onNameChanged`) | full | Worksheet Delta |
| Merged cells | event `(verify)` / snapshot fallback | full | merge/unmerge event firing unconfirmed |
| Row height / column width | snapshot `(verify)` | full | no reliable resize event |

## Tier 2 — config fidelity, snapshot-driven, best-effort (v1)

> **No fine-grained change events exist for any embedded object** (charts, pivots, shapes, slicers) — only collection add/delete + `onActivated`/`onDeactivated`. All Tier 2 objects are therefore captured by snapshotting an object inventory at Step boundaries, never observed mid-edit. (Verified against the API reference.)

| Feature | Capture | Restore fidelity | Notes |
|---|---|---|---|
| Conditional formatting | snapshot `(verify)` | good | rich rules may not round-trip exactly |
| Data validation | snapshot | good | |
| Defined names | snapshot | full | |
| Tables (ListObjects) — data | event (`Table.onChanged`) | full | data edits are Tier 1; style/structure Tier 2 |
| Tables — style/structure | snapshot | good | |
| Charts | snapshot (collection add/delete events only) | partial | rebuild from saved config; **chart source range is NOT a readable property** — must capture at creation, can't recover from an existing chart; custom formatting may drift |
| PivotTables — basic | snapshot | basic only | source range + row/col/data/filter hierarchies + layout; `showAs`/`PivotLayout`/`PivotFilters` reconstructable; **Data-Model/OLAP pivots cannot even be created** via `pivotTables.add` |
| Geometric shapes / text boxes | snapshot | good | round-trip well |
| Images | snapshot | existence/position only | **image bytes are NOT readable back** — cannot recreate byte-for-byte |
| Slicers (regular) | snapshot | good | `worksheet.slicers`/`slicers.add(source, field)`, caption/style/geometry/selected-items readable (ExcelApi 1.10) — **moved up from Tier 3** |
| Threaded comments | event (`comments.onChanged`) | good | |
| View settings (freeze panes, gridlines, zoom) | snapshot | good | |

## Tier 3 — existence-tracked only, restore NOT guaranteed (out of v1)

Rollback across these raises a **Fidelity Caveat** rather than corrupting silently.

| Feature | Why out |
|---|---|
| Data-model / OLAP / Power Pivot tables | cannot be created via `pivotTables.add` (range source only) |
| Pivot calculated fields / items, custom grouping | no API |
| Timeline slicers | no confirmed API (regular slicers are Tier 2) |
| Sparklines | no API |
| Form controls / ActiveX / OLE objects | no API |
| VBA / macros | not exposed to Office.js |
| External data connections / Power Query | not reconstructable |
| Linked data types (stocks, geography) | limited API |
| Legacy notes (non-threaded comments) | limited API |
| Worksheet/workbook protection & passwords | partial; security-sensitive |
