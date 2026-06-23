# Office.js Investigation — Findings

Pre-build verification of every load-bearing Office.js assumption behind the architecture. **Two source tiers:**

- **Context7** (`/officedev/office-js-docs-pr`) — conceptual docs: confirms that APIs/events *exist* and their shapes, but does **not** carry enum membership, the multi-cell rules, or per-API requirement-set versions.
- **Microsoft Learn API reference** (fetched directly) — reference-grade: enum members, property remarks, and exact `[API set]` versions. Used to fill the gaps Context7 structurally cannot.

Verified 2026-06-23. Cluster sections (1–5) from the Context7 pass are appended below this summary.

---

## Executive summary — assumptions vs. verdict

| Assumption (ADR) | Verdict | Authority |
|---|---|---|
| Multi-cell `onChanged.details` is null; must diff a Shadow State (ADR-0001) | **CONFIRMED** — docs: *"If the changed event is triggered on multiple cells, this property cannot be retrieved."* | MS Learn |
| Structural changes surface as a typed event (`changeType`) (ADR-0001) | **CONFIRMED** — full `DataChangeType` enum at ExcelApi 1.7 | MS Learn |
| Excel natively shifts formula refs on insert → HyperFormula unneeded (ADR-0003) | **CONFIRMED (insert)** / **SPIKE (delete)** — docs state refs adjust on insert; delete-path `#REF!` not explicit | Context7 |
| `source` can't tell our writes from the user's → need echo cancellation (ADR-0002) | **CONFIRMED but SUPERSEDED** — see `triggerSource` below | both |
| Preview Sheet can be hidden so the user can't reveal it (ADR-0005/0008) | **CONFIRMED** — `SheetVisibility.veryHidden`, ExcelApi **1.1** | MS Learn |
| Workbook stamp travels in the file (ADR-0006) | **CONFIRMED** — `workbook.settings` (in-file, travels); `customXmlParts` for hardening | Context7 |
| Co-authoring is detectable to disable tracking (ADR-0006) | **NUANCED** — only reactively, via `source: Remote`; no proactive "is shared" flag | Context7 |
| Lossless capture of all cell types | **NUANCED** — `.values` flattens linked/entity types; true lossless needs `valuesAsJson` at **ExcelApi 1.16** | Context7 |

---

## Reference-grade verification (Microsoft Learn API reference)

These are the facts Context7 could not provide — fetched from the live API reference.

### `Excel.DataChangeType` enum — **ExcelApi 1.7**
Members (string values): `rangeEdited` ("RangeEdited"), `rowInserted`, `rowDeleted`, `columnInserted`, `columnDeleted`, `cellInserted`, `cellDeleted`, `unknown` ("Unknown" — was `"Others"` in 1.7, renamed in 1.8).
**Impact:** Structural Delta detection (ADR-0001) is fully supported at the **1.7** baseline. Distinguishing a value edit (`rangeEdited`) from a structural op is a documented switch on `changeType`.

### `WorksheetChangedEventArgs.details` — **ExcelApi 1.9**
Docs verbatim: *"This property can be retrieved when the changed event is triggered on a **single cell**. If the changed event is triggered on **multiple cells, this property cannot be retrieved**."*
**Impact:** Authoritatively confirms the Shadow State premise (ADR-0001). The single-cell `details` is an *optimization* for 1-cell edits; everything else requires read-back + diff.

### ⭐ `WorksheetChangedEventArgs.triggerSource` — **ExcelApi 1.14** (NEW — changes ADR-0002)
Type: `Excel.EventTriggerSource | "Unknown" | "ThisLocalAddin"`. Docs: *"identifies whether this local add-in triggers the event."*
**Impact:** This is a **direct echo signal Context7 never surfaced.** At 1.14+, an event raised by *our own* reconcile writes reports `triggerSource === "ThisLocalAddin"` — so echo cancellation can simply **drop events where `triggerSource` is `ThisLocalAddin`**, instead of maintaining an expected-write set. The expected-write set becomes the **fallback for hosts below 1.14**. Revises ADR-0002.

### ⭐ `WorksheetChangedEventArgs.changeDirectionState` — **ExcelApi 1.14** (NEW — strengthens ADR-0001)
Carries `insertShiftDirection` / `deleteShiftDirection` (mutually exclusive). 
**Impact:** For a Structural Delta, the event itself tells you the **shift direction** (down/right vs up/left). Combined with `changeType` + `address`, the Structural Delta is capturable directly from the event at 1.14+ — no inference needed.

### `WorksheetChangedEventArgs.getRange(ctx)` — **ExcelApi 1.8**
The changed range is obtainable directly from the event (no address re-parsing).

### `Excel.SheetVisibility` enum — **ExcelApi 1.1**
Members: `visible`, `hidden`, `veryHidden`. `veryHidden` cannot be unhidden from the Excel UI (only programmatically).
**Impact:** The engine-owned Preview Sheet can be `veryHidden` from the **lowest** baseline. Spike resolved.

---

## The requirement-set / `MinVersion` decision (now informed by hard facts)

The manifest `MinVersion` gates installability. The facts pin three meaningful floors:

| Target floor | What you get | What you give up |
|---|---|---|
| **1.7** | Structural `changeType`, `source`, worksheet lifecycle, `veryHidden` | single-cell `details`, direct echo signal, shift-direction, lossless entity types |
| **1.9** | + single-cell `details` (1-cell fast path) | direct echo signal, shift-direction, entity types |
| **1.14** (recommended floor for full design) | + `triggerSource` echo signal + `changeDirectionState` shift info | entity-type losslessness (1.16) |
| **1.16** | + `valuesAsJson` lossless entity/linked types | broadest reach |

**Recommendation:** manifest-gate at **1.9** (broad reach; `details` fast path), and **runtime feature-detect** `triggerSource`/`changeDirectionState` (1.14) and `valuesAsJson` (1.16) via `isSetSupported`, degrading gracefully (expected-write-set echo cancellation below 1.14; flatten entity cells with a Fidelity Caveat below 1.16). This keeps the install base wide while using the better APIs where present. **To be confirmed with you.**

---

## Spikes the docs could not resolve (must run before/with the keystone)

1. **Delete-path reference adjustment** — insert is doc-confirmed; delete `#REF!` semantics are not. (ADR-0003)
2. **Event fan-out of one user action** — sets the Step debounce window. (Step / ADR-0001)
3. **Echo arrival timing** below 1.14 — async ordering of our own write-echoes vs guard clear. (ADR-0002)
4. **Native Undo (Ctrl+Z)** — does it fire `onChanged`, with what `changeType`? (ADR-0006)
5. **IndexedDB persistence/eviction inside the add-in WebView** + `navigator.storage.persist()` — not in Office.js docs. (ADR-0007)
6. **`settings` byte cap** — trivial for a GUID+hash, but unbounded if we ever store more. (ADR-0006)

---
---

# Cluster sections (Context7 pass)



## 1. Events & Change Capture

> Sourcing note: All claims below trace to Context7 results against `/officedev/office-js-docs-pr` (the official Office.js docs PR repo). The Context7 corpus for this library is largely conceptual/sample-driven: it confirms the *existence and shape* of events well, but it does **not** expose the per-event requirement-set version numbers, the full `changeType` enum membership, or several runtime-behavior caveats. Those are marked **NOT FOUND IN DOCS — needs empirical spike** rather than guessed from memory.

---

### A1. `Worksheet.onChanged` populates `details.valueBefore/valueAfter` ONLY for single-cell edits; `details` null/undefined for multi-cell (paste) — NUANCED / NOT FULLY CONFIRMED
- Detail: Docs confirm `WorksheetChangedEventArgs.details` (a `ChangedEventDetail`) carries `valueBefore`, `valueAfter`, `valueTypeBefore`, `valueTypeAfter`, and that the sample handler reads them for "a changed cell" (singular). The same pattern holds for `Table.onChanged` → `TableChangedEventArgs.details`. However, the Context7 corpus does **NOT** contain any explicit statement that `details` is null/undefined for multi-cell changes such as a paste, nor a documented guard (e.g. `getDetailsOrNullObject`) for that case. The single-cell framing is strongly *implied* by every sample but not stated as a hard rule.
- Requirement set: ExcelApi for `onChanged` event itself = **NOT FOUND IN DOCS** (version not surfaced); the `details` property is a later addition than the base event — exact version **NOT FOUND IN DOCS**.
- Source: excel-add-ins-worksheets.md ("Log worksheet data changes with onChanged event", "Detect data changes"); excel-add-ins-tables.md ("Detect Table Data Changes").
- Implication: Treat `details` as present only for single-cell edits and code defensively (null-check) — but VERIFY empirically that multi-cell paste yields null `details`; our capture must fall back to a range snapshot for multi-cell changes regardless.

### A2. `onChanged` `changeType` full enum incl. structural values — NOT FOUND IN DOCS
- Detail: Docs confirm `onChanged` fires not only for value/format edits but "also applies to changes within tables and worksheets," and that it triggers "whenever the format or value of data changes." The Context7 results do **NOT** enumerate the `changeType` property values. The specific structural members (rangeEdited, rowInserted, rowDeleted, columnInserted, columnDeleted, cellInserted, cellDeleted) were not returned by any query.
- Requirement set: n/a (enum not surfaced).
- Source: excel-add-ins-events.md ("Events in Excel"); excel-add-ins-worksheets.md.
- Implication: **Needs empirical spike** — confirm the exact `changeType` string set against the live API / TypeScript typings before relying on structural-change discrimination in our capture logic.

### A3. `onChanged` `source` = Local vs Remote, and CANNOT distinguish add-in's own writes from the user's — NUANCED (Local/Remote CONFIRMED; own-write indistinguishability NOT FOUND)
- Detail: CONFIRMED that events carry a `source` property whose values are `Local` vs `Remote`, used for coauthoring: `event.source == Local` (current user) vs `event.source == Remote` (remote coauthor). This same `source` appears on `onFormulaChanged` and `onProtectionChanged` handlers too. CRITICALLY: `source` distinguishes *which user/machine*, NOT *user-vs-add-in*. The docs do **NOT** provide any property to tell the add-in's own programmatic writes apart from the local user's manual edits — both are `Local`. The explicit "it cannot" statement was not found, but the absence of any disambiguating field supports the assumption.
- Requirement set: coauthoring `source` semantics = **NOT FOUND IN DOCS** (version not surfaced).
- Source: excel-add-ins-events.md ("Events and coauthoring"); excel-add-ins-worksheets.md (onProtectionChanged handler reading `event.source`).
- Implication: Our own writes WILL echo back as `Local` onChanged events — we must use an explicit self-write guard/flag (and `suspendApiCalculationUntilNextSync` does not suppress events). Verify echo timing empirically (see A9).

### A4. `Worksheet.onFormatChanged` — existence and which-property detail — CONFIRMED (existence) / NUANCED (no property-level detail found)
- Detail: CONFIRMED `onFormatChanged` exists and "occurs when the formatting of a worksheet is modified." The docs do **NOT** surface any payload carrying *which* format property changed; no `FormatChangedEventArgs` property breakdown was returned. Note also: the base `onChanged` event docs say it fires "whenever the format or value of data changes," so format edits may surface on `onChanged` too.
- Requirement set: **NOT FOUND IN DOCS**.
- Source: excel-add-ins-events.md ("Events in Excel").
- Implication: `onFormatChanged` tells us *that* formatting changed, likely not *what* — assume a format snapshot/diff is required to know specifics.

### A5. `Worksheet.onCalculated` — existence and trigger — CONFIRMED
- Detail: CONFIRMED. `onCalculated` "signals that a worksheet, or all worksheets in a collection, has finished its calculation process." It is available both per-worksheet and at the worksheet-collection level. Triggered by the recalc cascade completing.
- Requirement set: **NOT FOUND IN DOCS**.
- Source: excel-add-ins-events.md ("Events in Excel").
- Implication: Use `onCalculated` as the "recalc settled" signal to re-snapshot formula-driven values after edits cascade; it fires after calculation completes, not per formula.

### A6. Worksheet-collection events: onAdded, onDeleted, onActivated, onDeactivated, onNameChanged (rename) — CONFIRMED (existence) / NUANCED (payloads partial)
- Detail: CONFIRMED `onAdded` "is triggered when a new object is added to a collection... worksheets" (also charts, comments, tables). CONFIRMED `onNameChanged` "fires when the name of a worksheet is altered" (sheet rename). The events reference also lists `onMoved` (worksheet moved), `onVisibilityChanged`, `onSelectionChanged`, `onActivated`/`onDeactivated` are implied by the collection model but `onActivated`/`onDeactivated`/`onDeleted` were not individually quoted in returned snippets. Payload shapes (e.g. `WorksheetAddedEventArgs.worksheetId`, `WorksheetNameChangedEventArgs` old/new name) = **NOT FOUND IN DOCS** in detail, though `onProtectionChanged` is shown carrying `worksheetId` + `source`, suggesting the worksheet-event family carries `worksheetId`.
- Requirement set: **NOT FOUND IN DOCS** per event.
- Source: excel-add-ins-events.md ("onAdded", "onNameChanged", "Excel JavaScript API Event Reference" listing onMoved/onNameChanged/onVisibilityChanged/onSelectionChanged/etc.).
- Implication: Sheet add/rename/move are observable via collection events carrying at least `worksheetId`; confirm exact added/deleted/activated payloads empirically. Rename is a first-class event (good — no polling needed).

### A7. Table events (`Table.onChanged`) and comment events (`comments.onChanged`) — CONFIRMED
- Detail: CONFIRMED `Table.onChanged` exists, receives `TableChangedEventArgs` with a `details` (`ChangedEventDetail`) exposing before/after values+types — same shape as worksheet onChanged. CONFIRMED comment events live on the `CommentCollection` object: `onAdded`, `onChanged`, `onDeleted`. `onChanged` also covers comment reply add/change/delete and resolve/reopen. Comment event args (`CommentAddedEventArgs`, `CommentChangedEventArgs`, `CommentDeletedEventArgs`) carry **arrays of comment IDs** (`event.commentDetails[].commentId`). Important: "Each comment event only triggers once when multiple additions/changes/deletions are performed at the same time" — i.e. batched.
- Requirement set: **NOT FOUND IN DOCS**.
- Source: excel-add-ins-tables.md ("Detect Table Data Changes"); excel-add-ins-comments.md ("Comment events", onChanged/onDeleted samples).
- Implication: Comment capture must iterate the `commentDetails[]` array (batched, multi-ID) — not assume single. Table changes mirror worksheet onChanged handling.

### A8. Row/column RESIZE, cell MERGE/unmerge, conditional-format edits fire NO reliable event (snapshot needed) — NOT FOUND IN DOCS (no such events surfaced → supports assumption)
- Detail: No event for row/column *resize*, cell *merge/unmerge*, or conditional-format edits was returned by any query. The events reference does list `onRowHiddenChanged` (hidden state) and `onRowSorted`, but NOT a resize or merge event. Conditional formatting has no change event in the returned corpus. Absence in docs is consistent with the assumption that these require a snapshot/diff, but absence-of-evidence is not proof.
- Requirement set: n/a.
- Source: excel-add-ins-events.md ("Excel JavaScript API Event Reference" — lists onRowHiddenChanged, onRowSorted, onVisibilityChanged, etc.; no resize/merge/CF event).
- Implication: Plan for snapshot-based detection of resize, merge/unmerge, and conditional-format changes — VERIFY empirically that none of these surface as `onFormatChanged` or `onChanged` side effects.

### A9. Event handler async timing — handlers run async after context.sync; own-write echoes may arrive after guard clears — NOT FOUND IN DOCS (behavior not documented in returned corpus)
- Detail: Every handler sample wraps work in `Excel.run`/`context.sync`, confirming handlers are async and load data via their own sync round-trip. But the Context7 corpus does **NOT** document the firing order relative to the originating `context.sync`, nor whether own-write echo events can be delivered after a synchronously-cleared guard flag. `suspendApiCalculationUntilNextSync()` is documented but it suspends *recalculation*, not *event delivery* — so it will NOT prevent onChanged echoes.
- Requirement set: n/a.
- Source: excel-add-ins-events.md (handler registration sample); performance.md / excel-add-ins-workbooks.md (`suspendApiCalculationUntilNextSync`).
- Implication: **Needs empirical spike.** Assume echoes can arrive asynchronously after our write's sync completes; a synchronous boolean guard is unsafe. Design the self-write guard to span until the echo is observed/drained, or correlate by address+value.

### A10. Native Undo (Ctrl+Z) fires onChanged with reverted values — NOT FOUND IN DOCS
- Detail: No statement in the returned corpus about whether Ctrl+Z / native undo raises `onChanged` (or with what `changeType`/values). Not addressed.
- Requirement set: n/a.
- Source: (none — not found).
- Implication: **Needs empirical spike.** Likely undo fires onChanged with the reverted value (since the cell value genuinely changes), which would also echo through our guard — must be tested.

### A11. Per-event requirement-set version (ExcelApi x.y) that introduced each event — MOSTLY NOT FOUND IN DOCS
- Detail: The Context7 corpus surfaced the *mechanism* for requirement sets (`Office.context.requirements.isSetSupported('ExcelApi','1.7')`, manifest min/maxVersion, "each version is a superset of earlier versions") but did **NOT** return the introduction version for any individual event (onChanged, onFormatChanged, onCalculated, onNameChanged, comment events, etc.). Only incidental version strings appeared (1.6, 1.7, 1.10, 1.16) in unrelated config examples.
- Requirement set: per-event versions = **NOT FOUND IN DOCS** — must be read from the Excel requirement-sets reference / TypeScript typings (`@types/office-js`) directly.
- Source: develop/understand-requirement-configuration.md; develop/initialize-add-in.md; develop/specify-office-hosts-and-api-requirements-unified.md.
- Implication: We cannot pin a defensible minimum ExcelApi from Context7 alone — see note below; treat the figure as provisional pending the requirement-sets reference.

---

### Minimum ExcelApi version this cluster needs: 1.9 (PROVISIONAL — NOT FOUND IN DOCS)

> The Context7 corpus did not surface the introduction version for these events. Based on the *feature set* required (worksheet `onChanged` with the `details`/before-after `ChangedEventDetail`, `onFormatChanged`, `onCalculated`, worksheet-collection `onNameChanged`/`onAdded`, table `onChanged`, comment collection events), the binding constraint is the richest of these — the `onChanged` `details` payload and comment events, historically the latest additions in this group. **1.9 is a provisional placeholder and must be confirmed against the Excel requirement-sets reference / `@types/office-js` typings via an empirical spike.** Do not ship a manifest min-version off this number without verification.


## 2. Reading State

> Docs-grounded against Context7 library `/officedev/office-js-docs-pr`. Claims not traceable to a doc result are marked "NOT FOUND IN DOCS — needs empirical spike." Note: Context7 does not surface per-property `[ApiSet]` requirement-set tags in its snippets, so most ExcelApi version numbers below are marked "not stated in docs (Context7)" and need confirmation against the official `@types/office-js` / API reference.

### Bulk read of values, formulas, number formats, types, text (`Range.values`, `.formulas`, `.formulasR1C1`, `.numberFormat`, `.valueTypes`, `.text`) — CONFIRMED (partial)
- Detail: Docs confirm `Range.values`, `Range.formulas`, and `Range.text` are loadable properties returned as 2D arrays (examples serialize them with `JSON.stringify(range.values, null, 4)` etc.). `text` returns the formatted display string regardless of whether the cell holds a value or formula. `formulas` returns the formula, falling back to the raw value when a cell has no formula. These all index `[row][col]`.
- `.formulasR1C1`, `.numberFormat`, and `.valueTypes` were NOT surfaced in the Context7 results in this pass, though they are well-established sibling properties of the same shape. Treat their existence/2D-array shape as: NOT FOUND IN DOCS (this pass) — high confidence from API design, confirm against API reference / spike.
- Requirement set: ExcelApi 1.1 for values/formulas/text/numberFormat (not stated in returned docs; standard baseline). `valuesAsJson`-family is newer (see below).
- Source: `docs/excel/excel-add-ins-ranges-set-get-values.md` (values, text, formulas examples).
- Implication: Lossless value/formula/display capture is achievable in a single `load(["values","formulas","text","numberFormat","valueTypes"])` + one `sync()`. Each is an independent 2D array, so a shadow snapshot = parallel 2D arrays keyed by address.

### Richer value API: `Range.valuesAsJson` / linked & entity data types — CONFIRMED
- Detail: `Range.valuesAsJson` exists and is read/write. It exposes typed cell values via `CellValueType` (e.g. `double`, `string`, `Entity`, plus a `basicType`/`basicValue` and a `properties` bag). Entity/linked data types carry nested structured `properties` (e.g. `EntityCellValue` with `.text` and `.properties[attr].basicValue`). This captures strictly MORE than `.values`: `.values` flattens a rich/entity/linked cell to its basic scalar and loses the `properties`, `type`, and nested structure.
- Requirement set: ExcelApi 1.16 for `valuesAsJson`/data types (not stated in returned docs — needs confirmation; data types GA'd around ExcelApi 1.16).
- Source: `docs/excel/excel-data-types-add-properties-to-basic-cell-values.md` (`range.valuesAsJson = [[{type, basicType, basicValue, properties}]]`); `docs/excel/custom-functions-data-types-concepts.md` (`EntityCellValue`, `.properties[attribute].basicValue`).
- Implication: For lossless capture, prefer `valuesAsJson` over `values` for any workbook that may contain linked/entity/web-image/formatted-number-value cells. `values` alone is lossy for these. There is also a `valuesAsJsonLocal` variant (locale-aware) — NOT FOUND IN DOCS this pass; confirm.

### Format surface for lossless capture (`Range.format` → font/fill/borders/alignment/protection) — CONFIRMED (partial enumeration)
- Detail confirmed readable/settable via `Range.format`:
  - Font: `format.font.bold`, `.italic`, `.size`, `.name`, `.color`.
  - Fill: `format.fill.color`.
  - Alignment: `format.horizontalAlignment` (`Excel.HorizontalAlignment.*`), `format.verticalAlignment` (`Excel.VerticalAlignment.*`).
  - Docs explicitly state `Range.format` is the gateway to "font styles, background fills, cell borders, and alignment settings," and that format changes do NOT alter underlying data.
- NOT surfaced in this pass (existence high-confidence from API design, but unconfirmed by returned docs — mark NOT FOUND IN DOCS, spike to enumerate): font `.underline` (enum) / `.strikethrough`; per-edge `format.borders.getItem(edge)` → `.style`/`.color`/`.weight`; `format.wrapText`; `format.indentLevel`; `format.textOrientation`; `format.rowHeight`/`columnWidth`; `format.protection` (`.locked`, `.formulaHidden`).
- Requirement set: font/fill/alignment ExcelApi 1.1; `format.protection` ExcelApi 1.2 (not stated in returned docs).
- Source: `docs/excel/excel-add-ins-ranges-set-format.md`.
- Implication: Core format capture is well supported, but a true *lossless* format snapshot needs an empirical spike to enumerate the full readable property tree (borders are per-edge via a collection, not flat; underline/protection/wrap/indent unconfirmed here).

### Conditional formats: read existing rules & types (`Range.conditionalFormats`) — NUANCED
- Detail: `Range.conditionalFormats` is a collection. Docs show creating rules (`.add(Excel.ConditionalFormatType.cellValue|colorScale|dataBar|iconSet|...)`) and accessing per-type config objects (`.cellValue.rule = {formula1, operator}`, `.colorScale.criteria`, `.dataBar.barDirection`, etc.). Reading is implied by the same typed accessors (load `conditionalFormats` then inspect `.type` and the matching typed sub-object), but the returned docs only demonstrate the WRITE/create path — no explicit "load existing rules and enumerate them" example was surfaced.
- Limits (from docs + API shape): the conditional-format object is a **discriminated union by `type`** — you must read `.type` first, then read only the matching sub-property (`.cellValue`, `.colorScale`, `.dataBar`, `.iconSet`, `.preset`, `.textComparison`, `.topBottom`, `.custom`). The non-matching sub-properties are null. Rule `formula1`/`formula2` are strings.
- Requirement set: ExcelApi 1.6 (conditional formatting introduced; not stated in returned docs).
- Source: `docs/excel/excel-add-ins-conditional-formatting.md`.
- Implication: Reading existing CF rules is possible but requires a type-switch per format; confirm via spike that loading `conditionalFormats/type` + each typed sub-object round-trips existing (not just newly added) rules.

### Data validation: readable (`Range.dataValidation`) — CONFIRMED
- Detail: `Range.dataValidation` (a `DataValidation` object) has five properties, several readable: `rule` (`DataValidationRule` — WholeNumber/Date/TextLength/List/Custom etc., with `formula1`/`formula2`/`operator` or `custom.formula`), `errorAlert` (message/title/style), `prompt` (message/title/showPrompt), `ignoreBlanks`, and `type` — explicitly described as a **read-only** identification of the validation type (WholeNumber, Date, TextLength, …), set indirectly when `rule` is set.
- Requirement set: ExcelApi 1.8 (data validation; not stated in returned docs).
- Source: `docs/excel/excel-add-ins-data-validation.md` ("Programmatic control of data validation").
- Implication: Data validation is fully capturable for a shadow: load `dataValidation/rule`, `/type`, `/errorAlert`, `/prompt`, `/ignoreBlanks`.

### Defined names: readable (`workbook.names` / `worksheet.names`) — NUANCED
- Detail: Named items are usable (docs resolve a named range via `sheet.getRange("MyRange")`). The `NamedItemCollection` on `workbook.names` and `worksheet.names`, and per-`NamedItem` properties (`name`, `formula`, `type`, `value`, `scope`, `visible`, `comment`), were NOT directly surfaced in this pass.
- Requirement set: ExcelApi 1.1 (workbook.names; worksheet-scoped names 1.4) — not stated in returned docs.
- Source: `docs/excel/excel-add-ins-ranges-get.md` (named-range resolution only).
- Implication: NOT FOUND IN DOCS (this pass) for the readable property list of `NamedItem` — high confidence these are loadable (`name`/`formula`/`scope`/`type`/`visible`), but confirm formula+scope readability via spike or API reference before relying on it for lossless capture.

### Used range & sparse/non-contiguous reads (`Worksheet.getUsedRange`, `RangeAreas`) — CONFIRMED
- Detail: `Worksheet.getUsedRange()` returns the bounded used range and is the recommended way to scope reads (docs call it first to limit `getSpecialCells` scope). `RangeAreas` represents multiple discontiguous ranges: `sheet.getRanges("A1:A100, C1:C100")` and `usedRange.getSpecialCells(Excel.SpecialCellType.formulas)` both return `RangeAreas`, enabling formatting/reads over non-contiguous cells in one operation. `SpecialCellType.formulas` filters to formula cells specifically.
- Requirement set: `getUsedRange` ExcelApi 1.1; `RangeAreas`/`getSpecialCells`/`getRanges` ExcelApi 1.9 (not stated in returned docs).
- Source: `docs/excel/excel-add-ins-ranges-special-cells.md`; `docs/excel/excel-add-ins-multiple-ranges.md`.
- Implication: Use `getUsedRange()` to avoid snapshotting the full 1M×16k grid, and `RangeAreas` + `getSpecialCells` to target only formula/constant/blank cells — key for staying under the cell/payload limits.

### HARD LIMITS: 5 MB payload, 5,000,000-cell read, `range.cellCount`, `untrack()` — CONFIRMED (mostly)
- Detail (exact numbers from docs):
  - **Payload size limit: 5 MB** for BOTH requests and responses, **on Excel on the web specifically**. Exceeding it throws a `RichAPI.Error` / `RequestPayloadSizeLimitExceeded` error. Total payload is a function of number of API calls, number of objects, and length of values set/retrieved.
  - **Read/get cell limit: 5,000,000 cells**, and this applies to **get operations on ALL platforms** (not web-only). The 5 MB payload cap is the web-only one.
  - `range.cellCount` is loadable and is the documented pre-flight check: load `cellCount`, and if `> 5000000`, split the operation before `sync()`.
  - Recommended chunking pattern: break large operations into smaller sequential batches, calling `context.sync()` after each sub-operation (docs example splits rows 1–5000 then 5001–10000 into two `sync()` calls); validate sizes BEFORE `sync()`; use `RangeAreas` to target only needed cells.
  - `untrack()`: NOT FOUND IN DOCS (this pass) — the performance docs returned did not include the `untrack()`/proxy-object memory guidance. High confidence it exists (`ClientObject.untrack()` to release proxy objects and reduce memory in large loops), but confirm against `docs/excel/performance.md` / `docs/concepts/...` directly.
- Requirement set: n/a (platform limits, not gated by ExcelApi version).
- Source: `docs/concepts/resource-limits-and-performance-optimization.md`; `docs/excel/performance.md`.
- Implication: Snapshotting must pre-check `cellCount` (cap 5,000,000 cells per get, all platforms) AND keep each request/response under 5 MB on web — these are two distinct ceilings. For wide lossless captures (values + formulas + text + numberFormat + valuesAsJson all at once) the 5 MB web cap will bite far below 5M cells, so chunk by region and by property. Add `untrack()` in long shadow-maintenance loops (verify API).

### "Before" state of a MULTI-cell change from the event itself — REFUTED (for multi-cell) / CONFIRMED only for single-cell
- Detail: `Worksheet.onChanged` / `Table.onChanged` deliver a `WorksheetChangedEventArgs` / `TableChangedEventArgs` whose `.details` (`ChangedEventDetail`) carries `valueBefore`, `valueAfter`, `valueTypeBefore`, `valueTypeAfter` — i.e. the before-state IS available, **but only for a single changed cell**. Multiple docs explicitly caveat this: the formula-change handler notes "This method assumes only a single formula is changed at a time," and the `details` examples read scalar before/after (not arrays). When a change spans multiple cells (paste, fill, multi-select edit), `details` does not provide a 2D before-array; the docs do not expose any per-cell "before" grid on the event.
- Requirement set: `details`/before-after on changed event ExcelApi 1.9 (`onChanged` is 1.7; the before/after `details` are later — not stated in returned docs, confirm).
- Source: `docs/excel/excel-add-ins-worksheets.md` (onChanged before/after; onFormulaChanged single-formula caveat); `docs/excel/excel-add-ins-tables.md`.
- Implication: For lossless before-state of multi-cell changes you CANNOT rely on the event payload — you MUST maintain a shadow snapshot and diff against it on each change. The event's `valueBefore` is only trustworthy for single-cell edits and even then is the scalar value, not formula/format. Confirmed: the suspicion is correct.

### Minimum ExcelApi version this cluster needs: 1.16
- Rationale: baseline reads (values/formulas/text/numberFormat/format/getUsedRange) are 1.1; data validation 1.8; `RangeAreas`/`getSpecialCells` + onChanged `details` 1.9; conditional formatting 1.6 — but **lossless** value capture requires `valuesAsJson` / linked & entity data types, which is the newest dependency at ~**ExcelApi 1.16**. All version numbers above are inferred (Context7 snippets did not include requirement-set tags) and should be confirmed against the official API reference before finalizing the minimum.


## 3. Writing, Reconcile & Structural Ops

Docs-grounded against Context7 library `/officedev/office-js-docs-pr`. Every claim traces to a doc URL or is marked NOT FOUND.

### Bulk write of values/formulas/formats, batched before one sync — CONFIRMED
- Detail: `Range.values`, `Range.formulas`, `Range.numberFormat` are all settable as 2-D arrays. Format props (`range.format.fill.color`, `range.format.font.color`, etc.) are set on proxy objects. All mutations are queued on proxy objects and only applied to the document when `context.sync()` runs the batched command set — so many mutations can be staged before a single sync. `null` in a 2-D array means "leave this cell's property unchanged" (e.g. `range.numberFormat = [[null, null, null, 'm/d/yyyy;@']]`), which is useful for sparse reconcile writes.
- Requirement set: ExcelApi 1.1 (core Range values/formulas/numberFormat + sync model).
- Source: develop/application-specific-api-model.md (proxy queue + sync); excel/excel-add-ins-ranges-set-get-values.md; excel/excel-add-ins-ranges-set-format.md; excel/excel-add-ins-blank-null-values.md (null = no-op); excel/performance.md.
- Implication: Our reconcile loop can stage all value/formula/format diffs across many ranges and flush with one sync. Use `null` to avoid clobbering format on value-only writes.

### untrack() and large-batch write performance + 5MB / 5,000,000-cell limits — CONFIRMED
- Detail: For large per-cell write loops, call `cell.untrack()` after staging each cell to release the proxy from tracked memory — docs state this gives "noticeable performance improvements" on large batch operations. Hard limits: **5MB payload** per request AND per response on Excel **on the web** (throws `RichAPI.Error` if exceeded); **5,000,000-cell limit on read operations** for a range. Mitigations: validate before sync (e.g. load `range.cellCount` and check `> 5000000`), split large operations into chunks each with its own `context.sync()`, and use `RangeAreas` to target scattered cells instead of one giant contiguous range. Note: the explicit 5M-cell limit is documented for READ operations; for writes the governing documented constraint is the 5MB payload size, so very large writes should be chunked by payload, not just cell count.
- Requirement set: `untrack()` ExcelApi 1.3; `RangeAreas` ExcelApi 1.9; `cellCount` ExcelApi 1.4. (Limits are platform constraints, not versioned.)
- Source: concepts/resource-limits-and-performance-optimization.md (limits, untrack, cellCount check, RangeAreas); excel/performance.md (chunked sync).
- Implication: Big reconcile batches must be payload-chunked for web (5MB) and should untrack per-cell when looping. Prefer bulk 2-D array writes over per-cell loops where possible; reserve untrack for unavoidable per-cell loops.

### Structural ops: Range.insert / Range.delete (entire rows/columns) — CONFIRMED, with native formula-reference shifting CONFIRMED
- Detail: `Range.insert(shiftDirection)` takes `Excel.InsertShiftDirection` (`down` | `right`) and adds empty cells, shifting existing cells. `Range.delete(shift)` takes `Excel.DeleteShiftDirection` (`up` | `left`) and removes cells, shifting survivors to fill the gap. Inserting/deleting entire rows or columns is done by addressing a full-row/full-column range and using the same methods (full row shifts down/up, full column shifts right/left). **Native formula-reference shifting is explicitly documented**: the insert reference doc's Key Points state "Formulas with cell references automatically adjust after insertion." (This is the load-bearing claim for ADR-0003 / cutting HyperFormula.)
- Caveat / spike note: The explicit "formulas automatically adjust" sentence is documented for the INSERT path. The DELETE doc describes the shift-to-fill behavior but I did not find an equally explicit "references adjust" sentence for delete in the retrieved snippets. This is standard Excel behavior (delete shifts references and turns deleted-cell refs into `#REF!`), but to be docs-honest: insert-side reference shifting = CONFIRMED in docs; delete-side reference adjustment = NUANCED / recommend a small empirical spike to confirm reference rewrite + `#REF!` semantics on delete before relying on it in reconcile replay.
- Requirement set: ExcelApi 1.1 (`Range.insert`, `Range.delete`, both shift enums).
- Source: excel/excel-add-ins-ranges-insert.md (insert + "Formulas with cell references automatically adjust after insertion"); excel/excel-add-ins-ranges-clear-delete.md (delete + shift).
- Implication: Validates replaying structural ops natively without a JS formula engine for inserts. Add one spike to confirm delete-path reference adjustment / `#REF!` handling, then ADR-0003 is fully grounded.

### Worksheet lifecycle: add / delete / rename / reorder / copy — CONFIRMED
- Detail: Add via `worksheets.add(name?)` (returns the new `Worksheet`, has a `position`). Delete via `Worksheet.delete()` — cannot delete the only remaining sheet (guard by checking `sheets.items.length === 1`). Rename via assignment `worksheet.name = "New Name"`. Reorder via assignment `worksheet.position = 0` (0-based). Copy via `worksheet.copy(positionType, relativeTo)` using `Excel.WorksheetPositionType` (e.g. `after`) and a reference sheet. All applied on `context.sync()`.
- Requirement set: add/delete/rename/reorder ExcelApi 1.1; `Worksheet.copy` ExcelApi 1.7.
- Source: excel/excel-add-ins-worksheets.md (add, delete, rename, move/position, copy).
- Implication: Engine can fully own sheet topology. `copy` (1.7) is the version floor if we duplicate sheets for branch/preview snapshots.

### Worksheet visibility incl. veryHidden — CONFIRMED (visible/hidden), NUANCED (veryHidden)
- Detail: `Worksheet.visibility` is set with the `Excel.SheetVisibility` enum; docs explicitly show `Excel.SheetVisibility.hidden` and `Excel.SheetVisibility.visible`. The `veryHidden` member (a sheet hidden such that the user cannot unhide it from the normal Excel right-click/unhide UI — only programmatically or via VBA/XML) is a real `SheetVisibility` member in the API, but I did NOT find an explicit doc snippet in the retrieved Context7 results naming `Excel.SheetVisibility.veryHidden` or stating its requirement set. There is also an `onVisibilityChanged` worksheet event.
- Requirement set: `SheetVisibility.visible`/`hidden` ExcelApi 1.1; `veryHidden` member — NOT FOUND IN RETRIEVED DOCS (believed ExcelApi 1.2, needs confirmation).
- Source: excel/excel-add-ins-worksheets.md (visible/hidden); excel/excel-add-ins-events.md (onVisibilityChanged). veryHidden: NOT FOUND IN DOCS — needs confirmation/empirical spike.
- Implication: An engine-owned Preview Sheet that the user can't trivially reveal is achievable IF `veryHidden` is supported at our target requirement set. Spike: confirm `Excel.SheetVisibility.veryHidden` exists and its min ExcelApi version on our target hosts before depending on it; fall back to `hidden` (user-revealable) otherwise.

### Suspend recalculation / screen updating for fast bulk reconcile — CONFIRMED
- Detail: `Application.calculationMode` is readable (loads as e.g. "Automatic"). `Application.suspendApiCalculationUntilNextSync()` suspends recalc until the next `context.sync()` — docs show formulas NOT recalculating mid-batch (e.g. dependent `=SUM` stays stale until the sync after resume), giving a real perf win for bulk edits where intermediate computed values aren't needed. `Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates until next sync / end of `Excel.run`; docs warn NOT to call it repeatedly (e.g. in a loop) or the window flickers, and to show the user a progress indicator since the UI looks idle.
- Requirement set: `calculationMode` ExcelApi 1.1; `suspendApiCalculationUntilNextSync` ExcelApi 1.6; `suspendScreenUpdatingUntilNextSync` ExcelApi 1.9.
- Source: excel/performance.md (both suspend APIs + behavior/warnings); excel/excel-add-ins-workbooks.md (suspendApiCalculationUntilNextSync).
- Implication: Reconcile should wrap bulk writes in suspendApiCalculation (1.6) and optionally suspendScreenUpdating (1.9, call once, show progress). `suspendScreenUpdatingUntilNextSync` at 1.9 is the highest floor in this cluster if we use it.

### Writing values via the API fires onChanged; no suppress flag (echo cancellation needed) — CONFIRMED (with nuance)
- Detail: `Worksheet.onChanged` fires whenever the value OR format of data changes; it delivers `WorksheetChangedEventArgs` with `changeType`, `address`, and `details` (valueBefore/valueAfter/type). Docs do NOT document any flag to suppress the event when the change originates from your own add-in's writes — there is no "silent write" option. The available discriminator is `event.source`, which is `Local` vs `Remote` (Local = current user/this client, Remote = a coauthor). NOTE: `source` distinguishes this-client vs coauthor — it does NOT by itself distinguish "my add-in wrote it" from "the user typed it," since both are Local. So echo cancellation against our own programmatic writes must be done by US (e.g. set an in-flight write flag / diff against expected values around our sync), not by a built-in suppress.
- Requirement set: `onChanged` ExcelApi 1.7; `details` (WorksheetChangedEventArgs) ExcelApi 1.9; `event.source` (Local/Remote) ExcelApi 1.8.
- Source: excel/excel-add-ins-worksheets.md (onChanged, details, "verify if the actual data values were altered"); excel/excel-add-ins-events.md (event.source Local/Remote, coauthoring); excel/co-authoring-in-excel-add-ins.md. No suppress flag: NOT FOUND IN DOCS (confirms none exists).
- Implication: Our echo-cancellation design is required and correct. We cannot rely on `event.source` alone to filter our own writes (our writes are Local, same as user edits). Implement an explicit "suppress next onChanged for these addresses" window around our reconcile sync.

### Minimum ExcelApi version this cluster needs: 1.9
- Rationale: core writes/structural/lifecycle/visibility/suspend-calc all sit at 1.1–1.7, but the features we actually want pull the floor up: `suspendScreenUpdatingUntilNextSync` and `RangeAreas` and onChanged `details` are 1.9. If we drop screen-update suspension, RangeAreas, and change-details, the floor falls to 1.7 (driven by `Worksheet.copy` and `onChanged`). Target 1.9 for the full feature set; 1.7 is the hard minimum for a reduced set. (Plus: confirm `SheetVisibility.veryHidden` version separately — likely 1.2, currently unverified.)


## 4. Embedded Objects (Charts/Pivots/Shapes)

Docs-grounded validation against Context7 library `/officedev/office-js-docs-pr`. Claims trace to a doc source; gaps are marked "NOT FOUND IN DOCS — needs empirical spike."

### Charts — tier CONFIRMED (Tier 2)
- Read fidelity: Good config snapshot. Readable: `chart.type`, `chart.name`, position (`top`/`left`/`height`/`width`), `title.text`, `legend.position`/`legend.format`, `dataLabels.format`, `axes.categoryAxis`/`valueAxis`, `series` collection (per-series name, format, points), `getDataTableOrNullObject()`. Source data range is NOT directly readable as a property — `charts.add(type, sourceData, seriesBy)` consumes a range to build the chart, but the API exposes no `chart.getSourceData()`/source-range getter. So you reconstruct the *rendered* config, not necessarily the live source binding. `chart.getImage()` gives a raster fallback.
- Recreate fidelity: `sheet.charts.add(Excel.ChartType.<x>, dataRange, Excel.ChartSeriesBy.<x>)` then re-apply title/legend/dataLabels/axes/series names/position. Round-trips the common config; custom per-point formatting, trendlines, secondary axes, and exotic chart types may drift or be unsettable.
- Change event: NO data/config-change event. Only collection-level `onAdded`/`onDeleted` (charts are explicitly listed for `onDeleted`) and object-level `onActivated`/`onDeactivated`. There is NO "chart changed" event — edits to an existing chart's formatting/series are invisible except via re-snapshot. Suspicion CONFIRMED.
- Requirement set: chart core ExcelApi 1.1; richer formatting/series/axes/dataLabels/dataTable across 1.7–1.9; `chart.getImage` 1.9. (Exact per-property minimums NOT enumerated in retrieved snippets — needs requirement-set table cross-check.)
- Source: docs/excel/excel-add-ins-charts.md; excel-add-ins-charts-data-labels.md; docs/tutorials/excel-tutorial.md; docs/excel/excel-add-ins-events.md.
- Implication for capability-map.md: Current row ("snapshot; collection add/delete events only; rebuild from saved config; custom formatting may drift") is ACCURATE. Add a note that the chart *source data range* is not a readable property — capture it at creation time or infer it; this is a real fidelity gap worth flagging in the row's Notes.

### PivotTables — tier NUANCED (basic = Tier 2 CONFIRMED; advanced = Tier 3 CONFIRMED)
- Read fidelity: Good for structure. Readable/settable: `pivotTables` collection, `hierarchies`, `rowHierarchies`/`columnHierarchies`/`dataHierarchies`/`filterHierarchies`, per-data-hierarchy `showAs` (ShowAsCalculation incl. percentOfColumnTotal etc.), `summarizeBy`, PivotFilters (value/label/date/manual via `applyFilter`), and `layout` (`layoutType` Compact/Outline/Tabular, `emptyCellText`, `fillEmptyCells`, `preserveFormatting`, `showRowGrandTotals`/`showColumnGrandTotals`, subtotals).
- Recreate fidelity: BASIC pivots reconstruct well — `worksheet.pivotTables.add(name, sourceRange, destination)` then re-add the four hierarchy categories, showAs, filters, and layout. NOT reconstructable (no API — Tier 3): calculated fields, calculated items, custom grouping (date/numeric bucketing), and any pivot backed by the Data Model / OLAP / Power Pivot / external connection. `pivotTables.add` only takes a worksheet range as source, so Data-Model/OLAP pivots cannot be created or faithfully round-tripped at all. value-field "summarize by" and "show values as" ARE covered (`summarizeBy`, `showAs`), so that part of value-field settings is supported — but advanced number-format-per-field and calculated measures are not.
- Change event: NO pivot-specific change event. PivotTable edits surface only as `Worksheet.onChanged` over the spilled cell area (coarse, value-level), not as a structural pivot event, plus collection `onAdded`/`onDeleted`. Suspicion CONFIRMED.
- Requirement set: PivotTable management/hierarchies ExcelApi 1.8; filters, showAs, layout/PivotLayout, dataHierarchy details ExcelApi 1.9+. (Exact minimums per member NOT in retrieved snippets — needs requirement-set table cross-check.)
- Source: docs/excel/excel-add-ins-pivottables.md; docs/excel/excel-add-ins-events.md.
- Implication for capability-map.md: Tier 2 "PivotTables — basic" and Tier 3 "data-model/OLAP/Power Pivot" + "calculated fields/items, custom grouping" splits are CORRECT. Refine the Tier 2 Notes: explicitly include showAs + PivotLayout + filters as reconstructable (not merely "row/col/data/filter hierarchies + layout"), and state that Data-Model-backed pivots cannot even be created via `pivotTables.add` (source is range-only).

### Shapes / images / text boxes — tier CONFIRMED (Tier 2), with a read-back caveat
- Read fidelity: Create-side is well documented; read-back is partial. `worksheet.shapes` collection supports `addGeometricShape(Excel.GeometricShapeType.<x>)`, `addImage(base64)`, `addTextBox(text)`, and reading/setting `name`, `top`/`left`/`height`/`width`, `fill` (e.g. `setSolidColor`), `textFrame.textRange` (text/font/alignment), `geometricShapeType`, `lineFormat`. For images, the original image bytes are NOT exposed as a readable property (you set base64 in; there is no documented `image.getImageAsBase64()` getter), so an existing image shape cannot be recreated byte-for-byte from the API — existence + position only.
- Recreate fidelity: Geometric shapes and text boxes round-trip well (type, geometry, fill, text). Images are existence/position-only on round-trip (cannot re-extract source bytes). Connectors/grouped shapes/SmartArt/freeform: NOT FOUND IN DOCS — needs empirical spike.
- Change event: NO shape-changed event; only `onActivated`/`onDeactivated`. Movement/resize of an existing shape is not eventable. Suspicion CONFIRMED (no change event).
- Requirement set: Shapes API ExcelApi 1.9.
- Source: docs/excel/excel-add-ins-shapes.md; docs/excel/excel-add-ins-events.md.
- Implication for capability-map.md: Current "Images / shapes — snapshot — partial — limited shapes API" is ACCURATE but understated about the cause. Add to Notes: geometric shapes/text boxes recreate well; IMAGE shapes are existence/position-only because source bytes aren't readable back. That is the real fidelity boundary.

### Slicers — tier WRONG (should be Tier 2 config-snapshot, not Tier 3 existence-only)
- Read fidelity: There IS a real slicers API. `worksheet.slicers` and `workbook.slicers` collections; `slicers.add(source, field)` (source can be a PivotTable or a Table; field is a PivotField/column). Readable/settable: `name`, `caption`, position (`top`/`left`/`height`/`width`), `style`, and the slicer's selected/filtered items (`slicerItems`, `selectItems`). So slicers are config-reconstructable, not merely existence-trackable.
- Recreate fidelity: Good for slicers bound to a Table or worksheet-range PivotTable — recreate via `slicers.add` + restore caption/geometry/style/selected items. Slicers over Data-Model pivots inherit the pivot's Tier-3 limitation. Timeline slicers: NOT FOUND IN DOCS — needs empirical spike (likely no API; keep timeline slicers Tier 3).
- Change event: NO slicer-change event (no eventable selection change documented); only collection add/delete + onActivated. (Selection state must be snapshotted.)
- Requirement set: Slicers API ExcelApi 1.10.
- Source: docs/excel/excel-add-ins-pivottables.md (slicer create + style sections); docs/excel/excel-add-ins-events.md.
- Implication for capability-map.md: CORRECTION. The Tier 3 row "Slicers / timeline slicers — limited API" is wrong for regular slicers. Move *regular slicers* to Tier 2 (snapshot, good restore: create/caption/style/geometry/selected items). Keep *timeline slicers* in Tier 3 (no confirmed API).

### Sparklines — tier CONFIRMED (Tier 3, existence-only)
- Read fidelity: NONE. No `sparkline`/`sparklines` member found anywhere in the Excel JS API docs (no create, read, or event surface). Suspicion CONFIRMED (no API).
- Recreate fidelity: None via Office.js. Existence can only be inferred indirectly (e.g., a cell that visually hosts a sparkline is otherwise an ordinary cell to the API) — effectively not even reliably existence-detectable.
- Change event: n/a.
- Requirement set: none.
- Source: absence across docs/excel/* (no sparkline topic in office-js-docs-pr); NOT FOUND IN DOCS for any sparkline API.
- Implication for capability-map.md: Tier 3 "Sparklines — no API" is CORRECT. No change.

### Form controls / ActiveX / OLE objects — tier CONFIRMED (Tier 3, no API)
- Read fidelity: NONE. No form-control, ActiveX, or OLE-object API surface in the Excel JS docs. Suspicion CONFIRMED.
- Recreate fidelity: None.
- Change event: n/a.
- Requirement set: none.
- Source: absence across office-js-docs-pr; NOT FOUND IN DOCS.
- Implication for capability-map.md: Tier 3 "Form controls / ActiveX / OLE objects — no API" is CORRECT. No change.

### Tables (ListObjects) — tier CONFIRMED (data Tier 1, style/structure Tier 2)
- Read fidelity: Strong. `worksheet.tables`, `tables.add(range, hasHeaders)`, `table.name`, `columns` (add/read), header row range/values, `rows.add`, `style` (e.g. banded rows via style name), header/total/filter-button toggles. Data is eventable via `Table.onChanged` (Tier 1); structure/style is snapshot (Tier 2).
- Recreate fidelity: High — recreate via `tables.add` + name + column headers + rows + style. Edge cases (calculated columns with formulas, custom table styles beyond named built-ins): partially covered; custom style definitions NOT FOUND IN DOCS — minor spike.
- Change event: YES for data — `Table.onChanged` (and `onSelectionChanged`); structural/style changes are snapshot-only (no style-change event).
- Requirement set: Tables core ExcelApi 1.1; `Table.onChanged` 1.7+.
- Source: docs/excel/excel-add-ins-charts-data-labels.md (tables.add usage), docs/tutorials/excel-tutorial.md, docs/excel/excel-add-ins-events.md, capability-map existing Tier 1/2 split.
- Implication for capability-map.md: Existing split ("Tables — data" Tier 1 event; "Tables — style/structure" Tier 2 snapshot) is CORRECT. No change.

### Corrections needed to capability-map.md:
1. **Slicers (regular): move Tier 3 → Tier 2.** Real API exists (`worksheet.slicers`/`workbook.slicers`, `slicers.add`, caption/style/geometry/selected-items). Capture = snapshot; restore = good. Keep *timeline slicers* in Tier 3 (no confirmed API). Edit both the Tier 3 "Slicers / timeline slicers" row and add a Tier 2 "Slicers (regular)" row.
2. **Charts (Tier 2): add a Notes caveat** that the chart *source data range* is not a readable property — capture it at chart-creation time; cannot be recovered from an existing chart object.
3. **Images/shapes (Tier 2): refine Notes** — geometric shapes and text boxes recreate well; IMAGE shapes are existence/position-only because source image bytes are not readable back. State the cause, not just "limited shapes API."
4. **PivotTables — basic (Tier 2): refine Notes** — explicitly list showAs (ShowAsCalculation), PivotLayout (layoutType/grand totals/subtotals/preserveFormatting), and PivotFilters as reconstructable; and state Data-Model/OLAP pivots cannot even be created via `pivotTables.add` (range-only source), reinforcing their Tier 3 placement.

No corrections needed for: Charts tier, PivotTable advanced-features Tier 3, Sparklines, Form controls/ActiveX/OLE, Tables.


## 5. Persistence, Identity, Lifecycle & Manifest

Docs-grounded investigation via Context7 against `/officedev/office-js-docs-pr`. Each claim cites a doc or is flagged "NOT FOUND IN DOCS — needs empirical spike."

---

### 1. Workbook stamp — identity that travels with the .xlsx (ADR-0006) — CONFIRMED / NUANCED

- Detail: There are TWO in-file stores, and they have different shapes:
  - **`Office.context.document.settings`** (Common API) / **`workbook.settings`** (Excel API) — a property bag of serialized JSON name/value pairs. It is "specific to the add-in instance and the document it's saved in." Held in memory at runtime; `set`/`get`/`remove` mutate the in-memory copy; `saveAsync()` (Common) or `context.sync()` after `settings.add` (Excel) persists it INTO the document. It IS saved inside the document, so it travels with the file. BUT it is scoped to the add-in + that document, and is the right tool for small structured state (a GUID, a hash, a last-Step pointer).
  - **`workbook.customXmlParts`** (Excel API) / `Office.context.document.customXmlParts` (Common API) — embeds arbitrary custom XML directly into the Open XML (.xlsx) package. "This data persists independently of the add-in itself." The XML string MUST include an `xmlns` namespace attribute. `customXmlParts.add(xml)` returns a part with an `id`; the documented pattern is to store that `id` in `settings` so you can retrieve the part across sessions (the part itself is addressed by namespace via `getByNamespace`, or by the stored id).
- **Recommendation for the workbook stamp:** Use BOTH in the documented combination, but lead with `settings`. Store the workbook GUID + last-Step hash as small keys in `workbook.settings` (simplest, JSON, no namespace ceremony). If you need the stamp to be robust against other add-ins / external tooling and to live as a first-class, independently-addressable blob in the package, write it as a `customXmlPart` under your own namespace (e.g. `https://omototo.github.io/timeline/stamp/1.0`) and cache its `id` in settings. For a "workbook GUID + last-Step hash," plain `settings` is sufficient and lowest-friction; customXmlParts is the heavier, more isolation-resistant option. Both survive email/machine-move because both are serialized into the .xlsx itself.
- Size limits: **NOT FOUND IN DOCS** for the `settings` property bag specifically (no documented per-key or total byte cap was returned — needs empirical spike; historically ~2MB on the web but the docs returned did not state it). The 5MB / 5,000,000-cell limits found apply to the **API payload / `context.sync` transfer**, NOT to the persisted settings store. Treat "settings byte cap" as NOT FOUND — needs a runtime spike if the stamp could grow large. A GUID + one hash is tiny, so this is not a near-term risk.
- Requirement set: `settings` (Common API) is broadly available; `workbook.settings` and `workbook.customXmlParts` (Excel host objects) require **ExcelApi 1.4** (custom XML parts entered ExcelApi at 1.4). Common-API `CustomXmlParts` requirement set also exists and is checkable via `isSetSupported('CustomXmlParts')`.
- Source: `docs/develop/persisting-add-in-state-and-settings.md`; `docs/develop/support-for-task-pane-and-content-add-ins.md`; `docs/concepts/resource-limits-and-performance-optimization.md` (payload limits); `docs/excel/performance.md`.
- Implication: ADR-0006 identity-travels-with-file is satisfied. Use `workbook.settings` for the GUID + last-Step hash (call `context.sync()` to persist); optionally mirror into a namespaced `customXmlPart` for hardening. Do NOT rely on `localStorage`/IndexedDB for the travelling stamp — those are per-machine/per-origin and do not move with the file (see #5).

---

### 2. Add-in lifecycle — re-attach & rehydrate every launch — CONFIRMED

- Detail: Startup goes through **`Office.onReady(callback)`** (current) or the legacy **`Office.initialize = function(reason){...}`**. `Office.initialize` receives a `reason` ('inserted' vs 'documentOpened') so you can distinguish first-insert from reopening an existing stamped workbook. Event handlers (e.g. `sheet.onChanged.add(...)`, `addHandlerAsync`) must be **re-registered inside `onReady`/`initialize` on every load** — the docs' canonical pattern registers worksheet change handlers there so they are active "even when the task pane is closed." Dialogs run in a **separate execution context** and must call their OWN `Office.onReady`.
- JS memory across pane close/reopen: The docs strongly imply (and the navigation-state sample explicitly designs around) **state NOT surviving** a close/reopen: it saves to `settings` + `localStorage` on navigate and **rehydrates in `Office.onReady`** on reopen. So: assume JS in-memory state is LOST when the pane closes; rehydrate from the store on every launch. (A shared runtime keeps code alive longer, but the safe default is rehydrate.)
- Requirement set: `Office.onReady`/`initialize` are core (n/a / all versions). Persisting across close/reopen via the documented pattern is requirement-set-agnostic.
- Source: `docs/develop/initialize-add-in.md`; `docs/develop/run-code-on-document-open.md`; `docs/design/navigation-patterns.md`; `docs/develop/connect-to-javascript-frameworks.md`.
- Implication: On every launch — (a) run `Office.onReady`, (b) read the workbook stamp from `settings`/customXmlParts to verify identity & last-Step hash, (c) re-attach all event handlers, (d) rehydrate UI from the store. Never assume a variable set last session is still in memory.

---

### 3. Calculation — Frozen-Value / Preview (ADR-0008) — CONFIRMED

- Detail: `workbook.application.calculationMode` is **"Automatic"** by default (also `automaticExceptTables` and `manual`). `Application.calculate(calculationType)` forces a recalc — types are `'full'`, `'fullRebuild'`, `'recalculate'` (the REST/host doc lists 'full'|'fullRebuild'|'recalculate'). `suspendApiCalculationUntilNextSync()` temporarily suspends recalc for batched writes (perf). **Volatile functions** — `NOW`, `RAND`, `TODAY` are explicitly named as built-in volatile functions that "recalculate every time Excel performs a recalculation," independent of their arguments.
- Requirement set: `calculationMode` / `Application.calculate` are **ExcelApi 1.1**; `suspendApiCalculationUntilNextSync` is **ExcelApi 1.6**.
- Source: `docs/excel/excel-add-ins-workbooks.md`; `docs/excel/performance.md`; `docs/excel/custom-functions-volatile.md`.
- Implication: Frozen-Value/Preview is feasible. To freeze, you must either snapshot `range.values` (not formulas) or switch `calculationMode = 'manual'` to stop volatile recalcs, then `Application.calculate('full')` on demand for Preview. Be aware volatiles WILL re-fire on open/any recalc in Automatic mode — that's exactly the behavior ADR-0008 must defend against by capturing values rather than leaving live `=NOW()`/`=RAND()` formulas in a "frozen" cell.

---

### 4. Co-authoring detection (ADR-0006 puts co-authoring out of scope) — CONFIRMED (detectable per-event)

- Detail: Every relevant Excel event arg carries a **`event.source`** property = **`Local`** (current user) or **`Remote`** (a remote co-author). This is documented under "Events and coauthoring" and appears on `onChanged`, `onFormulaChanged`, `onProtectionChanged`, etc. So you CAN tell, per event, that a remote co-author triggered the change.
- NUANCE / NOT FOUND: There is **no documented "is this workbook currently shared / being co-authored?" boolean/property** returned in these queries — detection is **reactive per-event** (you only learn co-authoring is happening once a `Remote`-sourced event arrives), not a pre-flight query. A proactive "co-authoring active" flag is **NOT FOUND IN DOCS — needs empirical spike** (or treat any `source == Remote` event as the trigger to disable tracking).
- Requirement set: `event.source` on `onChanged` is **ExcelApi 1.7+** (events with source); other event-specific sources vary by their own set version.
- Source: `docs/excel/excel-add-ins-events.md` ("Events and coauthoring"); `docs/excel/excel-add-ins-worksheets.md`.
- Implication: ADR-0006 can be enforced reactively: when any tracked event has `source == Remote`, treat the workbook as co-authored and disable/skip tracking for that change. If you need to disable tracking BEFORE the first remote edit (proactively), there is no documented signal — spike required.

---

### 5. IndexedDB / `navigator.storage.persist()` / partitioning in the WebView — NUANCED / partially NOT FOUND

- Detail: The docs cover **`localStorage`** in add-ins and **storage partitioning**, but did NOT return anything specific about **IndexedDB** or **`navigator.storage.persist()`** or eviction quotas.
  - **Partitioning:** `Office.context.partitionKey` may be defined; when defined it is a hash of (top-level host domain, e.g. `excel.cloud.microsoft) + (add-in domain). You must prefix storage keys with it: `localStorage.setItem(partitionKey + key, value)`. The key is **undefined in WebView controls on Office on Windows** (no partitioning there) and defined on the web (Chromium ≥115 partitions storage to prevent cross-site tracking).
  - **Availability caveat:** "Browser-based storage techniques like cookies or localStorage might be blocked by certain browsers or user settings... test for their availability."
- IndexedDB specifically, `navigator.storage.persist()`, and eviction behavior: **NOT FOUND IN DOCS — needs empirical spike.** (Office-js docs do not document IndexedDB in the add-in WebView; this is a valid "not found" answer per the brief.)
- Requirement set: `Office.context.partitionKey` is a Common API property (n/a to ExcelApi versioning; relatively recent — feature-detect by `typeof`/truthiness rather than a requirement set).
- Source: `docs/develop/persisting-add-in-state-and-settings.md` (Browser storage / Storage partitioning sections).
- Implication: Browser storage (localStorage and, by extension, IndexedDB) is **per-machine, per-origin/partition** and does NOT travel with the file — so it is fine for a local cache/queue but MUST NOT hold the canonical identity/stamp (that goes in `settings`/customXmlParts, #1). Always (a) prefix keys with `Office.context.partitionKey` when defined, (b) feature-test storage availability, (c) treat IndexedDB persistence/eviction as unverified until a runtime spike confirms it inside the Windows WebView and web hosts.

---

### 6. Manifest requirement sets vs runtime feature-detection — CONFIRMED

- Detail: Two distinct gating mechanisms:
  - **Install-time gate (XML manifest):** `<Requirements><Sets DefaultMinVersion="1.1"><Set Name="ExcelApi" MinVersion="1.x"/></Sets></Requirements>` (also `<Methods>`). If the Office client does NOT support the listed sets/versions, the add-in **cannot be installed/activated**. Can also live inside `<VersionOverrides>` to gate add-in commands. (Unified/JSON manifest uses `requirements.capabilities: [{name, minVersion}]`.)
  - **Runtime feature-detection:** `Office.context.requirements.isSetSupported('ExcelApi', '1.7')` returns a boolean (2nd arg optional, defaults to '1.1'). Use this to branch on optional features without blocking installation.
  - **Important asymmetry:** Some sets (notably **DialogApi 1.2**) **cannot be specified in the manifest** and MUST be checked at runtime via `isSetSupported`. "Support for [some] manifest requirements is currently under development."
- Requirement set: n/a (this is the mechanism itself).
- Source: `docs/develop/xml-manifest-overview.md`; `docs/develop/specify-office-hosts-and-api-requirements.md`; `docs/develop/specify-api-requirements-runtime.md`; `docs/develop/understand-requirement-configuration.md`; `docs/develop/initialize-add-in.md`.
- Implication: Put the **floor you cannot run without** (your minimum ExcelApi — see below) in the manifest `<Requirements>` to block install on too-old Office. Feature-detect **optional / newer** capabilities (DialogApi 1.2 messaging, ExcelApi >1.4 niceties) at runtime with `isSetSupported`. Do NOT manifest-gate DialogApi 1.2 — runtime-check it.

---

### 7. Dialog API for license-key entry / Keygen activation (ADR-0010) — CONFIRMED

- Detail: `Office.context.ui.displayDialogAsync(url, options, callback)`. URL must be **HTTPS and same domain** as the host page. Options include `{height, width}` (percent). The returned `dialog` (in `asyncResult.value`) exposes `addEventHandler(Office.EventType.DialogMessageReceived, ...)`; the dialog page calls **`Office.context.ui.messageParent(...)`** to send the key/result back. The dialog runs in its **own execution context** and must call its own `Office.onReady`. Handle **error 12009** = user blocked the dialog (prompt them to click again and Allow). Parent→dialog messaging requires **DialogApi 1.2** (runtime-checked, not manifestable).
- Requirement set: **DialogApi 1.1** for basic open + `messageParent` (child→parent); **DialogApi 1.2** for parent→child `messageChild` (must be `isSetSupported`-checked at runtime).
- Source: `docs/develop/dialog-api-in-office-add-ins.md`; `docs/develop/dialog-handle-errors-events.md`; `docs/develop/dialog-best-practices.md`; `docs/develop/connect-to-javascript-frameworks.md`.
- Implication: License-key entry / Keygen activation flow works: open an HTTPS same-origin dialog, collect the key, `messageParent` it back, handle 12009 (blocked) with a re-prompt. Basic flow needs only DialogApi 1.1; only add DialogApi 1.2 if you must push data INTO the dialog after it opens — and feature-detect it.

---

### Minimum ExcelApi version this cluster needs: 1.7

Rationale: `customXmlParts` enters at **1.4**, `calculationMode`/`Application.calculate` at **1.1**, `suspendApiCalculationUntilNextSync` at **1.6**, and **`event.source` (Local/Remote co-authoring detection on `onChanged`) at 1.7** — the highest floor. Manifest-gate **ExcelApi 1.7**; runtime feature-detect DialogApi 1.2 and `Office.context.partitionKey` separately (neither is an ExcelApi version, and DialogApi 1.2 is not manifestable).
