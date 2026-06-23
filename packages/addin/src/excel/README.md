# Office.js adapters (Stream B)

The **shell** that wires the pure engine to Excel. Everything here is the only code allowed to touch Office.js. Each adapter implements a frozen seam and carries the constraints established in `docs/office-js-findings.md` and the ADRs.

## Adapters to build

### `OfficeChangeSource` → emits `Observation`s (ADR-0013)

- Registers `Worksheet.onChanged` / `onFormatChanged` and worksheet-collection events.
- **Debounces** the event burst of one user action into one Step's worth of Observations (window ~250–400 ms; confirm via the fan-out spike).
- **Reads back** the changed region to build the `after`-slab — **chunked** to respect the **5 MB payload** (web) and **5,000,000-cell** (all platforms) limits; validate `range.cellCount` before `sync()`; `untrack()` proxies in the loop. (`office-js-findings.md` §2.)
- **Echo filter:** drop events where `triggerSource === 'ThisLocalAddin'` (ExcelApi 1.14); below 1.14 fall back to the expected-write set. (ADR-0002.)
- Translates Office.js enums/types into engine-neutral `Observation` values (no Office types cross the seam).
- Maps structural `changeType` + `changeDirectionState` → `StructuralObservation`; `source: 'remote'` → signal the engine to suspend (co-authoring, ADR-0006).

### `RealSheetRenderTarget` / `PreviewSheetRenderTarget` → consume `ReconcilePlan` (ADR-0002/0008)

- Single **echo-cancelled choke point**: every write registers the expected-write set / runs under the re-entrancy guard so its own `onChanged` echoes are swallowed.
- `RealSheet`: live writes (`mode: 'formula'`).
- `PreviewSheet`: an engine-owned **`veryHidden`** worksheet (ExcelApi 1.1) rendered with **frozen values** (`mode: 'value'`); created/activated/deleted per the plan's ops.
- Apply structural ops via `Range.insert/delete` (native reference-shift — ADR-0003); batch before one `sync()`; `untrack()`.

### `IndexedDbStore implements HistoryStore` (ADR-0007)

- The 10 `HistoryStore` methods backed by IndexedDB; request `navigator.storage.persist()`.
- `InMemoryStore` (in the engine package) remains the test/bench double.

### `WorkbookStamp` (ADR-0006)

- `read()`/`write({ workbookGuid, tipHash })` via `Office.context.document.settings` (in-file, travels with the `.xlsx`); optionally mirror into a namespaced `customXmlPart`.

## Setup notes

- Add `@types/office-js` (or the `office-js` runtime types) as an addin dev dependency.
- Run `bun run typecheck && bun run lint && bun run test` green before the PR.
- Spikes that gate correctness here (delete-path refs, event fan-out, undo, IndexedDB eviction) are Stream D — coordinate before relying on unverified behavior.
