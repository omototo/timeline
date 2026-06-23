# Timeline Engine — Interface Spec (living)

Concrete interface shapes resolved during interface grilling. Decisions of record live in the ADRs (esp. [ADR-0013](./adr/0013-engine-seams-observation-boundary.md)); this doc holds the *types*. Types are engine-neutral (no Office.js) — see ADR-0001.

## Execution model (Q1 — decided)

Functional core, imperative shell. The engine is a synchronous, stateful, in-memory instance holding the Shadow State + HEAD. Every method returns an **effect description**; the engine never performs or awaits I/O. The shell executes effects (Excel writes via RenderTarget, persistence via HistoryStore).

## Observation — the input boundary (Q2 — decided)

Engine-neutral discriminated union. Carries the read-back **after**-slab for value changes only (the engine holds "before" in Shadow State). Multi-area aware. Sheet-scoped. Open to extension (an `ObjectSnapshotObservation` for Tier-2 objects arrives via a separate Step-boundary path later).

```ts
type SheetId = string;
type Rect = { startRow: number; startCol: number; rowCount: number; colCount: number };
type Area = Rect[]; // RangeAreas-aware: one logical change may span disjoint rectangles

type CellValue = /* engine-neutral scalar | rich-value JSON */ unknown;
type ValueType = 'empty' | 'string' | 'number' | 'boolean' | 'error' | 'richValue';

interface CellSlab {
  values: CellValue[][];
  formulas: (string | null)[][];
  numberFormats: string[][];
  valueTypes: ValueType[][];
}

interface ObservationMeta {
  triggerSource: 'thisLocalAddin' | 'unknown'; // ExcelApi 1.14; 'unknown' below
  source: 'local' | 'remote';                  // co-authoring signal
}

interface ValueObservation extends ObservationMeta {
  kind: 'value';
  sheetId: SheetId;
  area: Area;
  after: CellSlab;
}

interface StructuralObservation extends ObservationMeta {
  kind: 'structural';
  sheetId: SheetId;
  changeType: 'rowInserted' | 'rowDeleted' | 'columnInserted' | 'columnDeleted' | 'cellInserted' | 'cellDeleted';
  address: Rect;
  shiftDirection?: 'down' | 'right' | 'up' | 'left'; // from changeDirectionState (1.14); inferred below
  // no slab: a structural op is a coordinate transform, not a value change
}

interface WorksheetObservation extends ObservationMeta {
  kind: 'worksheet';
  op: 'add' | 'delete' | 'rename' | 'reorder';
  sheetId: SheetId;
  newName?: string;     // rename
  newPosition?: number; // reorder
}

type Observation = ValueObservation | StructuralObservation | WorksheetObservation;
```

## Effects — the output boundary (Q3 — decided)

Every method returns an `EffectEnvelope`. `ReconcilePlan` is a **host-neutral, minimal-diff** op list computed engine-side against the engine's tracked "currently-projected" state (the RenderTarget stays dumb and the diff logic is headless-testable). Each cell op carries a **write mode**: `value` (Frozen Value, for Preview — ADR-0008) or `formula` (live, for Present). Echo is handled entirely in the shell; the engine's Observation input is already echo-filtered.

```ts
type WriteMode = 'value' | 'formula';

type ReconcileOp =
  | { op: 'setCells'; sheetId: SheetId; area: Area; slab: CellSlab; mode: WriteMode }
  | { op: 'applyStructural'; sheetId: SheetId; changeType: StructuralObservation['changeType']; address: Rect; shiftDirection?: 'down'|'right'|'up'|'left' }
  | { op: 'createPreviewSheet'; previewSheetId: SheetId }
  | { op: 'activateSheet'; sheetId: SheetId }
  | { op: 'deletePreviewSheet'; previewSheetId: SheetId };

interface ReconcilePlan { target: 'realSheet' | 'previewSheet'; ops: ReconcileOp[]; }

type PersistOp =
  | { op: 'appendDelta'; branchId: string; delta: Delta }
  | { op: 'writeKeyframe'; branchId: string; stepIndex: number; state: /* serialized */ unknown }
  | { op: 'setHead'; head: Head }
  | { op: 'saveBranch'; meta: BranchMeta }
  | { op: 'deleteBranch'; branchId: string };

interface EffectEnvelope { reconcile?: ReconcilePlan; persist?: PersistOp[]; }
```

## Verb surface (Q4 — decided)

All mutators return `EffectEnvelope`; queries are pure. Lifecycle folds resume + drift + stamping into one safe entry point.

```ts
interface TimelineEngine {
  // Lifecycle
  attach(observed: WorkbookSnapshot, persisted: PersistedHead | null): EffectEnvelope; // hash+compare; clean resume OR drift -> Reconciliation Step
  detachToCoauthoring(): EffectEnvelope; // on source:'remote' -> suspend tracking (ADR-0006)

  // Recording — Present mode only; if an Observation arrives in Preview the engine returns a no-op + diagnostic (shell must lock the real sheet during Preview)
  ingest(obs: Observation): EffectEnvelope;

  // Navigation
  goto(ref: StepRef): EffectEnvelope;        // enter Preview (frozen values, fresh Preview Sheet)
  returnToPresent(): EffectEnvelope;         // discard Preview Sheet, reactivate real sheet
  branch(from: StepRef): EffectEnvelope;     // "Branch from here" -> provisional editable Present
  switch(branch: BranchId): EffectEnvelope;  // checkout another branch tip (non-destructive, NOT a Step)

  // Queries (pure)
  head(): Head;
  timeline(opts?: TimelineQuery): TimelineView; // histogram model: steps, bar magnitudes, branch splits
  inspectStep(ref: StepRef): StepDetail;        // formula text metadata for Preview
}
```

## Persistence seams (Q5 — decided)

Two **shell-side** ports — the engine never calls them; it emits `PersistOp`s and consumes loaded data passed into `attach`/`goto`/`switch`.

- **`HistoryStore`** (async; IndexedDB per ADR-0007, `InMemoryStore` first): `appendDelta(branchId, delta)`, `writeKeyframe(branchId, stepIndex, state)`, `loadKeyframeAtOrBefore(branchId, stepIndex)`, `loadDeltas(branchId, from, to)`, `getHead()`, `setHead(head)`, `saveBranch(meta)`, `listBranches()`, `getBranch(id)`, `deleteBranch(id)`.
- **`WorkbookStamp`** (in-file `workbook.settings` per ADR-0006): `read()`/`write({ workbookGuid, tipHash })` — tiny, travels with the `.xlsx`.

**Memory policy:** Shadow State + HEAD + current branch's delta log/magnitudes are **resident** (deltas are small; the histogram needs them). **Keyframes are lazy** — `goto` replays from the nearest keyframe; a keyframe miss emits a `{ op:'loadKeyframe', branchId, stepIndex }` effect and the shell re-invokes (two-phase only on miss). `switch` lazy-loads the target branch's delta log.

## Core types + keyframe cadence (Q6 — decided)

**Navigation is forward-replay-only**: reach Step N by loading the nearest keyframe ≤ N and replaying *forward*. Deltas are never inverted (no backward-stepping logic). `ValueDelta` still stores `before` — needed for the inspect/diff UI and lossless capture (ADR-0008) — but navigation never uses it.

```ts
type BranchId = string;
type StepRef = { branchId: BranchId; stepIndex: number }; // ordinal; stable (branches are append-only)
type Head = { branchId: BranchId; mode: 'present' | 'preview'; previewStepIndex?: number };
type BranchMeta = { id: BranchId; parentBranchId?: BranchId; forkedAt?: StepRef; order: number; name?: string; provisional: boolean };

type CellState = { value: CellValue; formula: string | null; valueType: ValueType; numberFormat: string }; // lossless
type ValueDelta          = { kind: 'value';          sheetId: SheetId; cells: { addr: Rect; before: CellState; after: CellState }[] };
type StructuralDelta     = { kind: 'structural';     sheetId: SheetId; changeType: StructuralObservation['changeType']; address: Rect; shiftDirection?: 'down'|'right'|'up'|'left' };
type WorksheetDelta      = { kind: 'worksheet';      op: WorksheetObservation['op']; sheetId: SheetId; newName?: string; newPosition?: number };
type ReconciliationDelta = { kind: 'reconciliation'; perSheet: SheetDiff[] }; // ADR-0006, inspectable
type Delta = ValueDelta | StructuralDelta | WorksheetDelta | ReconciliationDelta;
```

**Keyframe cadence — adaptive:** write a keyframe when **either** N steps (default 100) **or** cumulative delta bytes since the last keyframe exceed a threshold — whichever first. Configurable. Bursts of large pastes keyframe more often (cheap replay); long tails of micro-edits keyframe rarely (cheap storage).

---

## Status: interface resolved (Q1–Q6). Next: materialize as compiling TypeScript in `packages/engine` (types + interfaces + `InMemoryStore` + tests), then build the engine algorithm behind it.

---

## Implementation notes (behind the frozen interface)

The frozen `TimelineEngine` surface (Q4) is implemented behind, unchanged. The following are **additive, non-breaking** extensions made while building the engine algorithm.

### Additive query methods on `TimelineEngineImpl`

These do not appear on the `TimelineEngine` interface; they are concrete-class queries the shell/tests use to inspect engine state without exposing internals. All are pure.

- **Wave 1 (value path):** `readShadow(sheetId, row, col)`, `shadowCellCount(sheetId)`, `tipStepIndex(branchId?)`, `steps(branchId?)`, `lastDiagnostic()`.
- **Wave 2 (structural + worksheet paths):** `sheetMeta(sheetId)` → `SheetMeta | undefined`, `shadowSheets()` → `SheetMeta[]` (tab order). Both delegate to the Shadow State's sheet map.
- **Wave 3 (keyframes + reconstruction + navigation):** `keyframeIndices(branchId?)` → `number[]` (stepIndexes at which keyframes were written, ascending) and `readReconstructed(ref, sheetId, row, col)` → `CellState` (forward-replay reconstruct at `ref`, then read one cell). Both are pure inspectors over the resident keyframes + delta log.

### Engine construction options (Wave 3) — adaptive keyframe cadence

`new TimelineEngineImpl(options?)`. The constructor takes an optional `TimelineEngineOptions`; the adaptive keyframe cadence (Q6) is configurable. A keyframe is written after appending a Step when **either** trigger fires:

```ts
interface TimelineEngineOptions {
  keyframeStepInterval?: number; // steps since last keyframe (default 100)
  keyframeByteThreshold?: number; // cumulative delta bytes since last keyframe (default 64 KiB)
}
```

Delta byte size is estimated by JSON-encoding length (a deterministic proxy for the store's serialized size). The keyframe payload is a serialized **Shadow State snapshot** (`ShadowSnapshot` — see below) for the branch + stepIndex, emitted as a `writeKeyframe` PersistOp *and* kept resident so reconstruction replays forward from it.

### Pinned: `ShadowSnapshot` (keyframe payload)

The spec's `PersistOp.writeKeyframe` carries a `state: /* serialized */ unknown` but did not pin the serialized shape. Pinned (minimal, structurally-cloneable — no `Map`s):

```ts
interface ShadowSnapshot {
  sheets: { sheetId: SheetId; cells: [cellKey: string, state: CellState][] }[];
  sheetMeta: SheetMeta[];
}
```

### Reconstruction + navigation (Wave 3)

Forward-replay-only (Q6): reconstruct at a `StepRef` by seeding from the nearest resident keyframe ≤ `stepIndex` (single branch for now) and applying the deltas in the window `(keyframeStepIndex, stepIndex]` forward — deltas are never inverted. `goto(ref)` flips HEAD → preview, reconstructs the target, and returns a `ReconcilePlan` targeting `previewSheet` that is the **minimal** `value`-mode (Frozen Values, ADR-0008) diff between the engine-tracked currently-projected state and the target.

**Multi-sheet Preview surfaces (ADR-0005).** History is workbook-scoped, so Preview is too: each **logical** source sheet projects onto its **own** Preview surface whose id is `previewSheetIdFor(sheetId)` = `` `__preview__::${sheetId}` `` (`PREVIEW_SHEET_PREFIX = '__preview__::'`, both exported from `project.ts`). The projection diff is computed **per logical sheet** and each `setCells` op carries that sheet's per-sheet preview id, so colliding coordinates across sheets (e.g. `Sheet1!A1` and `Sheet2!A1`) land on distinct surfaces and never overwrite each other. The first time a given logical sheet is projected in a Preview session, the plan is prefixed with a `createPreviewSheet` for its surface; the **very first** surface created is also `activateSheet`d. `returnToPresent()` returns a `realSheet` plan that deletes **every** per-sheet Preview surface created during the session, then emits a single `activateSheet` carrying the **real** `SheetId` that was active when Preview began (captured at the first `goto`) — never a `BranchId` (branch ids and sheet ids are distinct namespaces). When no active real sheet is knowable, the `activateSheet` op is omitted and the shell restores the previously-active sheet itself. `returnToPresent()` flips HEAD back to present and is a no-op when not in Preview.

### Pinned placeholder: `SheetMeta`

The spec's `WorksheetDelta` (add/delete/rename/reorder — ADR-0005) implies the engine tracks per-sheet metadata, but the spec did not pin a shape for it. Pinned (minimal-but-sensible): the **stable `sheetId`** (a sheet keeps its id across a rename), its display **`name`**, and its **`order`** (0-based left-to-right tab position, kept dense by add/delete/reorder).

```ts
type SheetMeta = { sheetId: SheetId; name: string; order: number };
```

### Structural path semantics (Wave 2)

A `StructuralObservation` becomes a `StructuralDelta` applied as a **coordinate remap** of the Shadow State (insert opens blank space and shifts cells down/right; delete removes the spanned cells and shifts the rest up/left). Per ADR-0001 it emits **no value diff** (a structural op is a coordinate transform, not a value change), and per ADR-0003 it **never rewrites formula text** — cell formulas are opaque strings the engine only relocates, never edits. Forward apply is deterministic (needed for replay). The `unsupportedKind` `IngestDiagnostic` code is now unreachable for `structural`/`worksheet`/`value`; it is retained for any future un-handled kind.


