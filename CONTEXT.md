# Excel Parametric Timeline

A non-destructive, branchable history for an Excel workbook, surfaced as a zoomable timeline in a task pane. Tracks data, formulas, formatting, and embedded objects (charts, pivots) so any past state can be previewed or branched from without losing the present. History is scoped to the **whole workbook** (all worksheets), not a single sheet.

## Language

**Shadow State**:
The complete in-memory mirror of the live sheet (values, formulas, formats) that the engine maintains as its own source of truth. Office.js change events act only as a trigger and a bounding-box; the actual before/after data is derived by diffing freshly-read values against the Shadow State.
_Avoid_: cache, snapshot (a snapshot is a point-in-time copy; the Shadow State is continuously maintained)

**Step**:
The atomic unit of history — one user-initiated action and all of its synchronous fallout (e.g. an edit plus the recalc cascade it triggers), coalesced into a single timeline entry. Produced by debouncing the burst of Office.js change events a single action emits; distinct actions stay distinct because change events only fire on commit.
_Avoid_: edit, change, revision, version (a Step may bundle several raw change events)

**Observation**:
The raw host facts about one (debounced) change, emitted by the capture adapter before any history meaning is assigned: address, change type, shift direction, trigger source, and the read-back values/formulas/formats slab. The engine turns Observations into Deltas by diffing them against the Shadow State — the capture layer never computes a Delta itself.
_Avoid_: event, change, Step (an Observation is pre-history; the engine decides if/where it becomes a Step)

**Delta**:
The recorded content of a Step — the minimal information needed to move the sheet forward or backward across that Step. Exists in two distinct classes (Value Delta, Structural Delta) that replay differently.
_Avoid_: diff, patch, change-set

**Value Delta**:
A sparse list of `(address, before, after)` entries for value, formula, and format edits that occur within a fixed coordinate space. Replayed by writing values back.
_Avoid_: cell diff

**Structural Delta**:
A coordinate-remapping transform (insert/delete row or column, move range) that shifts the addresses of existing cells rather than changing their contents. Replayed by *applying the operation* and broadcasting it to every coordinate-keyed store (Shadow State, format map, chart/pivot anchors, the formula engine) — never as a set of value writes.
_Avoid_: layout change, shift

**Render Target**:
A live Excel surface that the engine projects (reconciles) a model state onto. The user's real sheet is the primary Render Target; the engine also owns throwaway Render Targets (Preview Sheets) for showing states other than the present. Excel holds no authority — it only displays what the engine reconciles onto it.
_Avoid_: view, output

**Present**:
The mode in which the user's real sheet is the active Render Target, the engine is recording, and new Steps append to the tip of the current branch. The only mode in which user edits become history.
_Avoid_: live, now, head

**Preview**:
The mode in which the engine reconciles a past Step (or another branch) onto a read-only Preview Sheet, leaving the user's real sheet frozen and untouched. The user looks but does not alter the present. Exited via "Return to Present".
_Avoid_: time-travel, view-history, rollback (rollback changes the present; Preview does not)

**Preview Sheet**:
A throwaway, engine-owned worksheet used as a Render Target during Preview, discarded on "Return to Present". Not part of the user's document.
_Avoid_: temp sheet, scratch sheet, ghost sheet

**Timeline**:
The task-pane visualization of history as a zoomable histogram, navigable on two axes: **temporal** (scrub/zoom across Steps; bar height reflects Delta size) and **structural** (zoom from the whole-workbook view down into a single worksheet's Steps). Also where branches appear as splits and the user switches between them.
_Avoid_: history bar, scrubber, log

**Worksheet Delta**:
A Structural Delta whose subject is a whole worksheet rather than a range — add sheet, delete sheet, rename sheet, reorder sheets. Replayed by performing the corresponding workbook-level operation and broadcasting it to every sheet-keyed store.
_Avoid_: tab change, sheet edit

**Frozen Value**:
The evaluated result of a cell captured at the moment of a Step and stored alongside its formula text. Preview renders Frozen Values — never live formulas — so the past is shown exactly as it was, immune to recalculation drift (volatiles like `=TODAY()`/`=RAND()`, cross-sheet references, or any dependency that differs now). Formulas are shown during Preview only as inert, inspectable metadata.
_Avoid_: cached value, snapshot value, static value

**Lossless Capture**:
The principle that each Step stores the full superset of a cell's state — formula text, evaluated value, value type, and number format — never a lossy subset. Guarantees any future capability (live recalc-in-preview, what-if branches, headless evaluation, semantic diffing) is buildable without re-architecting the capture layer.
_Avoid_: full snapshot, deep copy

**Fidelity Tier**:
The classification of an Excel feature by how faithfully the engine can capture and restore it, given Office.js limits. Tier 1 = full fidelity, event-driven; Tier 2 = config fidelity, snapshot-driven, best-effort; Tier 3 = existence-tracked only, restore not guaranteed. Maintained as a living capability map ([docs/capability-map.md](./docs/capability-map.md)).
_Avoid_: support level, coverage

**Fidelity Caveat**:
An honest marker attached to a Step when an object on it cannot be guaranteed to restore faithfully (a Tier 3 object, or a Tier 2 object the API can't fully round-trip). The system flags rather than silently degrades — it never pretends to a fidelity it cannot deliver.
_Avoid_: warning, error

**Reconciliation Step**:
A single Step recorded when the engine reattaches and finds the live workbook has drifted from its Shadow State (the workbook was edited without the engine running). It holds a full, itemized per-cell/per-sheet diff between live reality and the last known Tip — inspectable down to the cell. It captures *what* differs, deliberately not the *sequence* of untracked edits, which is unknowable.
_Avoid_: external change, sync, merge, conflict

**Branch**:
A line of history that diverges from a specific Step of a parent branch and accumulates its own Steps independently. Created deliberately ("Branch from here") while previewing; the parent branch's later Steps are preserved, never overwritten.
_Avoid_: fork, version, copy, scenario

**Tip**:
The latest Step of a branch — the state you edit when you are in Present on that branch.
_Avoid_: head, latest, end (HEAD is the active pointer, not the branch's end)

**HEAD**:
The persisted pointer to where the user currently is: which branch is active and whether they are at its Tip (Present) or viewing a past Step (Preview). Restored on reopening the workbook.
_Avoid_: cursor, position, current

**Switch** (branches):
Moving HEAD from one branch's Tip to another's and reconciling that Tip onto the real sheet. Non-destructive navigation — never deletes the branch left behind and never records a Step.
_Avoid_: checkout, jump, change branch

**Provisional Branch**:
A Branch that has been created but has no committed Steps yet. Auto-discarded when the user switches away, so abandoning a fork before editing leaves no trace. Becomes a persisted Branch once its first Step lands.
_Avoid_: draft branch, empty branch, temp branch

