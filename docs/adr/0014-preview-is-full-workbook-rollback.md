# Preview is a full-workbook rollback, not a single frozen sheet

## Context

ADR-0008 established that Preview renders **Frozen Values**, and the first implementation projected each touched logical sheet onto its own hidden Preview Sheet, activating one of them. That works for a single-sheet workbook but breaks down once a workbook has several sheets:

- Only one preview surface was activated; the user's *other* real sheets stayed visible showing their **present** state. The workbook was a confusing mix of live and historical sheets — "what am I looking at depends on which tab I'm on."
- The preview surface showed up as a raw, internally-named tab the user was standing on.
- Worksheet add/delete were never reflected: previewing *before* a sheet existed still showed it.

## Decision

Preview is a **full-workbook rollback**. Scrubbing to step N shows the entire workbook as of N:

- The engine reconstructs the **full set of sheets that existed at N** (`sheetMeta ∪ populated`) and creates a Preview surface for every one — deleting surfaces for sheets that did not yet exist (a later-added sheet "disappears" as you scrub back). The cell diff is restricted to the surviving surfaces. The user is **anchored** to their own sheet's preview across scrubs.
- `goto`/`returnToPresent` flag the **transition** on the `ReconcilePlan` (`enterPreview`/`exitPreview`). The shell hides **all real sheets** on entry (remembering their visibility) and restores them exactly on exit, so during Preview only the historical surfaces are visible — no live/historical mix.
- Preview surfaces are **read-only** (worksheet protection): a frozen snapshot cannot be edited.
- The task pane shows a prominent **PREVIEW banner** (step number, an Exit-preview action) so the mode is unmistakable.

This extends ADR-0008 (still Frozen Values, still read-only); it changes only *how many* surfaces are shown and that the real sheets are hidden around them.

## Consequences

- Resolves the multi-sheet confusion: the workbook is wholly the past or wholly the present, never a mix.
- Heavier than single-sheet preview — a surface per sheet, and a hide/restore pass per session — accepted for fidelity; the cell diff stays minimal (only changed cells between consecutive previewed steps) so scrubbing remains responsive.
- Hiding/restoring real-sheet **visibility** is a shell (Office) concern — the engine models sheet existence and content, not visibility — so the shell owns the enter/exit, driven by the engine's transition flags.
- Deferred: friendly per-sheet preview tab **names** (e.g. "Sheet1 (history)") — the surfaces are recognizable and the real sheets are hidden, so naming is cosmetic and follows separately.
