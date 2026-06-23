# All engine writes route through one choke point with echo cancellation

## Context

The engine writes historical values back to the live sheet during preview, replay, branch-switch, and rollback. Every such write fires `onChanged`, which the handler would otherwise record as a new user Step — corrupting the timeline and risking an infinite feedback loop. Office.js cannot distinguish the add-in's own writes from the user's: `event.source` reports `Local` vs `Remote` (us vs. a co-author), and both our writes and the user's typing are `Local`.

## Decision

Every write to the live sheet routes through a single choke point. No component may write to Excel outside it.

Echo detection is **version-tiered** (the tiers were pinned by the Office.js investigation — see `docs/office-js-findings.md`):

- **ExcelApi ≥ 1.14 (primary): use `WorksheetChangedEventArgs.triggerSource`.** Office.js reports `triggerSource === "ThisLocalAddin"` for events raised by our own writes — a *direct* echo signal. The handler simply drops those events. No bookkeeping required. (This supersedes the original plan; `triggerSource` was not surfaced by the conceptual docs and was found only in the API reference.)
- **ExcelApi < 1.14 (fallback): re-entrancy guard + expected-write set.** The choke point raises a guard for the synchronous apply window and registers the exact `(address → value)` set it is about to write; the handler swallows events matching the guard or draining the set. The guard alone is insufficient because `onChanged` arrives asynchronously after the guard may have cleared — the expected-write set catches the late echoes.

Note: `source` (`Local`/`Remote`) is co-authoring-only and can NEVER distinguish our writes from the user's (both are `Local`) — `triggerSource` is a distinct, newer property and is the correct signal.

## Consequences

- Single-writer invariant: any future code path that mutates the sheet directly will reintroduce the feedback loop. This must be enforced in review.
- On hosts ≥ 1.14 the false-swallow edge disappears (echo detection is identity-based via `triggerSource`, not value-based). The rare false-swallow (a real user edit equal to an expected echo value) only exists on the < 1.14 fallback path.
- `triggerSource`/`source`/`details` cannot be detected at install time via the manifest — they are gated by ExcelApi requirement sets and must be **runtime feature-detected** with `isSetSupported('ExcelApi','1.14')` to choose the echo path.
