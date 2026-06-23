# The engine is authoritative; drift collapses into one inspectable Reconciliation Step

## Context

The Shadow State is authoritative only while the engine witnesses every mutation. The workbook can change behind the engine's back — edited with the add-in not running, native Undo, co-authoring. Replaying history onto a drifted workbook silently produces garbage. The PRD did not address this.

The user's directive: the plugin is authoritative; edits made without it are not something we try to reconstruct; and the opaque "net blob" reconciliation (a black box) should be minimized.

## Decision

1. **Stamp the workbook** with a GUID + the hash of the last-committed Step, stored inside the file (custom XML part / document setting) so identity travels with the `.xlsx`. The history store is keyed by that GUID.
2. **Verify on every attach** (launch / pane open): hash the live workbook, compare to the persisted Tip hash. Match → clean resume.
3. **On mismatch, record one Reconciliation Step** holding a full, itemized per-cell/per-sheet diff (computed via a canonical, deterministic workbook serialization — the one idea borrowed from git-xl). It is inspectable down to the cell. We capture *what* changed, not the *sequence* of untracked edits.
4. **Native Undo (Ctrl+Z) is a new forward Step**, not a timeline pop — no attempt to unify with Excel's undo stack.
5. **Co-authoring is out of scope for v1.** If the file is shared or `source: Remote` events appear, tracking is disabled with a clear message. A branching timeline plus multi-author merge is two products.

## Consequences

- The timeline's granular fidelity is only as good as the add-in's uptime; untracked sessions appear as a single inspectable boundary.
- Pre-drift history remains previewable (non-destructively, on a Preview Sheet); it is never auto-applied over the user's current work.
- A canonical workbook serialization is now a shared dependency of keyframes and reconciliation.
