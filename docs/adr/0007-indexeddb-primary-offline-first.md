# History lives in IndexedDB, offline-first; portability is a deliberate later feature

## Context

History (Deltas, keyframes, branches) needs a home. Options weighed: IndexedDB only; embed in the workbook (custom XML / hidden sheet); sidecar file; cloud sync. The product is Offline First and wants the fastest local store. Embedding in the file causes severe bloat and leaks abandoned branches to anyone the file is sent to; cloud breaks offline-first.

## Decision

IndexedDB is the primary working store for v1, chosen for speed and offline operation. `navigator.storage.persist()` is requested at startup to reduce the chance the host WebView evicts the store. Export/import of a compressed history bundle (sidecar) and optional in-file embedding are deferred features, not v1 gates. Cloud sync is out.

## Consequences (accepted risks)

- History is keyed to the machine/WebView origin, **not** the file: it does not travel when the `.xlsx` is emailed or opened on another machine until export/import ships.
- Because the file holds only the current branch Tip, losing IndexedDB loses every other branch. Persistent-storage mitigates but does not eliminate eviction.
- The UI must state plainly that history is local to this machine. Durability/portability becomes an informed user action (export) rather than a silent guarantee.
