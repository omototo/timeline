# Open source under Apache-2.0; open-core with no licensing gate in the engine

Supersedes [ADR-0010](./0010-licensing-deterrent-keygen-stripe.md).

## Context

The project is going open source. A code review flagged that a `LicenseGate` consulted by the host-agnostic Timeline Engine (ADR-0010) is business-policy contamination of domain logic — true even before the open-source decision, and incompatible with an OSS core. We chose **Apache-2.0** (permissive + explicit patent grant, business-friendly, open-core compatible) and an **open-core** model.

## Decision

- **License: Apache-2.0** for the whole repository. Contributions under DCO sign-off (no CLA).
- **The engine contains no licensing concept at all.** No `LicenseGate`, no expiry check, no degraded-mode — the engine and the core add-in are fully functional OSS. Candidate 6 (LicenseGate seam) is removed from the architecture.
- **Open-core:** any paid edition (e.g. cloud sync, team branches, hosted backup) lives in a **separate distribution-layer wrapper**, not in this core repo. If such an edition uses Keygen/Stripe, that integration lives in the wrapper and gates only the paid features it adds — never the open-source core.
- **Distribution is dual** (revises ADR-0011): AppSource remains the official binary for non-technical users; **self-host / sideload is first-class and documented**.

## Consequences

- The engine is freely embeddable and testable with no policy coupling — strengthens the host-agnostic goal (ADR-0001) and the depth review's intent.
- ADR-0010's deterrent-enforcement and "lock creation not access" degraded mode no longer apply to the core; they would only ever be relevant inside a future closed paid wrapper, governed by its own decision record.
- OSS hygiene (LICENSE, NOTICE, CONTRIBUTING with DCO, SECURITY, CODE_OF_CONDUCT) is part of the repo.
