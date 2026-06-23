---
status: superseded by ADR-0012
---

# Licensing: Keygen + Stripe, deterrent-level enforcement, lock creation not access

> **Superseded by [ADR-0012](./0012-open-source-apache-2-open-core.md).** The project went open source (Apache-2.0, open-core). The engine now contains no licensing gate at all; this decision survives only as historical context for a possible future *closed paid wrapper*, which would record its own decision.

## Context

The PRD specified Stripe + Keygen with a locally-checked JWT and, on expiry, "lock IndexedDB write-access." Two problems: the engine that checks the license may be TypeScript, not Rust (ADR-0004); and for an offline-first app all enforcement runs in the user's WebView, so it is inherently bypassable — a signed license stops *forging* but not *patching out the check*. "Lock IndexedDB write-access" also risked holding the user's own captured history hostage.

## Decision

- **Keygen** for cryptographic offline license validation (signed license keys/files; public key in the client verifies the signature; machine fingerprint binds the seat). **Stripe** for payment.
- Enforcement is a **deterrent, not DRM**: verify signatures, do an anti-rollback clock check (persist the maximum timestamp ever seen; a backward clock forces online re-validation), and stop there. No deep anti-tamper — it cannot be won offline and would punish honest users with friction.
- **Degraded mode on expiry locks *creation*, never *access*:** stop recording new Steps and disable branching, but keep read-only history, Preview, and Export of everything already captured. Gate new value; never hold the user's own data hostage.
- **Trial** = a time-boxed signed license issued on first run, sharing the paid-license code path.

## Consequences

- A determined pirate runs free indefinitely. Accepted as correct for an offline tool.
- Licensing is auth-provider-agnostic and offline — explicitly NOT Azure AD SSO, which is why the project does not use the generator's SSO template (ADR-0011).
