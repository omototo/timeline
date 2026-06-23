# Distribution via AppSource with the XML manifest; scaffold with yo office (no SSO)

## Context

The add-in needs a distribution channel and a project scaffold. The Yeoman generator (`yo office`) / Microsoft 365 Agents Toolkit bootstraps the manifest, HTTPS dev certs, sideload/debug config, office-js typings, and a React+TS task pane — but only the shell; the engine architecture is all custom on top. The generator also offers an SSO (Azure AD) template, which conflicts with our offline Keygen licensing (ADR-0010). The manifest has two forms — mainstream XML vs emerging unified JSON — and the manifest form is AppSource-coupled and costly to migrate.

## Decision

- **Distribution channels (dual, per ADR-0012): AppSource + self-host/sideload.** AppSource is the official binary for non-technical users; self-host/sideload is first-class and documented in CONTRIBUTING. The XML manifest serves both.
- **Manifest: XML manifest** — fully AppSource-supported today; do not ride the unified-JSON-manifest curve without a specific reason.
- **Scaffold: `yo office`, Excel + React + TypeScript, plain task-pane template — NOT the SSO template** (our auth is offline Keygen, not Azure AD). Treat the scaffold as throwaway plumbing.
- Swap webpack → **Vite** early for faster cold start and dev loop (relevant if a `.wasm` payload ships).

## Consequences

- AppSource review/publishing constraints apply (review latency, content policies). 
- Choosing XML manifest defers unified-manifest features (Teams/M365 integration); revisit only if those become product requirements.
