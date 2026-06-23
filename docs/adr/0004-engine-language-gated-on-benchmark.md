---
status: proposed
---

# Engine language (TypeScript vs Rust/WASM) is gated on a benchmark

## Context

The PRD asserts JavaScript is too slow to diff large arrays and mandates Rust/WASM. That premise is unverified and likely misidentifies the bottleneck — the Office.js `getValues()` round-trip dominates, the diff is a linear scan, and the JS↔WASM boundary has its own marshalling tax. Committing to a Rust toolchain up front is a permanent cost paid against an unmeasured benefit.

## Decision

No engine-language decision until a benchmark is run. **TypeScript-in-a-Web-Worker is the default path** (the Worker already moves work off the UI thread — the real win attributed to Rust). Rust/WASM is adopted only for a *specific stage* that the benchmark proves exceeds budget with the I/O floor ruled out — introduced behind the engine interface as a drop-in, never as a wholesale rewrite.

Benchmark stages (all in TS-in-Worker first): (1) Office.js `getValues` for 50K/100K cells — the I/O floor; (2) JS diff vs Shadow State; (3) keyframe compress via native `CompressionStream`; (4) replay 100/1K/10K deltas; (5) end-to-end paste-50K → Step committed.

## Provisional budgets (to be confirmed by feel, not just numbers)

- Capture latency (paste → Step committed): **≤ 200 ms**
- Reconstruction / scrub (land on any Step): **≤ 500 ms**
- Per-stage Rust trigger: any single JS stage (2/3/4) exceeding **~50 ms** *while the I/O floor is materially smaller*

These thresholds are set provisionally; the prototype must be played with before they are accepted as final.

## The benchmark must run on-host, not only headless

The headless path (`ReplayChangeSource → Engine → fake RenderTarget → InMemoryStore`) measures **engine compute** and is necessary but **not sufficient**. The actual capture budget is dominated by Office.js I/O, which only a **real-Excel harness** can measure: stage 1 (`getValues` I/O floor) and stage 5 (end-to-end paste → Step committed) MUST run against a live host. Headless answers "is the engine fast enough"; on-host answers "do we hit the 200 ms capture budget." Both are required before the engine-language decision is made.
