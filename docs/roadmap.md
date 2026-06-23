# Roadmap — parallel work streams

The frozen interface ([`docs/engine-interface.md`](./engine-interface.md)) lets work proceed as independent streams that integrate through PRs into protected `main`. @omototo approves every merge.

## Streams

| Stream | Owner | Branch | File ownership | Depends on |
|---|---|---|---|---|
| **A — Engine algorithm** | Claude | `feat/engine-algorithm` | `packages/engine/**` | frozen interface |
| **B — Office.js adapters** (shell) | Claude | `feat/office-adapters` | `packages/addin/src/excel/**` | frozen ports + `docs/office-js-findings.md` |
| **C — Task-pane UI** | Codex | `feat/task-pane-ui` | `packages/addin/src/ui/**` | UI contract (`src/ui/contract.ts`) |
| **D — Spikes + on-host bench** | interactive | — | — | real Excel (sideload) |

A, B, C develop in parallel against the frozen interface (and fakes). D needs a live host and is done interactively.

## Integration discipline

1. **The interface in `docs/engine-interface.md` is frozen.** A stream that needs a change to it does not edit it unilaterally — it raises the change in its PR for a deliberate, reviewed update (a ripple across streams).
2. **File ownership is exclusive** (table above). The only shared seam inside `packages/addin` is a thin wiring file, written at integration time, not by B or C.
3. **All merges are PRs into `main`** — CI (`build`) green + @omototo approval. Squash merge; branch auto-deleted.
4. **No `Co-Authored-By` / AI-attribution trailers** in commits (repo rule).

## Integration order

A (engine) merges first; B (adapters) wires the engine to Excel; C (UI) consumes the engine's `timeline()` view. B and C are built against fakes, so they can land in any order and are connected at integration.

## Status

- A: Waves 1–3 merged on branch (value, structural, reconstruction); 4–5 (branching/lifecycle, queries/benchmark) in progress.
- B: scaffolding (this branch).
- C: brief + contract + fixture published (`feat/task-pane-ui`), handed to Codex.
- D: pending — needs a sideloaded build.
