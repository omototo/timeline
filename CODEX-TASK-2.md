# Codex Task 2 — Complete the Timeline Task Pane (`feat/task-pane-ui-v2`)

Round 2 of the UI. You shipped the MVP in PR #1 (`src/ui/TimelinePane.tsx` + `contract.ts` + `sample-timeline.ts`). Now make it **complete, polished, and demoable** — a task pane that feels like a real product. Branch off `main` (already done: this branch). Same rules as before (read `CODEX-BRIEF.md` if present, `CONTEXT.md`, ADR-0005/0006/0008, `docs/engine-interface.md`).

## Boundaries (expanded this round)

- **Own `packages/addin/src/ui/**` AND the add-in entry** (`src/App.tsx`, `src/main.tsx`) — you're mounting the pane as the actual add-in UI. Do NOT touch `packages/addin/src/excel/**`(adapters),`packages/engine/**`, or `docs/engine-interface.md`.
- Still **no Office.js / engine imports** in the UI. The pane talks only to a `TimelineDataSource` (provided — see below). At integration we swap the fake source for an engine-backed one; your code shouldn't change.
- You MAY extend the **UI-side** contract (`src/ui/contract.ts`) — e.g. add `TimelineCommand` variants for branch management. That's UI-owned; integration maps new commands to engine ops.

## The data-source seam (provided: `src/ui/data-source.ts`)

```ts
interface TimelineDataSource {
  getView(): TimelineView;
  subscribe(cb): () => void;
  dispatch(cmd): void;
}
```

`FakeTimelineDataSource` is a stateful fake that already reduces `goto` / `returnToPresent` / `switch` / `branch` so the pane feels live. Mount the pane against it (a `useSyncExternalStore` hook over `subscribe`/`getView` is the clean React pattern). Extend the fake's reducer as you add commands.

## Scope (expand freely within these themes)

1. **Round out the MVP**
   - **Temporal zoom**: a density/windowing control to zoom and pan across many Steps (a 1,000-step history must be navigable), in addition to the existing worksheet drill-down. The two axes (time + worksheet) are the Fusion-360 feel.
   - **Step inspector** panel: kind, magnitude, sheet, label; leave a slot for formula metadata (`inspectStep`, wired later).
   - **"Branch from here"** affordance: visible on a previewed Step, emits `branch`.
2. **Branch-management UX** (extend the contract as needed)
   - Name / rename a branch; delete a branch (with the provisional vs persisted distinction from ADR-0006); polish the split-timeline toggle; a branch-compare view is a nice-to-have.
3. **Visual fidelity**
   - Playhead/current-position marker, fork connectors between a branch and its parent at `forkedAt`, clear value/structural/worksheet/reconciliation styling, empty/loading states.
   - Office-native look: **Fluent UI React (`@fluentui/react-components`)** is the Office-standard component lib and is a reasonable dependency here — use it, or hand-roll; justify the choice in the PR. Keep the histogram itself SVG/divs for control.
4. **Mount + dev harness**
   - Render the pane in `App.tsx` via the `TimelineDataSource` (default `FakeTimelineDataSource`).
   - A dev way to view it with rich fixtures (multiple branches, deep history) — extend `sample-timeline.ts` with a larger fixture and/or a harness component. It should run in `bun run --filter @timeline/addin dev` (Vite).
5. **Cleanup**: delete `CODEX-BRIEF.md` and this `CODEX-TASK-2.md` at the end (handoff docs, not project docs).

## Conventions / workflow

- React 19 + strict TS (`verbatimModuleSyntax`, `exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`). Vitest + jsdom + @testing-library/react.
- **Tests** for the new behavior: temporal-zoom windowing, inspector contents, branch-from-here emits `branch`, branch rename/delete reduce correctly in the fake, the pane re-renders on data-source change. Keep coverage ≥80%.
- Conventional Commits, **NO `Co-Authored-By` / AI-attribution trailers**.
- Before the PR: `bun install` then `bun run typecheck && bun run lint && bun run format:check && bun run test` green.
- Open a PR into `main` (protected: CI `build` + @omototo approval). Use the PR template; in "Notes for reviewer" call out any contract extensions you made and anything you'd want the integration layer to honor.

## Definition of done

A demoable task pane: zoomable on both axes, scrub + preview + return, branch split with create/switch/rename/delete, a Step inspector, Office-faithful styling, mounted in the add-in entry against the fake data source, with green tests — and the handoff briefs removed.
