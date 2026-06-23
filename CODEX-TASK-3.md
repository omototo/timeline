# Codex Task 3 — Production-harden the pane + Branch Compare (`feat/task-pane-ui-v3`)

Round 3. The pane is feature-complete against the fake (PRs #1, #2). Now make it **production-grade for the Office task pane** and add the **branch-compare** experience. Branch off `main` (done: this branch). Same rules as before.

## Boundaries

- Own `packages/addin/src/ui/**` + `src/App.tsx` (theme prop only). Do NOT touch `src/excel/**`, `packages/engine/**`, or `docs/engine-interface.md`.
- No Office.js / engine imports. Theme arrives via a **prop** (integration feeds the real Office theme later).
- You MAY extend the UI-side contract, but **prefer computing compare from the existing `TimelineView`** (no engine call). A cell-level state diff needs engine data → out of scope; leave a clearly-marked slot for it.

## Scope

1. **Branch compare view** — pick two branches (or a branch vs its parent); show the divergence point (`forkedAt`), a side-by-side of the two Step tracks, per-branch step counts + total magnitude, and where they diverge. Computed from `TimelineView`. Mark a slot for the future cell-level diff (wired at integration via `inspectStep`/state).
2. **Accessibility** — full keyboard nav (arrows move playhead / select steps; logical Tab order; visible focus; ARIA roles/labels on histogram, sliders, branch tracks, inspector); announce mode changes (Present/Preview) to screen readers. No obvious WCAG violations.
3. **Narrow-pane responsive layout** — the Office task pane is **~320px wide**. The pane must work in a narrow column: collapsible/overflowing controls, vertical stacking, horizontal scroll for the histogram, nothing clipped. Test at ~320px.
4. **Theming** — a `theme` prop (light/dark, token-based via CSS variables); default light; support dark. No host reads.
5. **Histogram richness** — tooltips on bars (on hover AND keyboard focus), a step/index axis with labels, clear "tall bar = big change" emphasis (PRD), subtle transitions on scrub/zoom.
6. **Cleanup**: remove `CODEX-TASK-3.md` at the end.

## Conventions / workflow (unchanged)

- React 19 + strict TS; Vitest + jsdom + @testing-library/react. Conventional Commits, **NO AI-attribution trailers**.
- Tests for: compare divergence computation, keyboard nav moves playhead/selection, narrow layout renders at 320px, the `theme` prop applies, tooltip appears on focus. Coverage ≥80%.
- Green (`typecheck && lint && format:check && test`) then PR into protected `main` (CI `build` + @omototo approval). Use the PR template; note any contract extensions.

## Definition of done

A production-grade, accessible, theme-able, narrow-pane-friendly timeline with a branch-compare view and green tests, PR'd. The cell-level diff is the only deferred piece (integration).
