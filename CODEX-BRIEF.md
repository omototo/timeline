# Codex Brief — Task-Pane UI (`feat/task-pane-ui`)

You are building the **React task-pane UI** for the Timeline Excel add-in. Work only on this branch. This brief is self-contained; deeper context is in `docs/` (read `CONTEXT.md`, `docs/adr/0005`, `0006`, `0008`, and `docs/engine-interface.md`).

## Mission

A zoomable, branchable **timeline histogram** in the Office task pane. The user sees their entire edit history as a histogram of "Steps," scrubs through time, previews any past state, and toggles between branches — Fusion 360 meets Git, for a spreadsheet.

## Hard boundaries (do not cross)

- **Own only `packages/addin/src/ui/`.** Do NOT touch `packages/addin/src/excel/` (the Office.js adapters — another stream), `packages/engine/**` (the engine), or the frozen interface in `docs/engine-interface.md`. If you think the engine/interface needs a change, write it in the PR description — do not edit it.
- **No Office.js in the UI.** The UI never imports `Office`/`Excel`. It consumes a view-model and emits commands; the integration layer (later) wires those to the engine + adapters.
- **Build against the fake.** Use `src/ui/sample-timeline.ts` and the `TimelineView` / `dispatch` contract in `src/ui/contract.ts` (provided). Do not depend on the real engine — it lands separately and is wired in at integration.

## The contract (provided in `src/ui/contract.ts`)

The UI is a pure function of a **view-model** plus a **command dispatcher**:

```ts
TimelinePaneProps = { view: TimelineView; dispatch: (cmd: TimelineCommand) => void }
```

- `TimelineView` — branches, each with ordered Steps; each Step has a `magnitude` (bar height), a `kind`, and a `sheetId`; plus `head` (active branch + present/preview + previewStepIndex). See `contract.ts`.
- `TimelineCommand` — what the user can do: `{type:'goto', ref}`, `{type:'returnToPresent'}`, `{type:'branch', from}`, `{type:'switch', branchId}`. The UI **emits** these; it does not implement them.

`sample-timeline.ts` gives you a realistic multi-branch fixture (a main line with a fork) to develop and test against.

## What to build (MVP, in order)

1. **Histogram** — render each branch's Steps as vertical bars; **bar height = `magnitude`** (a 1,000-cell paste towers over a one-cell edit). Color/marker by `kind` (value / structural / worksheet / reconciliation).
2. **Two-axis zoom** (ADR-0005): temporal (zoom/scroll across Steps; bars get denser/sparser) AND structural (a control to filter the histogram to a single `sheetId` — "drill into a worksheet" — or show the whole workbook).
3. **Scrub** — a slider/playhead; dragging emits `goto`. The current `head` position is marked. "Return to Present" button emits `returnToPresent` and is visible only in preview mode.
4. **Branch split view** (ADR-0006) — when more than one branch exists, show them as split tracks with the fork point connected to the parent; clicking a branch emits `switch`. A "Branch from here" affordance on a previewed Step emits `branch`.
5. **Step inspector** — selecting a Step shows its summary (kind, magnitude, sheet, and — later — formula metadata via `inspectStep`). A stub panel is fine for MVP.

Preview _rendering_ onto the sheet (frozen values, ADR-0008) is NOT your concern — the engine/adapters do that. You only emit `goto` and reflect `head.mode === 'preview'`.

## Tech + conventions

- React 19 + TypeScript (strict — match the repo's tsconfig; `verbatimModuleSyntax`, `exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`).
- Tests with **Vitest + jsdom + @testing-library/react** (the addin project is already configured). Cover: bar heights map to magnitude; scrub emits `goto`; branch click emits `switch`; preview-mode shows "Return to Present".
- Keep dependencies lean. If you want a charting lib, prefer hand-rolled SVG/divs for the histogram (full control, no heavy dep) — but if you add one, justify it in the PR.
- **Commit messages: Conventional Commits, and NO `Co-Authored-By` / AI-attribution trailers** (repo rule).

## Workflow

- Branch: `feat/task-pane-ui` (this one). Commit here.
- Before opening the PR: `bun install` then `bun run typecheck && bun run lint && bun run format:check && bun run test` must be green.
- Open a PR into `main`. **`main` is protected**: CI (`build`) must pass and **@omototo must approve** — your PR will not merge without that review. Use the PR template.
- Fill the PR "Notes for reviewer" with: what's done, what's stubbed, and any place you felt the contract was missing something.

## Definition of done (MVP)

A task pane that renders the sample timeline as a zoomable histogram, lets you scrub (emitting `goto`), drill into a worksheet, see a branch split and switch branches, with green tests — all against the fake. Integration with the real engine is a follow-up we handle on merge.
