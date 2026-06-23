# Codex Task 4 — Run the add-in in Excel (Office shell + sideload) (`feat/excel-sideload`)

The timeline UI is feature-complete against the fake (PRs #1, #2, #10). Now make it **actually load and run inside Excel as a sideloaded task pane**, demoable against the fake data source. This is the "see it work in Excel" milestone.

## ⚠️ Setup — avoid the collision we had last time

- **Work in your OWN checkout of this branch, not the shared main working directory.** Last round, edits landed uncommitted on `main` in the shared dir and collided. Check out `feat/excel-sideload` in your own clone/worktree and **commit to it regularly.**
- First: `git checkout feat/excel-sideload`, then `bun install`. Run `bun run verify` before every push (it's the exact CI gate).

## Boundaries

- **Own**: `packages/addin/src/main.tsx` (the Office bootstrap), `packages/addin/manifest.xml`, dev/sideload tooling + npm scripts, the README "Run in Excel" docs, and a small Office-theme bridge module.
- **Do NOT touch**: `packages/addin/src/excel/**` (Office.js adapters — another stream), `packages/engine/**`, or the existing `packages/addin/src/ui/**` **components** (they're done — you only mount them).
- The pane consumes a `TimelineDataSource`. **Use `FakeTimelineDataSource`** — the real engine-backed source is the integration team's job and will be swapped in via a one-line provider change. Don't build it.
- `main.tsx` and the theme bridge MAY import Office.js (the add-in host is allowed Office.js; only the `src/ui` components must stay pure).

## Scope

1. **Office bootstrap** — `main.tsx` waits for `Office.onReady` (Excel host) and then mounts the React app (`App` / `TimelinePaneContainer` with `FakeTimelineDataSource`) into the task pane root. Handle the case where Office isn't present (dev fallback) so `bun run --filter @timeline/addin dev` still renders in a plain browser.
2. **HTTPS dev + sideload** — wire `office-addin-dev-certs` so the Vite dev server serves **HTTPS on port 9588** (the manifest already points there). Add scripts (e.g. `start` / `sideload` / `stop`) using `office-addin-debugging` (or document the manual sideload). Make `manifest.xml` dev-ready (it already targets `https://localhost:9588` and `ExcelApi` 1.9; add placeholder icon assets if needed).
3. **Office theme bridge** — read `Office.context.officeTheme` and map it to the UI's existing `theme` prop (light/dark) so the pane matches Excel's theme. This is the one place the bootstrap touches Office.js for the UI.
4. **Docs** — a "Run it in Excel" section in `README.md` (or CONTRIBUTING): the sideload steps for Excel on the web and desktop, and the dev-cert trust step.
5. **Tests** — bootstrap-level tests where feasible (mock `Office.onReady` / `Office.context.officeTheme` in jsdom): the app mounts when Office is ready, the theme bridge maps `officeTheme` → `theme`, and the dev fallback renders without Office.
6. **Cleanup** — remove `CODEX-TASK-4.md` at the end.

## Conventions / workflow (base rules — see CONTRIBUTING.md)

- React 19 + strict TS. Tests with Vitest + jsdom.
- **`bun run verify` green before every push** (typecheck → lint → format:check → test:cov → build).
- Conventional Commits, **signed off** (`git commit -s`), **NO AI co-author/attribution** in commits or the PR.
- Open ONE PR into `main` (protected: CI `build` + @omototo approval; the Opus review bot will also review it). Use the PR template; note anything the integration layer should honor (e.g. how the theme prop is fed).

## Definition of done

`bun run --filter @timeline/addin dev` serves HTTPS on 9588; the manifest sideloads into Excel (web + desktop) and the timeline task pane renders against the fake, matching Excel's light/dark theme; sideload steps documented; green tests; PR'd. Swapping the fake for the real engine is the follow-up integration step (not yours).
