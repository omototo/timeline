# Timeline — Excel Office Add-in

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](./LICENSE)
[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-support-yellow.svg)](https://buymeacoffee.com/omototo)

A Bun monorepo for an Excel Office Add-in that reconstructs past sheet state from a Shadow State engine (see `docs/adr/`). A non-destructive, branchable timeline — Fusion 360-style history with Git-style branching.

**Open-core, Apache-2.0.** The engine and core add-in are free and open source with no licensing gate (ADR-0012). Distributed via both Microsoft AppSource (official binary) and self-host / sideload — see `CONTRIBUTING.md`.

## Quick start

```sh
git clone https://github.com/omototo/timeline.git
cd timeline
bun install
bun run dev      # serves the pane at localhost:9588 against sample data
bun run verify   # full gate (check:text, typecheck, lint, format, tests, build)
```

`bun run dev` serves the task pane against the fake timeline data source, so you
can develop in a plain browser. It uses HTTPS when the Office dev certs are
installed and falls back to plain HTTP otherwise — no certs required for
browser-only work. Run `bun run verify` before opening any PR. To load the
add-in inside Excel, see [Run It In Excel](#run-it-in-excel).

## Layout

| Package                                | Purpose                              | Purity                                                      |
| -------------------------------------- | ------------------------------------ | ----------------------------------------------------------- |
| `packages/engine` (`@timeline/engine`) | Pure history/diff engine.            | No DOM, no Office.js (enforced by tsconfig `lib` + ESLint). |
| `packages/addin` (`@timeline/addin`)   | React 19 + Office.js task pane host. | Wires the engine to Excel.                                  |

## The engine-purity wall

ADR-0001 requires the engine to stay free of Office.js and the DOM. This is enforced two ways:

1. **Compiler:** `packages/engine/tsconfig.json` sets `"lib": ["ES2023"]` with **no `"DOM"`** — referencing `window`, `document`, or `Office` fails to compile.
2. **Lint:** an ESLint `no-restricted-imports` override bans `office-js`, `@microsoft/office-js`, `react`, and `react-dom` under `packages/engine/**`.

## Scripts

| Command                           | Action                                    |
| --------------------------------- | ----------------------------------------- |
| `bun install`                     | Install all workspace deps.               |
| `bun run typecheck`               | `tsc -b` across project references.       |
| `bun run lint`                    | ESLint (typed, flat config).              |
| `bun run format` / `format:check` | Prettier write / check.                   |
| `bun run test`                    | Vitest (engine = node, addin = jsdom).    |
| `bun run test:cov`                | Vitest with v8 coverage (80% thresholds). |
| `bun run build`                   | Typecheck + build the addin with Vite.    |

## Run It In Excel

The task pane runs against the fake timeline data source for local demos. The
manifest targets `https://localhost:9588/index.html`, so Excel sideload requires
the dev server to use the **trusted Office localhost certificate**. Browser-only
development needs no certs (see [Quick start](#quick-start)); Excel sideload
does.

1. Install and trust the Office dev certificate:

   ```sh
   bun run certs
   ```

   The OS will prompt you to trust the certificate — **accept the prompt**. This
   is required so Excel can load the HTTPS dev server. Without the trusted cert,
   `bun run dev` still works in a plain browser (over HTTP), but Excel will
   refuse to load the add-in.

2. Sideload in Excel desktop:

   ```sh
   bun run sideload
   ```

   This starts the HTTPS Vite dev server on port 9588 and opens Excel desktop
   with the add-in loaded. Use `bun run start` when you
   want the Office debugging session, and `bun run stop`
   to remove the sideloaded add-in registration when finished.

3. Run the pane in a plain browser or prepare Excel on the web:

   ```sh
   bun run dev
   ```

4. Sideload in Excel on the web:

   - Open a workbook in Excel on the web.
   - Choose **Home** → **Add-ins** → **More Add-ins** → **Upload My Add-in**.
   - Upload `packages/addin/manifest.xml` while the dev server is running.

The add-in waits for `Office.onReady` inside Excel and falls back to the same
fake-backed pane in a plain browser. Excel's `officeTheme` is mapped to the
pane's `light`/`dark` theme prop at bootstrap.

## Conventions

- TypeScript strict (strictest practical flags — see `docs/engineering-standards.md`).
- Conventional Commits, enforced on `commit-msg` via commitlint.
- husky + lint-staged format/lint on commit.

## Documentation

- [`CONTEXT.md`](./CONTEXT.md) — ubiquitous-language glossary.
- [`docs/adr/`](./docs/adr/) — architecture decision records (the design's spine).
- [`docs/engine-interface.md`](./docs/engine-interface.md) — the Timeline Engine interface spec.
- [`docs/capability-map.md`](./docs/capability-map.md) — what Excel features the timeline can/can't track.
- [`docs/office-js-findings.md`](./docs/office-js-findings.md) — Office.js API investigation.
- [`docs/engineering-standards.md`](./docs/engineering-standards.md) — stack and conventions.

## Support

If this project is useful to you, you can support its development:

☕ [**Buy Me a Coffee** — buymeacoffee.com/omototo](https://buymeacoffee.com/omototo)
