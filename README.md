# Timeline — Excel Office Add-in

A Bun monorepo for an Excel Office Add-in that reconstructs past sheet state from a Shadow State engine (see `docs/adr/`).

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

## Conventions

- TypeScript strict (strictest practical flags — see `docs/engineering-standards.md`).
- Conventional Commits, enforced on `commit-msg` via commitlint.
- husky + lint-staged format/lint on commit.
