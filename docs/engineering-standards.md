# Engineering Standards

Foundation conventions for the project, grounded in current upstream docs (verified via Context7, 2026-06). Keep this in sync with the configs; the configs are the enforcement, this is the rationale.

## Stack

| Concern | Choice | Notes |
|---|---|---|
| Package manager / runtime | **Bun** | workspaces, fast installs, script runner |
| Repo | **monorepo** | `packages/engine` (pure) + `packages/addin` (host) — ADR-0001 enforced by build |
| Build / dev server | **Vite** | ADR-0011; HTTPS dev server for sideloading |
| Language | **TypeScript strict** | strictest practical flags (below) |
| Lint | **ESLint flat config + typescript-eslint** | type-checked rules |
| Format | **Prettier** | `eslint-config-prettier` disables conflicting lint rules |
| Tests | **Vitest** (`test.projects`) | `node` env for engine, `jsdom` for addin; `bun test` optional for engine |
| Hooks | **husky + lint-staged** | format/lint/typecheck on commit |
| CI | **GitHub Actions** | typecheck → lint → format:check → test (coverage) → build |

## TypeScript — strictest practical

Base `tsconfig` enables, beyond `"strict": true`:
`noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`, `noImplicitOverride`, `noFallthroughCasesInSwitch`, `noImplicitReturns`, `noUnusedLocals`, `noUnusedParameters`, `verbatimModuleSyntax`, `isolatedModules`, `forceConsistentCasingInFileNames`.

**The engine-purity wall (build-enforced):** `packages/engine/tsconfig.json` sets `"lib": ["ES2023"]` with **no `"DOM"`** — referencing `window`, `document`, or `Office` fails to compile. The addin tsconfig adds `"DOM"`, `"DOM.Iterable"`, and `jsx: react-jsx`. This makes "the engine imports zero Office.js" (Candidate 1) a compiler error, not a convention.

## ESLint — typed linting (typescript-eslint best practice)

- Flat config via `defineConfig` from `eslint/config`.
- Extend `js.configs.recommended`, `tseslint.configs.strictTypeChecked`, `tseslint.configs.stylisticTypeChecked`.
- Enable type information with `languageOptions.parserOptions.projectService: true` (preferred over manual `project: [...]` arrays in monorepos).
- **Engine override:** `no-restricted-imports` bans `office-js`, `@microsoft/office-js`, `react`, `react-dom` inside `packages/engine/**` — a second guard alongside the tsconfig lib wall.
- **Addin override:** `eslint-plugin-react`, `react-hooks`, `jsx-a11y`.
- `@vitest/eslint-plugin` on test files. `eslint-config-prettier` last.

## Vitest — projects

One root config with `test.projects`: an `engine` project (`environment: 'node'`) and an `addin` project (`environment: 'jsdom'`). Coverage via `@vitest/coverage-v8` with thresholds (start 80% global; raise as the engine matures — the engine is pure and should approach 100%).

## Office.js hard limits (shape the host adapters)

Verified from the Office.js resource-limits doc — these are **not negotiable** and constrain `OfficeChangeSource` / `RenderTarget`:

- **5 MB payload limit** per request/response on Excel on the web. A large read-back (the Shadow State diff source) must be **chunked** below this.
- **5,000,000-cell limit** for a single range read. Validate `range.cellCount` before `context.sync()`; split with `RangeAreas` or tiling when exceeded.
- **`untrack()` proxy objects** after use in large batch loops — proxies live until `sync()` and otherwise grow memory unboundedly. The choke-point writer and the read-back path must untrack aggressively.
- **Batch before `sync()`** — minimize round-trips (each `sync()` is the real latency cost, per ADR-0001/ADR-0004).
- **Manifest requirement sets** — declare the minimum `ExcelApi` version the capture features need; gates installability to hosts that actually support the events we rely on.

## ExcelApi version policy (from the Office.js investigation)

Manifest **install-time floor = ExcelApi 1.9** (structural `changeType` at 1.7; single-cell `onChanged.details` fast path at 1.9). Do NOT raise the manifest floor for newer APIs — **runtime feature-detect** them with `Office.context.requirements.isSetSupported('ExcelApi','<v>')` and degrade gracefully:

| API | Set | Used for | Fallback below |
|---|---|---|---|
| `triggerSource` (`ThisLocalAddin`) | 1.14 | echo cancellation (ADR-0002) | expected-write-set |
| `changeDirectionState` | 1.14 | Structural Delta shift direction | infer from changeType + address |
| `valuesAsJson` | 1.16 | lossless linked/entity cell capture (ADR-0008) | flatten value + emit Fidelity Caveat |

See `docs/office-js-findings.md` for the full verification and the spike list (delete-path reference adjustment, undo behaviour, event fan-out, IndexedDB eviction).

## Commits

Conventional Commits (`@commitlint/config-conventional`), enforced on `commit-msg`. Keep commits small and scoped to one module/seam.
