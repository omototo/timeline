# Contributing to Excel Parametric Timeline

Thanks for your interest in contributing. This project is open-source under
**Apache-2.0** and developed in the open. The core engine and add-in are free
and gate-free; this repository contains the open-core only.

## Prerequisites

- [Bun](https://bun.sh) (the package manager and task runner for this monorepo).
- A recent Node-compatible toolchain is pulled in via Bun; no separate Node
  install is required to run the scripts below.

## Getting started

```sh
bun install
```

This installs all workspace dependencies and wires up Git hooks (husky, via the
root `prepare` script).

## Monorepo layout

This is a Bun workspace with two packages:

| Package                                | Role                                 | Purity                                             |
| -------------------------------------- | ------------------------------------ | -------------------------------------------------- |
| `packages/engine` (`@timeline/engine`) | Pure history/diff engine.            | **No** Office.js, DOM, or React. Host-agnostic.    |
| `packages/addin` (`@timeline/addin`)   | React 19 + Office.js task pane host. | Wires the engine to Excel; owns all host concerns. |

The engine is the reusable, testable core. The add-in is the Excel host that
adapts the engine to Office.js.

## The engine-purity rule

The engine must stay free of Office.js, the DOM, and React. Keeping it pure is
what makes it portable and unit-testable without an Office host. This is
**enforced two ways** — a violation fails CI, not just review:

1. **Compiler.** `packages/engine/tsconfig.json` sets `"lib": ["ES2023"]` with
   **no `"DOM"`**. Referencing `window`, `document`, or `Office` in the engine
   fails to typecheck.
2. **Lint.** An ESLint `no-restricted-imports` override (see
   `eslint.config.mjs`) bans importing `office-js`, `@microsoft/office-js`,
   `react`, and `react-dom` under `packages/engine/**`.

If you need host data inside the engine, pass it in as plain values — don't
reach for the host API from the engine.

## Scripts

Run these from the repo root:

| Command                             | Action                                      |
| ----------------------------------- | ------------------------------------------- |
| `bun install`                       | Install all workspace deps.                 |
| `bun run typecheck`                 | `tsc -b` across project references.         |
| `bun run lint` / `bun run lint:fix` | ESLint (flat config); `:fix` autofixes.     |
| `bun run format` / `format:check`   | Prettier write / check.                     |
| `bun run test`                      | Vitest (engine = node, addin = jsdom).      |
| `bun run test:cov`                  | Vitest with v8 coverage.                    |
| `bun run build`                     | Typecheck, then build the add-in with Vite. |

Before opening a pull request, run the full gate (mirrors CI exactly):

```sh
bun run verify
```

## Commit messages — Conventional Commits

Commits must follow [Conventional Commits](https://www.conventionalcommits.org).
This is enforced on the `commit-msg` hook via commitlint
(`@commitlint/config-conventional`). Examples:

```
feat(engine): coalesce change bursts into a single Step
fix(addin): cancel echo on the single write choke point
docs: document the engine-purity wall
```

husky + lint-staged also run Prettier (and ESLint `--fix` on `.ts`/`.tsx`) on
staged files at commit time.

## Developer Certificate of Origin (DCO)

This project uses the [Developer Certificate of Origin](https://developercertificate.org)
instead of a CLA. By signing off on a commit, you certify that you wrote the
patch or otherwise have the right to submit it under the project's license.

Sign off every commit by adding a `Signed-off-by` trailer with your real name
and email:

```
Signed-off-by: Jane Doe <jane@example.com>
```

Git can add this for you automatically:

```sh
git commit -s
```

Pull requests whose commits are not signed off will be asked to amend before
merge. There is no separate CLA to sign.

## Base rules (working agreement)

These exist because skipping them has bitten us before — a branch that was
"green locally" but red in CI, and a PR that merged red and broke `main`.

1. **Run `bun run verify` before every push and before opening or updating a
   PR.** It runs the **exact** CI gate in one command — `typecheck → lint →
format:check → test:cov → build`. If `verify` is green, CI will be green.
   (A partial gate that skips `format:check` or `build` is what reddened `main`.)
2. **Never merge a PR with a failing or pending check** — not even with admin
   rights. Red CI on `main` blocks everyone; wait for green.
3. **Update a PR branch one way only** — either GitHub's "Update branch" button
   _or_ a local `git merge origin/main`, never both. Doing both diverges the
   branch and forces a messy reconcile.
4. **One stream owns its files.** Engine → `packages/engine`, Office adapters →
   `packages/addin/src/excel`, UI → `packages/addin/src/ui`. Don't edit another
   stream's files; raise interface changes in the PR instead of editing the
   frozen `docs/engine-interface.md` unilaterally.
5. **Branches are squash-merged and auto-deleted.** Delete your local branch
   after merge (`git branch -D <branch>`); don't leave stale branches/worktrees.
6. **No AI co-authorship or attribution** (Claude or Codex) in commit messages
   or PR titles/bodies.

## Pull requests

1. Branch off the default branch (`main`).
2. **`bun run verify` is green** (rule 1 above).
3. Conventional Commits, signed off (`git commit -s`).
4. Describe the change and link any relevant ADRs in `docs/adr/`.
5. CI green + a maintainer approval before merge (`main` is protected).

By contributing, you agree that your contributions are licensed under the
Apache-2.0 license that covers this project.
