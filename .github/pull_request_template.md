## Summary

<!-- What does this PR change, and why? Link the related issue (e.g. Closes #123). -->

## Checklist

- [ ] PR title is a Conventional Commit (e.g. `feat(engine): ...`, `fix(addin): ...`)
- [ ] `bun run typecheck && bun run lint && bun run format:check && bun run test` pass locally
- [ ] Tests added or updated for the change
- [ ] If `packages/engine` was touched, it stays **pure** (no Office.js / DOM / React imports)
- [ ] If the design changed, the relevant ADR / `CONTEXT.md` / `docs/engine-interface.md` is updated

## Notes for the reviewer

<!-- Anything specific you'd like looked at, trade-offs, or follow-ups. -->
