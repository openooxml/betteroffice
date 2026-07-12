# Changesets

This folder holds the release queue. Every code PR that changes a published
package (`@betteroffice/*` under `packages/`) adds a changeset:

```bash
bun changeset
```

Pick the affected packages and the bump (`patch` / `minor` / `major`), then
commit the generated `.changeset/*.md`. On merge to `main`, the Release
workflow opens a `chore: release` PR that applies the bumps and CHANGELOG
entries; merging that PR publishes to npm.

Each format is its own fixed group — `xlsx` and `xlsx-react` version in
lockstep, independent of `docx`. Apps under `apps/` are private and never
published.

Skip the changeset for test/docs/CI-only PRs.
