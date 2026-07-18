# Changesets

This folder holds the release queue. Every code PR that changes a published
npm package or Rust crate adds a changeset:

```bash
bun changeset
```

Pick the affected packages and the bump (`patch` / `minor` / `major`), then
commit the generated `.changeset/*.md`. Select `@betteroffice/rust-crates` for
Cargo changes. On merge to `main`, the Release workflow opens a
`chore: release` PR that applies the bumps and CHANGELOG entries; merging that
PR publishes to npm and crates.io.

Each format is its own fixed group — `xlsx` and `xlsx-react` version in
lockstep, independent of `docx`. The eight publishable Rust crates use one
lockstep version independent from npm. Apps and other private packages are
never published.

Skip the changeset for test/docs/CI-only PRs.
