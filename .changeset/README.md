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

The DOCX, XLSX and PPTX packages share one fixed version; fonts and VSDX
each have their own. The published Rust crates use one
lockstep version independent from npm. Apps and other private packages are
never published.

Skip the changeset for test/docs/CI-only PRs.
