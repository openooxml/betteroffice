# Releasing

Changesets drive npm and crates.io releases through the same release PR.

## Changesets

Run `bun changeset` and select every affected npm package. Select
`@betteroffice/rust-crates` when a change affects the published Rust API or
implementation. The eight Rust crates version and publish in lockstep,
independently from npm versions.

Merging a changeset opens or updates `chore: release`. Merging that release PR
publishes every unpublished npm and Cargo version. The Cargo publisher checks
crates.io before each upload, so rerunning a partial release resumes safely.

## Initial crates.io release

The first publication requires a crates.io API token because Trusted
Publishing can only be configured after a crate exists.

1. Create a short-lived crates.io token authorized to publish new crates.
2. Add it to the repository as `CRATES_IO_BOOTSTRAP_TOKEN` before merging the
   initial release PR.
3. Merge the release PR and confirm all eight crates were published.
4. Add a GitHub Trusted Publisher to each crate with owner `openooxml`,
   repository `betteroffice`, and workflow `release.yml`.
5. Remove the GitHub secret and revoke the bootstrap token.

Subsequent releases use `rust-lang/crates-io-auth-action` and GitHub OIDC to
obtain a short-lived crates.io token.

## Python bindings

`scripts/python-bindings.mjs` is the single registry of published Python
distributions. `scripts/version-packages.mjs` imports it, and both the CI gate
and the release workflow read it at runtime — `release.yml` turns it into the
build and publish matrices — so the list is never duplicated in a workflow.

To launch `betteroffice-docx`:

1. Add `bindings/python-docx` to `PYTHON_BINDINGS`. That is the only code
   change: CI installs and tests it, and the release builds, packages, and
   publishes it.
2. Publish once with `PYPI_API_TOKEN`, because a Trusted Publisher can only be
   configured after the project exists.
3. Create the GitHub environment `pypi-docx`. The publish job runs in
   `pypi-<binding>`, which is what scopes a Trusted Publisher to one project.
4. Add a Trusted Publisher to the PyPI project with owner `openooxml`,
   repository `betteroffice`, workflow `release.yml`, and environment
   `pypi-docx`. Publishing must stay a direct job in `release.yml`: PyPI
   [cannot name a reusable workflow][reusable] as a Trusted Publisher.

Removing `PYPI_API_TOKEN` switches every binding to Trusted Publishing, so
configure it for each project before dropping the secret.

Each binding builds and publishes independently. A broken wheel in one
distribution fails only that distribution's upload.

[reusable]: https://docs.pypi.org/trusted-publishers/troubleshooting/#reusable-workflows-on-github

## Publish order

The workflow publishes dependencies before consumers:

```text
betteroffice-opc, betteroffice-xlsx-model
betteroffice-xlsx-parse, betteroffice-xlsx-calc, betteroffice-xlsx-render
betteroffice-xlsx-ops, betteroffice-xlsx-raster
betteroffice-xlsx
```

Cargo versions and internal registry requirements live in the root
`Cargo.toml`. `scripts/version-packages.mjs` synchronizes them with the private
Changesets marker in `crates/package.json`.
