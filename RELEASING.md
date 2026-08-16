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

## Initial npm release of a new package

`release.yml` authenticates to npm with OIDC Trusted Publishing and carries no
`NPM_TOKEN` fallback. Trusted Publishing can only be configured on a package
that already exists, so **a brand-new package name cannot be published by the
workflow** — its first version must be pushed by hand.

Do this *before* merging a release PR that would publish the new name.
`changeset publish` publishes each package independently: a run where the new
package fails auth but its dependents succeed leaves those dependents live on
npm depending on a name that 404s.

1. Confirm the name is new and unclaimed: `npm view <name>` returns E404.
2. Build the publishable artifacts: `bun run build:packages`.
3. Pin workspace dependency ranges: `bun scripts/rewrite-workspace-deps.ts`.
   This publish-path-only script rewrites the source package manifests in
   place, so run it from a clean checkout and restore them after publishing.
4. From the package directory, publish the current version by hand with an
   account that has `publish` rights on the `@betteroffice` scope:
   `npm publish --access public --provenance`.
5. From the repository root, restore the rewritten source manifests with
   `git restore -- packages/*/package.json`, then confirm the worktree is clean.
6. Add a Trusted Publisher to the new package on npmjs.com with owner
   `openooxml`, repository `betteroffice`, and workflow `release.yml`.
7. Verify: `npm view <name> version` matches, and the package page shows the
   provenance attestation.
8. Only then merge the release PR. Subsequent versions publish through OIDC
   like every other package.

An optional peer that 404s does not make installation fail: npm silently omits
it. Consumers then fall back to synthetic metrics while being told to install a
package that does not exist. Nothing in CI enforces this bootstrap ordering. If
the manual publish is skipped and a release runs anyway, publish the missing
package and republish the dependent as soon as possible.

## Python bindings

`scripts/python-bindings.mjs` is the single registry of Python distributions.
`scripts/version-packages.mjs` imports it, and both the CI gate and the release
workflow read it at runtime — `release.yml` turns it into the build and publish
matrices — so the list is never duplicated in a workflow.

`bindings/Cargo.lock` pins the workspace crates by version, so every bump
invalidates it and CI's `cargo clippy --locked` would fail.
`scripts/version-packages.mjs` regenerates it after writing the new versions, and
the release commit carries it alongside the manifests.

Each entry carries a `publish` flag. Every registered binding is versioned by
Changesets, installed and tested by the CI gate, and built into wheels and an
sdist by the release; only a flagged one enters the publish matrix. A binding
lands with `publish: false`, so merging it cannot create a PyPI project.
Flipping that flag is what arms the upload, and it is the last step of a launch.

Every registered binding is at `publish: true` today; none is held back.

### Launching a binding

A **pending** Trusted Publisher is configured before the project exists and
creates it on first upload, converting itself into a normal publisher, so a new
distribution never needs an API token ([PyPI docs][pending]). A pending
publisher reserves nothing until it is used — if someone else registers the
project name first, it is invalidated.

To launch `betteroffice-<format>`:

1. Add `bindings/python-<format>` to the registry with `publish: false`, and give
   the directory a private `package.json` at the version its `Cargo.toml`
   carries, so Changesets versions it. Those are the only two files to write:
   the `members = ["python-*"]` glob in `bindings/Cargo.toml` already covers the
   crate — which is also why every `bindings/python-*` **directory** must be a
   cargo crate, or the whole workspace fails to load — and `bindings/*` is
   already a Bun workspace. Regenerate and commit both lockfiles: `bun install`
   writes `bun.lock`, and `cargo metadata --manifest-path bindings/Cargo.toml`
   writes `bindings/Cargo.lock`. `bun install` never touches the Cargo lock. CI
   fails on either being stale (`bun install --frozen-lockfile`, and
   `cargo clippy --locked` on the bindings workspace).
2. Create the GitHub environment `pypi-<format>`. The publish job runs in
   `pypi-<binding>`, which is what scopes a Trusted Publisher to one project.
3. Add a pending publisher at <https://pypi.org/manage/account/publishing/> with
   project name `betteroffice-<format>`, owner `openooxml`, repository
   `betteroffice`, workflow `release.yml`, and environment `pypi-<format>`.
   Publishing must stay a direct job in `release.yml`: PyPI [cannot name a
   reusable workflow][reusable] as a Trusted Publisher.
4. Flip the registry entry to `publish: true`. The next push to `main` with no
   pending changesets uploads the distribution and converts the publisher. The
   site's PyPI downloads badge reads the same flag, so no other list needs it.

No step needs a `PYPI_API_TOKEN`.

### The API token path

`release.yml` keeps a token path that runs only when the `pypi-<binding>`
environment holds a `PYPI_API_TOKEN`. It exists for `betteroffice-xlsx`, which
shipped on a token before pending publishers were used here. Delete both steps
once that project has a Trusted Publisher; a new binding must never take it.

A PyPI API token is scoped either to a single project or to the entire account,
and a project that does not exist yet cannot be named — so a bootstrap token is
necessarily account-wide. That blast radius is what pending publishers remove.

A token that is kept anyway must be **project-scoped** and held as an environment
secret on `pypi-<binding>`, never at repository or organization scope. The guard
below proves only that a token is not repository- or organization-scoped; nothing
in CI can prove a token is project-scoped, so that part is on whoever adds it.

**Outstanding:** `PYPI_API_TOKEN` is a **repository** secret today, and it
created `betteroffice-xlsx`, so it is almost certainly account-scoped. Delete it
under Settings → Secrets and variables → Actions, revoke it on PyPI, and add a
Trusted Publisher to `betteroffice-xlsx` (owner `openooxml`, repository
`betteroffice`, workflow `release.yml`, environment `pypi-xlsx`). Until it is
gone the release fails at `Refuse a repository-scoped PyPI token`: that job
declares no environment, so anything it can read is repository- or
organization-scoped, and a repository secret reaches every binding's publish leg.

Naming an environment that does not exist does not fail a job — GitHub creates it
empty — so `environment: pypi-<binding>` cannot block a publish by itself. The
`publish` flag and that guard are what block it.

Each binding builds and publishes independently. A broken wheel in one
distribution fails only that distribution's upload.

[pending]: https://docs.pypi.org/trusted-publishers/creating-a-project-through-oidc/
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
