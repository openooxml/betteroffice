# Releasing

Changesets drive npm and crates.io releases through the same release PR.

## Changesets

Run `bun changeset` and select every affected npm package. Select
`@betteroffice/rust-crates` when a change affects the published Rust API or
implementation. Every Rust crate in `scripts/rust-crates.mjs` versions and
publishes in lockstep, independently from npm versions.

Merging a changeset opens or updates `chore: release`. Merging that release PR
publishes every unpublished npm and Cargo version. The Cargo publisher checks
crates.io before each upload, so rerunning a partial release resumes safely.

## Initial crates.io release

The first publication requires a crates.io API token because Trusted
Publishing can only be configured after a crate exists.

1. Create a short-lived crates.io token authorized to publish new crates.
2. Add it to the repository as `CRATES_IO_BOOTSTRAP_TOKEN` before merging the
   initial release PR.
3. Merge the release PR and confirm every crate was published.
4. Add a GitHub Trusted Publisher to each crate with owner `openooxml`,
   repository `betteroffice`, and workflow `release.yml`.
5. Remove the GitHub secret and revoke the bootstrap token.

Subsequent releases use `rust-lang/crates-io-auth-action` and GitHub OIDC to
obtain a short-lived crates.io token.

`scripts/check-publish-targets.mjs --crates` runs before the crates publish and
fails the release, naming every crate that is not on crates.io yet — unless
`CRATES_IO_BOOTSTRAP_TOKEN` is set, which is the one credential that can create
a crate.

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
package that does not exist. `scripts/check-publish-targets.mjs --npm` fails the
release, naming every package that is not on npm yet, so a skipped bootstrap
stops the run instead of half-publishing it. Both guards run before the first
upload of either registry — a guard that failed after the crates went out would
leave the release half-published, and a crate cannot be unpublished. The
versions it reads are the ones about to be published: the publish path runs only
when no changesets are pending, and `changesets/action` runs no `version`
command then.

## Python bindings

`scripts/python-bindings.mjs` is the single registry of Python distributions.
`scripts/version-packages.mjs` imports it, and both the CI gate and the release
workflow read it at runtime — `release.yml` turns it into one publish dispatch
per binding — so the list is never duplicated in a workflow.

`bindings/Cargo.lock` pins the workspace crates by version, so every bump
invalidates it and CI's `cargo clippy --locked` would fail.
`scripts/version-packages.mjs` regenerates it after writing the new versions, and
the release commit carries it alongside the manifests.

Each entry carries a `publish` flag. Every registered binding is versioned by
Changesets and installed and tested by the CI gate; only a flagged one is built
into wheels and an sdist and uploaded. A binding lands with `publish: false`, so
merging it cannot create a PyPI project. Flipping that flag is what arms the
upload, and it is the last step of a launch.

Every registered binding is at `publish: true` today; none is held back.

### One publish path

`publish-python-binding.yml` is the only workflow that uploads to PyPI. A
Trusted Publisher is scoped to the workflow file that runs the upload, because
that is what the OIDC claim names, and every publisher here names that file — so
`release.yml` cannot upload from a job of its own. Its `python-pypi` job
dispatches `publish-python-binding.yml` once per publishable binding with
`dry_run=false` and `sha=<the release commit>`. The dispatch API takes only a
branch or tag as its `ref`, and `main` moves during a release, so that `sha` is
what every checkout in the build uses — the wheels carry the versions just
released, not whatever landed on `main` meanwhile.

The release then waits for each dispatched run and fails if one of them does, so
a green **Release** means the wheels are on PyPI. Each upload is still its own
**Publish a Python binding** run rather than a leg of the release run, and the
same dispatch by hand — with the `sha` you mean, which is required — is how a
wheel that a release dropped gets filled in.

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
   `betteroffice`, workflow `publish-python-binding.yml`, and environment
   `pypi-<format>`. Publishing stays a direct job in that file: PyPI [cannot name
   a reusable workflow][reusable] as a Trusted Publisher, which is why the
   release dispatches it instead of calling it.
4. Flip the registry entry to `publish: true`. The next push to `main` with no
   pending changesets uploads the distribution and converts the publisher. The
   site's PyPI downloads badge reads the same flag, so no other list needs it.

No step needs a `PYPI_API_TOKEN`.

### No API token path

Neither workflow can upload with a token. The release train fails at `Refuse a
repository-scoped PyPI token` before it dispatches anything: that job declares no
environment, so anything it can read is repository- or organization-scoped, and a
repository secret would reach every binding. `publish-python-binding.yml` repeats
the guard inside `pypi-<binding>`, where an environment secret is visible too.

A PyPI API token is scoped either to a single project or to the entire account,
and a project that does not exist yet cannot be named — so a bootstrap token is
necessarily account-wide. That blast radius is what pending publishers remove.

**Outstanding:** `betteroffice-xlsx` was created by such a token before pending
publishers were used here, and it has no Trusted Publisher — its wheels carry no
provenance, unlike `betteroffice-docx` and `betteroffice-pptx`, whose
attestations name `publish-python-binding.yml`. Add one (owner `openooxml`,
repository `betteroffice`, workflow `publish-python-binding.yml`, environment
`pypi-xlsx`) or its upload fails the OIDC exchange.

Naming an environment that does not exist does not fail a job — GitHub creates it
empty — so `environment: pypi-<binding>` cannot block a publish by itself. The
`publish` flag and that guard are what block it.

Each binding builds and publishes independently. A broken wheel in one
distribution fails only that distribution's upload.

[pending]: https://docs.pypi.org/trusted-publishers/creating-a-project-through-oidc/
[reusable]: https://docs.pypi.org/trusted-publishers/troubleshooting/#reusable-workflows-on-github

## Publish order

The workflow publishes dependencies before consumers, in the order
`RUST_CRATES` lists them in `scripts/rust-crates.mjs`: the shared crates, then
each format's own layers, then the crate that ties them together.

Cargo versions and internal registry requirements live in the root
`Cargo.toml`. `scripts/version-packages.mjs` synchronizes them with the private
Changesets marker in `crates/package.json`.
