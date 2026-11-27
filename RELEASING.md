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

Trusted Publishing can only be configured after a crate exists, so a crate
that is not on crates.io yet cannot be created with OIDC. Every release
attempts to obtain an OIDC token via `rust-lang/crates-io-auth-action` (pinned to the
official v1 commit) and publishes per crate: a crate that does not exist yet
goes straight to `CRATES_IO_BOOTSTRAP_TOKEN`, while a crate that already
exists tries OIDC first and falls back to the bootstrap token only if the
OIDC token is unavailable or the publish fails. The choice is independent per crate.
`scripts/publish-crates.mjs` passes only the selected token to that crate's
`cargo publish` (under both `CARGO_REGISTRY_TOKEN` and
`CARGO_REGISTRIES_CRATES_IO_TOKEN`, with `CRATES_IO_BOOTSTRAP_TOKEN` removed
from every cargo child environment), rechecks the target version after a
failed OIDC upload before falling back, and records each bootstrap use in the
step summary before the registry wait for that crate can fail.
If a failed attempt is followed by a visible registry version, the summary
reports an unconfirmed publishing credential without attributing it to either
token. Only a successful Cargo publish confirms the credential used.

1. Create a short-lived crates.io token authorized to publish new crates and
   any existing crates that may need the fallback.
2. Add it to the repository as `CRATES_IO_BOOTSTRAP_TOKEN` before merging the
   release PR that introduces the new crate names.
3. Merge the release PR and confirm every crate was published.
4. Add a GitHub Trusted Publisher to each newly created crate with owner
   `openooxml`, repository `betteroffice`, and workflow `release.yml`, before
   its next release.
5. Remove the GitHub secret and revoke the bootstrap token once OIDC works
   for all crates, not just once all names exist. The release records each
   crate it creates with the bootstrap token and each existing crate it
   publishes with the bootstrap fallback separately in the step summary,
   recording each use before that crate's registry wait runs.

`scripts/check-publish-targets.mjs --crates` runs before the crates publish and
fails the release, naming every crate that is not on crates.io yet — unless
`CRATES_IO_BOOTSTRAP_TOKEN` is set, in which case it reports the missing
crates as "will be created with the bootstrap token" and passes.

`release.yml` always attempts the OIDC exchange on the publish path. That
step may continue on failure only when the bootstrap token is detected, so an
OIDC exchange that yields no token falls back per crate as above; without the
bootstrap token an OIDC failure still fails the release. A registry lookup
error fails the release instead of guessing a crate is new.

## Initial npm release of a new package

`release.yml` authenticates to npm with OIDC Trusted Publishing. When the
optional `NPM_FALLBACK_TOKEN` secret is set, npm uses it only for a package whose
OIDC exchange fails (see "npm fallback token"). Trusted Publishing can only be configured on a package
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

## npm fallback token

For a package that has no Trusted Publisher yet, create a granular access token
on npmjs.com with "Read and write (publish and stage)", "Bypass two-factor
authentication", only the packages that need it, and a short expiry. Add it as
the repository secret `NPM_FALLBACK_TOKEN`. The npm publish steps of
`release.yml` and `publish-native.yml` pass it as `NODE_AUTH_TOKEN`; npm still
tries OIDC first and uses the token only where the exchange fails. Delete the
secret after the release and add the missing Trusted Publishers.

## Canary releases

Once a day `release.yml` publishes the pending npm changes as a canary: the
document, React and i18n packages of each format and the font packages, versioned
below the coming release (for example `0.7.0-canary.20261127031700`) under the
`canary` dist-tag. Native packages, the packages that need them, crates and
Python bindings have no canary. A run builds the newest main commit whose CI
passed, publishes nothing when that commit already has a canary or no changeset
is pending, and never moves `latest`.

- Install: `npm install @betteroffice/docx@canary`.
- Trigger: run **Release** with `canary` checked; `dry_run` (on by default) packs
  and install-checks the tarballs and uploads them as the `npm-canary` artifact.
- Turn on or off: the schedule runs only while the repository variable
  `CANARY_RELEASE` is `true`.

## Native packages

`@betteroffice/{docx,pptx,xlsx}-native` are three roots with eight platform packages
each (macOS arm64/x64, GNU Linux arm64/x64, musl Linux arm64/x64, Windows
arm64/x64): 27 npm names.
Each root joins its core's fixed Changesets group. `version-packages.mjs` syncs
the platform manifests and the Cargo version.
Repository roots are private so `changeset publish` skips them. The dispatch-only
`publish-native.yml` drops `private` and uploads with OIDC Trusted Publishing in
environment `npm-<format>`, refusing `NPM_TOKEN` and falling back to
`NPM_FALLBACK_TOKEN` when it is set.
After a successful npm release, `release.yml`'s `native-npm` job dispatches that
workflow for formats whose root version is missing and waits for every run.
The release holds packages that need a native version missing from npm. The
`held-npm` job publishes their tarballs after the required native versions appear.
`RELEASE_SKIP_NATIVE=true` disables native publishing and leaves held packages
unpublished while the required versions are missing.

Trusted Publishing cannot create a package: each of these names has to be on npm
before the release that publishes it. Unless `RELEASE_SKIP_NATIVE` is `true`,
read ignoring case as the workflow reads it, `scripts/check-publish-targets.mjs --npm`
fails the release before anything is published while one of them is not on npm.

### A new format

Publish all nine names of the format by hand first:

1. Run `publish-native.yml` with `dry_run=true` at the release commit. In a clean
   checkout of that commit, download its eight `<format>-native.*` artifacts into
   `bindings/<format>-native/artifacts` and the whole `<format>-native-loader`
   artifact, `index.js` and `dist/`, into `bindings/<format>-native`.
2. Run `bun install --frozen-lockfile`, then `bunx napi artifacts --cwd bindings/<format>-native`,
   then `bun scripts/rewrite-workspace-deps.ts` from the repository root.
3. From `bindings/<format>-native`, run `npm pkg delete private`, then check
   `npm pack --dry-run`: it must list `index.js`, `index.d.ts` and every `dist/`
   file the manifest's `exports` names, and no dependency in `package.json` may
   start with `workspace:`.
4. Run `npm publish --access public --provenance=false`. Its `prepublishOnly` publishes the eight platform packages first.
5. Throw the checkout away: steps 2 and 3 rewrote its manifests.
6. Configure all nine names' Trusted Publisher: owner `openooxml`, repository
   `betteroffice`, workflow `publish-native.yml`, environment `npm-<format>`.
   Restrict each `npm-<format>` environment to the main branch.
   Allow a direct `npm publish` in each: a configuration created after
   2026-09-03 allows only `npm stage publish` unless that is chosen.
   Create it no earlier than two days before the release: npm lets a new
   configuration expire unless it publishes successfully within two days.

### A new platform for existing roots

A new target adds one name per format. Reserve each before the release that
first publishes it; that release publishes the real version.

1. Outside the checkout, make one directory per name with a `README.md` and a
   `package.json`: the name, version `0.0.0`, a description that says it is a
   placeholder, and the `license`, `repository`, `publishConfig`, `os`, `cpu`
   and `libc` of `bindings/<format>-native/npm/<platform>/package.json`. No
   `main` and no `files`: there is no binary yet.
2. Check that `npm pack --dry-run` lists only `package.json` and `README.md`.
3. Run `npm publish --access public --provenance=false` in each directory.
4. Configure each name's Trusted Publisher as in step 6 above.

Not verified:

- That `--provenance=false` on the command line wins over a `publishConfig` that
  asks for provenance. npm documents `provenance=false` in `publishConfig`, in
  `.npmrc` and as `NPM_CONFIG_PROVENANCE=false`.
- What the native publish does when a platform version is already on npm, whether
  it pins the roots' optional dependencies exactly, and under which tag it publishes.
- That the release's publish through Trusted Publishing goes through over a
  placeholder, and that `latest` then moves from `0.0.0` to the released version.
  The two font names reserved this way show only that a first publish at `0.0.0`
  leaves `latest` there.
- That every name the release publishes has its Trusted Publisher configured.

## Python bindings

`scripts/python-bindings.mjs` is the single registry of Python distributions.
`scripts/version-packages.mjs` imports it, and both the CI gate and the release
workflow read it at runtime — `release.yml` turns it into one publish dispatch
per binding — so the list is never duplicated in a workflow.

The standalone workspaces `bindings/`, `fuzz/` and `apps/native-viewer/` each
carry a `Cargo.lock` that records the workspace crates by version, so every bump
invalidates them and CI's `cargo clippy --locked` would fail.
`scripts/rust-crates.mjs` lists them in `STANDALONE_WORKSPACES`;
`scripts/version-packages.mjs` regenerates each lock after writing the new
versions, and the release commit carries them alongside the manifests.
`scripts/standalone-workspaces.test.ts` fails when a committed lockfile is not on
that list or a path dependency on a workspace crate carries a `version`.

Each entry carries a `publish` flag. Every registered binding is versioned by
Changesets and installed and tested by the CI gate; only a flagged one is built
into wheels and an sdist and uploaded. A binding lands with `publish: false`, so
merging it cannot create a PyPI project. Flipping that flag is what arms the
upload, and it is the last step of a launch.

VSDX is held back with `publish: false`. Its npm packages are private and its Rust
crates declare `publish = false`; all three remain available for local builds.

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

### Font pack wheels

`betteroffice-fonts-latin`, `betteroffice-fonts-intl` and `betteroffice-fonts-cjk`
are pure `py3-none-any` wheels of data: each holds one font pack under
`betteroffice_fonts_packs/<pack>/<version>/`, the directory layout the Rust font
manager reads from a local base directory. Every format binding requires all
three at the exact versions it pins in `pyproject.toml`, and its module points
the default font source at the installed directory on import, so no format wheel
embeds a font and nothing is fetched at run time.

`scripts/python_font_packs.py` builds them from `packages/fonts-<pack>/assets`,
and a pack's version is its `manifest.json` version, which the Rust catalog reads
too. `FONT_PACK_NAMES` in `scripts/python-binding-names.mjs` lists them apart
from the binding registry: they have no crate and no Changesets entry. A pack is
pending whenever PyPI does not serve its manifest version, so regenerating a pack
at a new version publishes it with the next release. `scripts/python-bindings.test.ts`
fails when a binding's pins drift from the manifests.

`python-dist.yml` builds a `fonts-<pack>` binding as a single wheel with no Rust
and no sdist. The release dispatches the font packs first and waits for them;
when one fails it dispatches no format binding and fails, because every format
wheel would be uninstallable without it. Launching them is the same pending
publisher setup as a binding, once per pack:

1. Create the GitHub environments `pypi-fonts-latin`, `pypi-fonts-intl` and
   `pypi-fonts-cjk`.
2. Add a pending publisher per project (`betteroffice-fonts-latin`,
   `betteroffice-fonts-intl`, `betteroffice-fonts-cjk`) with owner `openooxml`,
   repository `betteroffice`, workflow `publish-python-binding.yml`, and the
   matching environment.

Until all three exist, the Python release train stops at the first font pack
with an error naming this section, and no format binding is uploaded. A
**Publish a Python binding** dispatch with `binding=fonts-<pack>` publishes one
pack by hand.

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
distribution fails only that distribution's upload; a font pack that fails stops
the format bindings, which require it.

[pending]: https://docs.pypi.org/trusted-publishers/creating-a-project-through-oidc/
[reusable]: https://docs.pypi.org/trusted-publishers/troubleshooting/#reusable-workflows-on-github

## Publish order

The workflow publishes dependencies before consumers, in the order
`RUST_CRATES` lists them in `scripts/rust-crates.mjs`: the shared crates, then
each format's own layers, then the crate that ties them together.

Cargo versions and internal registry requirements live in the root
`Cargo.toml`. `scripts/version-packages.mjs` synchronizes them with the private
Changesets marker in `crates/package.json`.
