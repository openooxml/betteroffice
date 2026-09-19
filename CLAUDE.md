# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

@AGENTS.md

## Review artifacts

Follow the [review artifact rules](AGENTS.md#review-artifacts). Attach screenshots,
before/after comparisons, recordings, and logs to the PR; never commit them as
review evidence. Keep required product assets and automated test fixtures or
golden baselines. If uploads are unavailable, say so and omit the attachment.

## Commands

Package manager is **bun**, not npm.

```bash
bun install

bun run dev              # apps/web
bun run dev:demo         # apps/demo
bun run dev:docs         # apps/docs

bun run test             # builds xlsx+docx wasm, then runs packages, apps/web/app, apps/demo/lib, scripts
bun run typecheck        # every workspace
bun run rust:check       # cargo fmt --check && clippy -D warnings && cargo test
```

### wasm must be built before tests

The generated wasm bindings under `packages/*/src/wasm/generated/` are **gitignored build
artifacts**. `bun run test` builds xlsx and docx first via `pretest`; running `bun test <path>`
directly does not. A failure like `Cannot find module './generated/xlsx_wasm.js'` means the
bundle is missing, not that the code is broken:

```bash
bun run build:xlsx-wasm   # also: build:docx-wasm, build:pptx-wasm, build:vsdx-wasm
```

Requires `wasm-pack` **0.15.0** exactly (CI pins it; do not bypass the pin).

### Running a single test

```bash
bun test packages/vsdx/src/wasm/loader.test.ts   # one file
bun test packages/xlsx -t "formula precedence"   # by name

cargo test -p betteroffice-vsdx-parse                       # one crate
cargo test -p betteroffice-vsdx --test facade               # one integration target
cargo test -p betteroffice-vsdx-edit --features wasm        # wasm-gated tests
```

Rust tests behind `#[cfg(target_arch = "wasm32")]` never compile under a normal `cargo test` —
they will silently look like they passed. Check what actually ran.

### Building distributables

```bash
bun run build:packages   # xlsx, docx, pptx, vsdx TS packages (each builds its wasm first)
bun run build            # apps/web
```

## Architecture

Four document formats — **docx, xlsx, pptx, vsdx** — each implemented as the same layered
pipeline. Learning one teaches the others.

```
crates/<fmt>-parse    file bytes  <->  lossless document model
crates/<fmt>-edit     yrs CRDT session, mutation policy, transactional save
crates/<fmt>-render   model -> display list   (docx uses docx-layout; xlsx/vsdx add -raster)
crates/<fmt>-wasm     wasm-bindgen surface consumed by the TS package
crates/betteroffice-<fmt>   native Rust facade — the public crates.io API
packages/<fmt>        framework-free TS core wrapping the wasm bundle
packages/<fmt>-react  drop-in React editor
packages/<fmt>-i18n   translations
```

Shared, format-agnostic crates sit underneath:

- `ooxml-opc` — OPC/zip container reader and writer with zip-bomb and path-traversal limits,
  plus the package sanitizer and content-type/document-kind detection
- `ooxml-drawingml` — colors, themes and geometry
- `ooxml-text` — text shaping
- `ooxml-redact` / `ooxml-redact-cli` — redaction, behind a per-format enum

`bindings/` is a **separate cargo workspace** (its own `Cargo.toml` and lockfile) exposing
PyO3/maturin bindings. The root workspace is `crates/*` only, so `cargo test` at the root does
not build it.

`apps/` holds the Next.js sites (`web`, `docs`, `demo`) plus two Cloudflare Workers
(`relay` for collaboration, `redact-worker`).

### Cross-cutting patterns

**Lossless round-trip is the core invariant.** Parsers preserve unknown XML, attribute order
and source ordering, so unmodified parts re-serialize byte-identically. Saving is a lexical
span patcher over the intact source part, not generic reserialization — unchanged parts are
byte-identical; changed parts preserve untouched spans.

**Editing goes through a mutation policy**, never raw value writes. Ops emit formulas; guarded,
redirected or locked cells refuse the edit and report a typed receipt.

**Collaboration is yrs CRDT** with staged-clone validation: a remote update is applied to a
clone, validated, and adopted only if valid — so a rejected update leaves state untouched.

**One state, two projections.** The live render projection and the save projection must agree.
Nine consecutive review rounds on vsdx each found a different instance of them disagreeing;
they are now unified behind a shared serializability predicate. When touching identity,
ordering or cell values, verify both projections still agree, and prefer making divergence
impossible by construction over adding another check.

**Public capability copy is a contract.** `apps/web/public/llms.txt`, `apps/web/app/content.ts`
and `README.md` state what ships, and tests pin some of it. Any capability added or removed
updates them in the same commit.

## Format-specific notes

### VSDX

Visio is **not** DrawingML — it is a ShapeSheet model: inches (not EMU), Y-up from bottom-left,
and every cell carries an independent formula `F` and cached value `V`. See `openspec/vsdx-plan.md`
for the format reality check and phase history, and `openspec/vsdx-overnight.md` for the
environment and full gate.

Tests read a **proprietary corpus** of real diagrams via `VSDX_CORPUS_DIR`. Those files must
never be committed — the repo is open source. Tests skip silently when the variable is unset,
which loses the most valuable coverage, so always set it. Committed fixtures are synthetic,
generated by `scripts/create-vsdx-fixture.ts` with a fixed zip date for deterministic bytes.

Gate every vsdx crate, including `-p betteroffice-vsdx` — a narrower gate hid a failing facade
test for nine review rounds.

## Environment

`cargo`, `rustc` and `wasm-pack` live in `~/.cargo/bin`, which may not be on PATH in a fresh
shell:

```powershell
$env:PATH = "$env:USERPROFILE\.cargo\bin;$env:PATH"
```

Two `crates/docx-layout/tests/display_list.rs` snapshot tests fail on any Windows checkout —
`core.autocrlf=true` with no `.gitattributes`, so LF-stored snapshots check out as CRLF. Linux
CI passes. Do not "fix" this from a feature branch.
