# Editor end-to-end tests and corpus integration scenarios

This suite has two layers. Chromium drives the production DOCX, XLSX, and PPTX
React editors through the browser-capable desktop app. Bun and Python run a
larger set of engine integration scenarios against pinned corpus documents.

## Browser tests

```bash
bun install --frozen-lockfile
bunx playwright install chromium
bun run test:e2e:browser
```

The command builds the production packages and app, starts a local preview
server, and runs headless Chromium. Each format opens a committed demo document,
edits using browser input, verifies undo and redo, downloads the saved document,
reopens it, and checks that attempting to open a corrupt document preserves the
current editor. The tests inspect the saved OOXML as well as the editor state.
They use bundled fonts and reject external HTTP requests. No deployed service,
credentials, or desktop window is required.

Playwright manages server startup and shutdown. Tests run with one worker and no
retries. Failure screenshots, traces, reports, and saved files go under
`.source/e2e/browser` (`BETTEROFFICE_BROWSER_OUTPUT` overrides). After a build,
`bunx playwright test --config scripts/e2e/browser/playwright.config.ts` reruns
just the browser tests.

This covers browser editing and persistence for all three formats. Browser
collaboration transport/reconnect, additional browsers, and visual fidelity
across the full corpus remain separate coverage areas. The existing
[visual-quality tooling](../office-quality/README.md) covers rendering comparisons.

## Corpus scenarios

```bash
uv tool install maturin==1.14.1
bun scripts/e2e/python-env.ts
bun run test:e2e
bun run test:e2e:harness
bun run typecheck:e2e
```

The Python bootstrap uses Python 3.13 and locked release builds of all three
bindings. The source fingerprint prevents a venv built from a different checkout
from silently satisfying cross-SDK tests. `BETTEROFFICE_E2E_VENV` overrides the
default `~/.cache/betteroffice/e2e-venv`.

There are 37 scenario definitions, each run against three pinned documents of its
format (111 cases, including 18 Python cases). `corpus.ts` records the document
identity, byte count, and SHA256; the asset cache verifies downloaded and cached
bytes. Set `QUALITY_ASSET_CACHE` to override `~/.cache/betteroffice/e2e-assets`.
The first run needs network access to fetch the samples.

The corpus suite exercises editing, formulas, formatting, tables, pagination,
slides, proposals, save/reopen, multi-replica updates, and cross-SDK behavior.
Its `web` actor names mean the WASM SDK running in Bun. Updates are exchanged
in-process; these scenarios do not exercise browser events or a relay. XLSX
`incremental-commits` measures repeated engine commits of string prefixes; the
browser test separately exercises draft typing and commit. DOCX native/WASM
layout parity compares page dimensions and fragment placement from the same
WASM-measured kernel; it does not claim independent native text measurement.

Missing Python bindings fail the standard run and CI. For exploratory local
runs only, `BETTEROFFICE_E2E_ALLOW_SKIP=1 bun run test:e2e` explicitly permits
reported skips. Such a run cannot record or compare baselines. Ordinary
`bun run test` skips the expensive corpus cases and runs the harness tests.

## Results and optional local comparisons

```bash
BETTEROFFICE_E2E_OUTPUT=/tmp/current-e2e bun run test:e2e
bun run test:e2e:record
bun run test:e2e:compare
bun scripts/e2e/report.ts /tmp/current-e2e
bun scripts/e2e/report.ts /tmp/current-e2e --json
bun scripts/e2e/report.ts --diff /tmp/before-e2e /tmp/current-e2e
```

Baselines live in ignored `.source/e2e/baseline`; set
`BETTEROFFICE_E2E_BASELINE` to choose another directory. They are local machine
measurements, not portable benchmark fixtures. Record promotion happens only
after all formats and all required scenarios succeed. Failed, skipped, filtered,
interrupted, or incomplete runs cannot replace the baseline. Partial failure
artifacts still include the expected scenario inventory and failure reasons.

Schema version 3 records the format, commit, dirty-worktree flag, timestamp, machine/runtime metadata,
pinned sample hash, case status, raw operation sequence, actor, stage timings,
and summaries. Failed operations retain their duration and error. Reports and
comparisons reject missing/empty files, unsupported schemas, duplicate cases,
invalid timings, and summaries inconsistent with their raw operations.

Comparison requires identical machine/runtime metadata, sample hashes, scenario
and participant inventories, actor/operation groups, operation counts, and error
counts. It compares each actor separately. A median regression must exceed 25%
and 10 ms (2 ms for operations with at least five observations); p95 is also
checked for groups with at least 20 observations. Stage means use only operations
that actually reported that stage.

These are diagnostic comparisons of scenario observations. Repeated operations
may have different document state or cold/warm behavior; they are not independent
benchmark trials. Run on a quiet machine and inspect raw timings. CI collects
profiles and gates correctness, including browser tests; it does not compare
Apple laptop baselines to Linux runners or claim a reliable performance gate.
A strict performance gate needs repeated base/PR runs on a controlled runner.

## CI

`.github/workflows/e2e.yml` runs two independent jobs on every PR:

- **Corpus scenarios** checks the harness, builds Python and WASM, executes all
  111 required cases, and uploads results, including partial failures.
- **Browser editors** installs Chromium with its Linux dependencies, builds the
  production app, starts the local server, runs all three formats, and uploads
  the Playwright report and failure traces.

Both jobs fail on test failures, have a 40-minute deadline, and cancel superseded
runs. Repository branch protection must require their check names if they are to
block merges; workflow code does not change that repository setting.
