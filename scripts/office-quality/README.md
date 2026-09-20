# Office visual quality harness

Compare DOCX, PPTX, and XLSX renders against Microsoft Office. Inputs and generated files stay in ignored `.source/office-quality/`; references come from the public corpus.

## Local benchmark

Run from the repository root with Bun, Node.js, Python 3.13, and the project's Rust/Wasm toolchain installed:

```sh
bun install --frozen-lockfile
bun run build:packages
bunx playwright install chromium
python3 -m venv .source/office-quality/venv
.source/office-quality/venv/bin/pip install -r scripts/office-quality/requirements.txt

QUALITY_PYTHON=.source/office-quality/venv/bin/python \
QUALITY_OUTPUT=.source/office-quality/run \
  node scripts/office-quality/run.mjs
node scripts/office-quality/readme.mjs .source/office-quality/run/report.json
```

The runner compares the latest npm releases with the checked-out source using pinned CDN fonts. Commit source changes first and choose an empty output directory. The single [`office-quality` collection](https://corpus.betteroffice.dev/collections/office-quality.json) contains all three formats, including PPTArena. Set `QUALITY_FORMAT=docx`, `pptx`, or `xlsx` to measure one format, or `QUALITY_SAMPLES='["betteroffice-demo"]'` for explicit samples. Runs support up to 1,000 samples.

DOCX references and page-bounds profiles support 1–250 pages. PPTX, XLSX, and VSDX references remain limited to 100 pages, matching their capture harness. Extra rendered DOCX pages retain their native dimensions and contribute to the page-count penalty. Source documents may be up to 128 MiB; individual reference images remain limited to 32 MiB. Assets are verified against their recorded sizes and SHA-256 hashes before capture.

Corpus downloads retry transient failures (HTTP 408/429/5xx plus connection and body interruptions) with bounded attempts, exponential backoff with jitter, and capped `Retry-After` waits. Set `QUALITY_ASSET_CACHE` to a directory outside `QUALITY_OUTPUT` to reuse verified source and reference assets across runs; collection, metadata, and npm registry responses are always fetched fresh, and each cached blob is re-verified by exact size and SHA-256 on every hit. CI restores and saves this cache under runner temp storage. A persistent repair marker changes the archive key when corrupt blobs are replaced; unchanged warm caches retain their key.

Capture and comparison failures are recorded per sample and channel without stopping the remaining samples. The report includes scored/total coverage and concise failure reasons; the generated README shows scores and coverage. Failed comparisons have no SSIM and are excluded from the means; differing coverage can make channel means incomparable. Source/reference metadata, downloads, and hash validation still fail the run before capture. Invalid successful comparison records cannot be published.

XLSX capture uses `printDisplayList` when available, with font metrics measured
at 72 layout DPI for the frozen Mac Office capture and the worksheet's explicit defaults.
The Normal style resolves through its built-in ID and `xfId`; missing declarations fall back to style XF zero, then font zero.
Older packages use the screen-range capture. Both keep the recorded ranges,
150-DPI output, page margins, and SSIM scoring. Per-capture metadata identifies
the mode and metrics. Unspecified column defaults, locale-specific dates, and
chart typography can still differ from Excel.

## Manual CI and generated README

After the [workflow](../../.github/workflows/visual-fidelity.yml) lands on `main`, use **Actions → Benchmarks → Run workflow**, or:

```sh
gh workflow run visual-fidelity.yml --ref main -f branch=main
```

Keep `--ref main`; set `branch` to the repository branch to measure. The optional `samples` input selects explicit samples. A preparation job freezes the source revision, package versions, the published DOCX package's npm `gitHead`, and corpus metadata. DOCX, PPTX, and XLSX then evaluate in parallel, each with its own asset cache and 90-minute budget. Each browser capture retains its 600-second deadline. Two independent native build jobs feed up to four DOCX benchmark jobs, with whole documents distributed by reference page count. Every timed competitor for a document runs sequentially on the same worker; browser capture and compilation run on other workers.

The final job requires every selected format and sample, combines the original comparisons, and updates the [README scores](../../README.md#benchmarks) as `openooxml-bot[bot]`. Unchanged results create no commit; a changed branch head requires a rerun. Runs are manual only. Score JSON and generated Markdown are retained for 30 days; per-format reports and optional commit PNGs are transferred as seven-day artifacts.

## Native DOCX and LibreOffice

The DOCX table adds LibreOffice fidelity and a native CLI render-time row. BetterOffice fidelity still measures the browser renderer at 150 DPI. Native timings exercise the Rust rasterizer at 96 DPI; they do not claim that it has the browser renderer's fidelity or feature coverage. PPTX and XLSX keep their existing fidelity comparisons.

Both BetterOffice executables use the same [`native/main.rs`](native/main.rs) host, compiled in release mode (`opt-level=3`, thin LTO), against the frozen current source and the exact source of the published npm version. The host reads the original DOCX, imports it into the engine, loads fonts, lays out the document, constructs the display list, rasterizes page one, PNG-encodes it, and writes the file. All of that work and process startup are timed. No prepared layout, display list, resident engine, Wasm, or browser is supplied to the executable. Unsupported raster operations, skipped images, blank output for a nonblank reference, and invalid output dimensions count as failures.

LibreOffice uses the official prebuilt Linux 26.2.3.2 release, verified against its pinned SHA-256, on Ubuntu 24.04. Its native CLI imports the same DOCX and exports page one directly through [`writer_png_Export`](https://help.libreoffice.org/latest/en-US/text/shared/guide/graphic_export_params.html). Each document has an isolated LibreOffice profile, initialized during the discarded warmup. No resident office server is used. PNG export requests the reference page's physical size at 96 DPI; validation permits the native exporters' one-pixel rounding difference without resampling.

The driver discards one warmup per engine/document, then starts five fresh processes per engine, rotating execution order. Timings include process startup and file I/O; validation, scoring, installation, and builds are outside the timer. OS file caches remain warm. The row is an arithmetic mean of per-document means, using only documents that complete every trial in all three engines. `Timed/total` reports each engine's coverage, and the footnote states the common subset. A failed trial removes that document from all three aggregate means; it never becomes a zero or a partial average. There are no speedup thresholds or cached timing results. Hosted-runner timing varies, so compare columns within the same run rather than treating small changes between runs as regressions.

The current source's bundled fonts are frozen once and shared by both native builds and LibreOffice. Linux workers use a Fontconfig configuration restricted to that bundle, including its Word-family aliases. Build metadata records source/binary/host hashes, Rust version and optimization settings; reports record fonts, LibreOffice build, runner image, architecture, and comparison-library versions. Local macOS LibreOffice uses system font discovery and reports that distinction.

LibreOffice fidelity separately exports the entire document to PDF, rasterizes every page through the pinned PyMuPDF version at 150 DPI, and uses the same Word references, grayscale SSIM, page penalties, and at-most-one-pixel edge adjustment as the existing DOCX comparison. It does not replace the Office references. Sources and references are hash-verified. Scoring omits redundant HTML-gallery PNGs and logs page progress; individual LibreOffice renders remain in artifacts. All selected documents are attempted on every run. Failures stay visible in coverage and diagnostics.

The reconciler requires every planned shard and document, matching plan hashes, builds, font bundles, versions, and measurement settings before updating the README or R2. Download `docx-benchmark-diagnostics-*` for per-trial durations, commands, logs, first-page native PNGs, and full LibreOffice renders. These artifacts expire after seven days. `docx-native-*` contains the executable builds and font bundle. Artifacts are never committed.

To run a local shard after preparing `plan.json` and checking out its published `gitHead`:

```sh
bun scripts/office-quality/native-fonts.ts .source/native-fonts
node scripts/office-quality/build-native.mjs . .source/native/commit
node scripts/office-quality/build-native.mjs /path/to/published-checkout .source/native/published
QUALITY_PLAN=.source/office-quality/plan/plan.json QUALITY_SHARD=0 \
  QUALITY_OUTPUT=.source/docx-benchmark node scripts/office-quality/docx-benchmark.mjs
QUALITY_OUTPUT=.source/docx-benchmark QUALITY_NATIVE=.source/native \
  QUALITY_FONTS=.source/native-fonts/manifest.json \
  python3 scripts/office-quality/docx_benchmark.py
```

Set `SOFFICE` to the native LibreOffice executable when it is not available as `libreoffice`. Use a fresh output directory for each run. `docxShards(plan)` in `docx-benchmark.mjs` lists the complete shard assignment; small explicit selections create fewer workers automatically.

## Published renders and the viewer

With `publish_renders` left on, the final job uploads the reconciled `commit` channel to the
`betteroffice-fidelity` R2 bucket, keyed by the measured commit:

```text
renders/<sha>/<sample>/page_0001.png
renders/<sha>/report.json
renders/latest.json
```

`latest.json` names the current SHA, its report key, and the published page count per sample. It is
written last, so it never points at an incomplete upload. Every published SHA is kept.

Missing format reports or incomplete render artifacts prevent publication. An R2 upload failure
also stops the final job before the README update. [`apps/fidelity`](../../apps/fidelity)
serves those renders next to the public Office references, so a page can be compared by swiping or
by a difference blend. It reads scores only from `report.json` and never derives its own.

## Office references

Requires macOS and desktop Word, PowerPoint, or Excel. Run exports serially into fresh directories:

```sh
.source/office-quality/venv/bin/python scripts/office-quality/reference.py \
  /path/to/example.docx --out .source/office-quality/example/office
```

The extension selects the app. Output includes `reference.pdf`, numbered PNGs at 150 DPI, and `result.json` with source hashes, Office/macOS versions, font names, UTC timestamps, and export status. The timeout defaults to 120 seconds; override it with `--timeout`.

For XLSX comparisons, add `--xlsx-profile /path/to/profile.json`. The profile specifies `scale_percent`, `margin_pt`, and `pages` with zero-based `sheet` indices and A1 `range` values. Each range must fit one page. The demo's nine-page profile is stored in its metadata under `reference.capture_profile`. This measures worksheet range rendering, not automatic print pagination. The print API ignores frozen panes; the screen-range fallback rejects them. Print settings apply only to the temporary workbook, leaving source bytes unchanged.

## macOS permissions

Verified with Office 16.112.3 on macOS 26.2.

- In **System Settings → Privacy & Security → Automation**, allow the app running the command to control Word, PowerPoint, or Excel.
- Allow Office-container access if prompted. The exporter stages both input and PDF inside `~/Library/Containers/<Office bundle>/Data/Documents/BetterOfficeBenchmark/`. This avoided Word's repeated per-file grants without changing Full Disk Access. Successful staging directories are removed; failed exports retain them for inspection.
- Complete Office's first-run screens before exporting. PowerPoint's welcome screen can cause errors `-9074` or `-2710`. Check for open dialogs after a timeout.

The scripts use local PDF export. For manual Word exports, choose **Best for printing** or **Print → Save as PDF** to keep processing local.

## Compare local captures

Capture BetterOffice using the [DOCX instructions](../docx-quality/README.md). For PPTX/XLSX/VSDX, set `QUALITY_FORMAT=pptx`, `xlsx`, or `vsdx` on the shared server and use the matching `?format=` in the capture URL. Document bytes stay local; external browser requests are limited to pinned font files.

`reference.py` exports no Visio references and the benchmark run covers DOCX, PPTX, and XLSX only, so captured VSDX pages have no baseline to score against. `apps/demo/public/betteroffice-demo.vsdx` is a committed diagram to capture. Pages composite onto white because Visio pages carry no painted background.

DOCX references record each PDF page's physical and raster bounds. The capture profile permits a one-pixel canvas extent difference on each axis, adding white space or clipping at the right/bottom edge without moving or resampling rendered content. Larger differences fail; extra renderer pages retain their native bounds. Captures record the original and output dimensions for each page.

For DOCX and XLSX, load the reference profile before running the capture command. Collection runs do this automatically:

```sh
export QUALITY_CAPTURE_CONFIG="$(jq -c '.capture_profile // null' .source/office-quality/example/office/result.json)"
```

```sh
.source/office-quality/venv/bin/python scripts/office-quality/compare.py \
  .source/office-quality/example/office .source/office-quality/example/betteroffice \
  --out .source/office-quality/example/diff
open .source/office-quality/example/diff/index.html
```

Inputs may be PDFs or directories of `page_0001.png`, etc. The comparator checks available source hashes, render status, page counts, and DPI, then writes `score.json` and an HTML image diff. SSIM uses grayscale pixels at 150 DPI; missing or extra pages reduce the score. Dimensions must match unless `--resize` is explicitly requested. The benchmark does not resize or align captures.

## Public corpus

The `betteroffice-corpus` R2 bucket is served at **https://corpus.betteroffice.dev**. Each sample has its own folder:

```text
<sample>/
  source.<format>
  metadata.json
  reference/page_0001.png
  reference/page_0002.png
  ...
```

Browse the [corpus index](https://corpus.betteroffice.dev/) for current samples, source downloads, Office references, provenance, and third-party notices. Collection manifests record the exact sample IDs selected for each benchmark.

`bo-corpus-2` uses Arial throughout, resolved by the renderer to the registry's metric-compatible Liberation Sans. Its Word reference was regenerated from that converted source; the pair uses different font files, and the metadata records the font normalization and previous reference.

Collection manifests live at `collections/<id>.json` with `schema_version: 1`, the collection `id`, and a `samples` array of folder names. Sample metadata links the source and PNGs and records capture provenance, hashes, comparisons, and licensing. Public sample files, references, and license notices live in the bucket; original private inputs stay local. Publishing additional authorized samples requires authenticated Wrangler:

```sh
bunx wrangler r2 object put betteroffice-corpus/<key> --file <local-file> --remote
```
