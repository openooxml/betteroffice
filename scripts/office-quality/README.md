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

The runner compares the latest npm releases with the checked-out source using pinned CDN fonts. Commit source changes first and choose an empty output directory. The default [`office-quality` collection](https://corpus.betteroffice.dev/collections/office-quality.json) selects the current samples across all three formats. The [`docx` collection](https://corpus.betteroffice.dev/collections/docx.json) selects Word documents. Set `QUALITY_COLLECTION` to another collection or `QUALITY_SAMPLES='["betteroffice-demo"]'` to select a subset, overriding the collection. Runs support up to 100 samples.

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

After the [workflow](../../.github/workflows/visual-fidelity.yml) lands on `main`, use **Actions → Visual fidelity → Run workflow**, or:

```sh
gh workflow run visual-fidelity.yml --ref main -f branch=main
```

Keep `--ref main`; set `branch` to the repository branch to measure. The optional `collection` and `samples` inputs select the corpus as above. The action uses existing bot credentials to update the [README scores](../../README.md#visual-fidelity) as `openooxml-bot[bot]`. Unchanged results create no commit; a changed branch head requires a rerun. Runs are manual only. CI retains score JSON and generated Markdown; documents and page images are excluded from uploaded artifacts.

The measurement job allows 90 minutes for setup, builds, and both comparison channels. Each browser capture retains its default 600-second deadline. The longer job budget accommodates larger collections; incomplete captures remain recorded failures.

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
