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

The runner compares the latest npm releases with the checked-out source using pinned CDN fonts. Commit source changes first and choose an empty output directory. The default [`office-quality` collection](https://corpus.betteroffice.dev/collections/office-quality.json) selects the three demos and `bo-corpus-1`, a redacted industry document. The [`docx` collection](https://corpus.betteroffice.dev/collections/docx.json) contains exactly `betteroffice-demo` and `bo-corpus-1`. Set `QUALITY_COLLECTION` to another collection or `QUALITY_SAMPLES='["betteroffice-demo"]'` to select a subset, overriding the collection. Runs support up to 100 samples.

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

Capture BetterOffice using the [DOCX instructions](../docx-quality/README.md). For PPTX/XLSX, set `QUALITY_FORMAT=pptx` or `xlsx` on the shared server and use the matching `?format=` in the capture URL. Document bytes stay local; external browser requests are limited to pinned font files.

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

| Sample metadata | Format | Reference pages |
| --- | --- | ---: |
| [betteroffice-demo](https://corpus.betteroffice.dev/betteroffice-demo/metadata.json) | DOCX | 2 |
| [bo-corpus-1](https://corpus.betteroffice.dev/bo-corpus-1/metadata.json) | DOCX | 17 |
| [betteroffice-slides](https://corpus.betteroffice.dev/betteroffice-slides/metadata.json) | PPTX | 3 |
| [betteroffice-workbook](https://corpus.betteroffice.dev/betteroffice-workbook/metadata.json) | XLSX | 9 |

Collection manifests live at `collections/<id>.json` with `schema_version: 1`, the collection `id`, and a `samples` array of folder names. Sample metadata links the source and PNGs and records capture provenance, hashes, comparisons, and licensing. The DOCX corpus contains only the demo and the authorized redacted sample; private inputs and captures stay local. Publishing additional authorized samples requires authenticated Wrangler:

```sh
bunx wrangler r2 object put betteroffice-corpus/<key> --file <local-file> --remote
```
