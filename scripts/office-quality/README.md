# Local Office visual checks

Small scripts for documents you choose. Keep inputs, PDFs, captures, and reports under ignored `.source/office-quality/`. The manual CI runner downloads explicitly selected public references; there is no scheduled run or upload service.

Three explicitly published, project-authored demos live in the public `betteroffice-corpus` R2 bucket: DOCX, PPTX, and XLSX. See the [reference links and SSIM results](../../README.md#visual-fidelity). Each completed sample contains the original source, Office PNG pages, and JSON metadata with UTC capture times, Office version, hashes, and license. Everything else stays local. Public access permits downloads; writes use authenticated Wrangler. There is no anonymous upload endpoint.

The canonical origin is **https://corpus.betteroffice.dev**, with one top-level folder per sample. The source files for all three formats are public. DOCX references are complete; PPTX and the XLSX print-profile references are awaiting the local Office dialogs described below. Their intended layout is:

```text
betteroffice-demo/
  source.docx
  reference/page_0001.png
  reference/page_0002.png
  metadata.json
betteroffice-slides/
  source.pptx
  reference/page_0001.png … page_0003.png
  metadata.json
betteroffice-workbook/
  source.xlsx
  reference/page_0001.png … page_0009.png
  metadata.json
```

[Demo source](https://corpus.betteroffice.dev/betteroffice-demo/source.docx), [Word page 1](https://corpus.betteroffice.dev/betteroffice-demo/reference/page_0001.png), [Word page 2](https://corpus.betteroffice.dev/betteroffice-demo/reference/page_0002.png), and [metadata](https://corpus.betteroffice.dev/betteroffice-demo/metadata.json) are public. The [slide source](https://corpus.betteroffice.dev/betteroffice-slides/source.pptx) uses `apps/demo/public/betteroffice-demo.pptx`; the [workbook source](https://corpus.betteroffice.dev/betteroffice-workbook/source.xlsx) uses `apps/demo/public/showcase.xlsx`. The metadata's `format`, `reference`, `reference_pages`, and asset hashes describe the sample independently of Git.

## Manual CI and generated README

After this workflow lands on `main`, run **Actions → Visual fidelity → Run workflow**. Keep the workflow ref on `main`; choose the repository branch to measure and update, including an open PR's branch. The sample input is a JSON array of folder names and defaults to `["betteroffice-demo","betteroffice-slides","betteroffice-workbook"]`.

```sh
gh workflow run visual-fidelity.yml --ref main \
  -f branch=main -f 'samples=["betteroffice-demo","betteroffice-slides","betteroffice-workbook"]'
```

The action builds all three branch renderers, fetches each current npm release, verifies source/Office image hashes from R2, captures both channels with pinned CDN fonts, and computes fresh SSIM. The saved Office references remain unchanged. DOCX captures document pages, PPTX captures complete slides through the core canvas renderer, and XLSX paints the recorded worksheet print ranges. Unmeasured formats show `—`; a selected sample with a missing, invalid, or failed reference stops the run. Until the pending references are published, select `["betteroffice-demo"]` to run the available measurements.

The generated block sits immediately above Contributing. Scores are pinned to the tested source revision; README-only commits are excluded from revision selection. A separate job uses the existing `OPENOOXML_BOT_APP_ID` / `OPENOOXML_BOT_PRIVATE_KEY` secrets to commit **only README.md** as `openooxml-bot[bot]`. It checks that the target branch still matches the measured head and uses a normal fast-forward push. A changed branch requires a rerun; unchanged generated text produces no commit. There is no push, PR, merge, or scheduled trigger.

Only `report.json` and the generated Markdown are retained as CI artifacts. Documents, PDFs, and page images are not uploaded to Actions or committed. Benchmark jobs have read-only repository access and no bot or R2 credentials.

To run the same measurements locally after `bun run build:packages` and the Python setup below:

```sh
QUALITY_PYTHON=.source/office-quality/venv/bin/python \
QUALITY_OUTPUT=.source/office-quality/run \
  node scripts/office-quality/run.mjs
node scripts/office-quality/readme.mjs .source/office-quality/run/report.json
```

Use a fresh output directory and commit source changes first; the runner refuses to label an uncommitted source tree with a commit SHA. Additional selected samples are supplied with `QUALITY_SAMPLES='["sample-folder"]'`. The generator never reuses a score for a different published version or source commit.

## Setup

```sh
python3 -m venv .source/office-quality/venv
.source/office-quality/venv/bin/pip install -r scripts/office-quality/requirements.txt
```

Reference exports require macOS and the corresponding installed desktop app: Microsoft Word, PowerPoint, or Excel. The comparison script also works on other platforms.

Run exports serially. Word and Excel PDF exports were verified on macOS 26.2 with Office 16.112.3. PowerPoint can reject opening files (`-9074`) or creating presentations (`-2710`) while its first-run “Start Using PowerPoint” screen is present. Clear that screen once before running exports. On this Mac, UI control is not granted, so it requires a manual click. An Excel repair dialog from a discarded temporary print-setup rewrite also needs to be dismissed with **No** before the profile export can be verified; the current exporter keeps the source bytes unchanged.

## Export an Office reference

```sh
.source/office-quality/venv/bin/python scripts/office-quality/reference.py \
  /path/to/example.docx --out .source/office-quality/example/word
```

Use `.pptx` or `.xlsx` to select PowerPoint or Excel. Each invocation exports one document, saves `reference.pdf`, rasterizes every page at 150 DPI, and records the source SHA-256, Office version, macOS version, font names, and page count in `result.json`. Choose a fresh output directory for each revision. The timeout defaults to 120 seconds and can be changed with `--timeout`.

`result.json` also records UTC start/finish timestamps and the PDF's own creation metadata. To publish a reference manually, use `bunx wrangler r2 object put betteroffice-corpus/<key> --file <local-file> --remote` for the source, PNGs, and provenance JSON. Publishing documents requires redistribution rights for their contents and embedded assets; keep personal or confidential documents out of this public bucket.

Excel's default export follows the workbook's print settings. For automatic comparisons, pass `--xlsx-profile profile.json` to export recorded rectangular ranges individually: landscape pages, a fixed percentage scale, explicit margins, gridlines, and no row/column headings or headers/footers. Other worksheets are hidden only in the temporary open workbook so Excel exports one selected range at a time; the source bytes remain unchanged. Each range must fit one printed page, or the exporter fails. The profile records the actual Office paper dimensions and each constituent PDF's metadata.

The demo profile uses 75% scale, 18-point margins, Dashboard `A1:N16` (including the chart), Formulas `A1:K14`, and all 301 Data rows in blocks of at most 45 rows, for nine pages. The profile will live in the completed sample metadata as `reference.capture_profile`, not in Git. Candidate capture passes all nine ranges; the Office profile export remains pending the dialog cleanup. Reuse it when retaking the reference:

```sh
.source/office-quality/venv/bin/python scripts/office-quality/reference.py \
  apps/demo/public/showcase.xlsx --xlsx-profile /path/to/profile.json \
  --out .source/office-quality/workbook/excel
```

A profile contains `scale_percent`, `margin_pt`, and `pages`, each with a zero-based `sheet` index and an A1 `range`. The candidate adapter uses these settings directly and requires the complete range to fit; it does not resize or align images based on reference pixels. This measures worksheet range rendering, not the engine's automatic print pagination. Frozen panes are currently rejected by the adapter. Keep the Office paper size consistent when retaking references on another machine.

## macOS permission learning

**Automation and Office file access are separate permissions.** In System Settings → Privacy & Security → Automation, allow the app running the command (for example, Terminal or Codex) to control the relevant Microsoft Office app. Each Office app may ask once. See [Apple's Automation instructions](https://support.apple.com/en-ie/guide/mac-help/mchl108e1718/mac).

Word's repeated **Grant File Access** prompts occurred when AppleScript wrote each new PDF outside Word's sandbox. Moving both input and output into Word's own container allowed two differently named fixtures to export successfully on macOS 26.2 / Word 16.112.3. No Full Disk Access change was needed for those tests.

The exporter stages each job under the corresponding app's local container:

```text
~/Library/Containers/com.microsoft.Word/Data/Documents/BetterOfficeBenchmark/
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/BetterOfficeBenchmark/
~/Library/Containers/com.microsoft.Excel/Data/Documents/BetterOfficeBenchmark/
```

After a successful export, Python copies the PDF into your chosen local output directory and removes the staged files. Failed jobs retain their staging directory for inspection. If macOS asks the script-running app to access another app's data, approve that access when requested. These are normal app-container permissions; the script does not edit permission databases, bookmarks, or system settings.

For external files, Microsoft's supported alternative is [`GrantAccessToMultipleFiles`](https://learn.microsoft.com/en-us/office/vba/office-mac/grantaccesstomultiplefiles): approve the requested paths once; Office retains the grants. Granting Automation alone does not grant arbitrary filesystem access.

The scripts use desktop PDF export. When exporting manually from Word, select **Best for printing** or the local Print → Save as PDF path. Do not select the option labeled **uses Microsoft online service** when keeping documents local. On a timeout, inspect Office's permission or document dialog before retrying; the exporter does not terminate your Office app or close unrelated documents.

## Capture BetterOffice

Follow [the DOCX capture instructions](../docx-quality/README.md), or use the all-format runner above. For a standalone slide/workbook capture, start `scripts/docx-quality/server.mjs` with `QUALITY_FORMAT=pptx` or `xlsx`, pass `?format=pptx` / `?format=xlsx` in the browser-task server URL, and provide the XLSX profile as JSON in `QUALITY_CAPTURE_CONFIG`. The browser receives the document bytes locally, and external requests are restricted to pinned font files. No document body is sent to a CDN. The demo and harness disable automatic Google Fonts lookups; the existing SDK default remains unchanged.

## Compare

```sh
.source/office-quality/venv/bin/python scripts/office-quality/compare.py \
  .source/office-quality/example/word .source/office-quality/example/betteroffice \
  --out .source/office-quality/example/diff
open .source/office-quality/example/diff/index.html
```

Each input can be a PDF or a directory containing `page_0001.png`, `page_0002.png`, etc. When both directories contain source hashes, they must match. Recorded render failures, inconsistent page counts, and different DPI are rejected. The output contains `score.json` and a local HTML report showing the reference, candidate, and amplified pixel difference.

Scores are grayscale SSIM on common pages and page-penalized SSIM (`sum(page scores) / max(page counts)`). Missing or extra pages count against the penalized score. Image sizes must match unless you explicitly pass `--resize` for Lanczos resize-to-match scoring. Record renderer versions and use the same fonts and export settings across revisions.

## Relation to Oxi

We checked the canonical GitLab repository, GitHub mirror/fork, releases, packages, Actions artifacts, and benchmark-directory history. Both main branches were at `e3b653d81394d49d7a5c9b649e595b488eebd299`; the public artifacts contained site/Wasm builds, not the frozen DOCX Word-reference bundle. Oxi's [CI configuration](https://gitlab.com/Ryujiyasu/oxi/-/blob/e3b653d81394d49d7a5c9b649e595b488eebd299/.gitlab-ci.yml) explicitly says Office fidelity gates run locally on Windows. Its [reference renderer](https://github.com/Ryujiyasu/oxi/blob/e3b653d81394d49d7a5c9b649e595b488eebd299/pipeline/word_renderer.py) uses Windows Word COM.

Mac Office references let us measure repeatable local before/after changes. They are a separate oracle from Oxi's Windows Office build and cannot establish a rank on its published leaderboard. Matching that table still requires the original frozen selection and Windows references, or a fresh comparison of all engines against one newly generated reference set.
