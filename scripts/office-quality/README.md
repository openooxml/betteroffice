# Local Office visual checks

Small scripts for documents you choose. There is no bundled dataset, downloader, scheduled run, or upload service. Keep inputs, PDFs, captures, and reports under ignored `.source/office-quality/`.

One explicitly published, project-authored demo lives in the public `betteroffice-corpus` R2 bucket; see the [reference links and SSIM results](../../README.md#visual-fidelity). The reference consists of a DOCX, two Word PNG pages, and JSON metadata with UTC capture times, Office version, hashes, license, and scores. Everything else stays local. Public access permits downloads; writes use authenticated Wrangler. There is no anonymous upload endpoint.

## Setup

```sh
python3 -m venv .source/office-quality/venv
.source/office-quality/venv/bin/pip install -r scripts/office-quality/requirements.txt
```

Reference exports require macOS and the corresponding installed desktop app: Microsoft Word, PowerPoint, or Excel. The comparison script also works on other platforms.

Run exports serially. Word and Excel PDF exports were verified on macOS 26.2 with Office 16.112.3. The PowerPoint adapter compiles against the installed dictionary, but this Mac rejected opening files (`-9074`) and creating presentations (`-2710`); its end-to-end export remains unverified pending resolution of that app state. A manually exported local PDF works with the comparator meanwhile.

## Export an Office reference

```sh
.source/office-quality/venv/bin/python scripts/office-quality/reference.py \
  /path/to/example.docx --out .source/office-quality/example/word
```

Use `.pptx` or `.xlsx` to select PowerPoint or Excel. Each invocation exports one document, saves `reference.pdf`, rasterizes every page at 150 DPI, and records the source SHA-256, Office version, macOS version, font names, and page count in `result.json`. Choose a fresh output directory for each revision. The timeout defaults to 120 seconds and can be changed with `--timeout`.

`result.json` also records UTC start/finish timestamps and the PDF's own creation metadata. To publish a reference manually, use `bunx wrangler r2 object put betteroffice-corpus/<key> --file <local-file> --remote` for the source, PNGs, and provenance JSON. Publishing documents requires redistribution rights for their contents and embedded assets; keep personal or confidential documents out of this public bucket.

Excel references follow the workbook's print settings. Match its print areas, scaling, margins, page breaks, and sheet selection when producing the candidate. A spreadsheet editor viewport is not the same surface as an Excel printed page. The current automatic BetterOffice capture below supports DOCX; PPTX/XLSX candidates can be supplied as PDFs or PNG pages from their renderer/exporter.

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

## Capture BetterOffice DOCX

Follow [the DOCX capture instructions](../docx-quality/README.md). The browser receives the document bytes locally, and external requests are restricted to pinned font files. No document body is sent to a CDN. The demo and harness disable automatic Google Fonts lookups; the existing SDK default remains unchanged.

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
