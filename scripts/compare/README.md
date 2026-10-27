# Editor comparisons

Reproduces the measurements in the comparison posts on the BetterOffice blog. Run on Linux from the repository root after `bun install --frozen-lockfile`, with Node.js 24, Python 3, the Rust/Wasm toolchain and Playwright Chromium installed where the scripts look for it (`PLAYWRIGHT_BROWSERS_PATH=/opt/ms-playwright bunx playwright install chromium`). The competitor pages are standalone npm projects outside the Bun workspace, installed from the exact npm lockfiles used for the published runs. Results land in `scripts/compare/out/<run>/` and `scripts/compare/sheets/out/<run>/`.

## DOCX: main thread during open (docx-editor.dev)

```sh
python3 -I scripts/compare/build-public-composite.py
CMP_RUN=base CMP_DOCS=public-800pp CMP_N=6 CMP_BURN=1 CMP_OPEN_CAP=600000 bash scripts/compare/run-open.sh
CMP_RUN=fonts CMP_DOCS=public-800pp CMP_N=6 CMP_BURN=1 CMP_OPEN_CAP=600000 \
  CMP_REUSE_RUNTIME=1 CMP_RUNTIME_ARCHIVE=scripts/compare/out/base/runtime-dist.tgz \
  CMP_RUNTIME_LOCK=scripts/compare/out/base/package-lock.json \
  CMP_PREPARE_FONTS=1 CMP_FONT_PRESET=all-defaults-and-pinned-catalog VITE_CMP_EXTENDED_FONTS=1 \
  bash scripts/compare/run-open.sh
```

`build-public-composite.py` downloads the 15 public documents listed with their licences and attribution in [`public-composite.json`](public-composite.json), checks each sha256, and composes them with `compose.py` into `docs/public-800pp.docx` plus its attribution list, `docs/public-800pp.PROVENANCE.txt`. It fails unless the result matches the measured composite byte for byte. The second run is the font control: docx-editor.dev with all default faces plus its pinned font catalog. `open.json` holds every session, including the longest main-thread task and the laid-out page count.

## XLSX: scrolling and cell edits (Univer)

```sh
CMP_RUN=scale CMP_INSTALL=install CMP_WBS=w300k CMP_N=6 CMP_BURN=1 CMP_EDITS=4 bash scripts/compare/sheets/run.sh
```

[`sheets/gen.mjs`](sheets/gen.mjs) writes the synthetic 30,000-row workbook (`bun scripts/compare/sheets/gen.mjs <dir> w300k`) as `.xlsx` for BetterOffice and as an IWorkbookData snapshot for Univer, and prints both sha256 values. `results.json` and `summary.md` hold the frame rates and edit timings.

## XLSX: formulas (Univer)

```sh
bash scripts/compare/sheets/edge/run-remeasure.sh
```

[`sheets/edge/formula_probe.py`](sheets/edge/formula_probe.py) writes one-formula workbooks with typed inputs and no cached values: ISO week (`isoweeknum*`), ranking (`rankref-*`) and time rounding (`mround-time*`). The script recalculates them with the native BetterOffice build and with Univer, then opens each in both editors in Chromium (`capture.json`, `shots/`).
