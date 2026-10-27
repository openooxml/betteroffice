#!/usr/bin/env bash
# VM only. Builds bo-xlsx (workspace packages) and univer-sheets (own lockfile), generates the workbooks, runs bench.mjs.
# Env: CMP_RUN (output name under scripts/compare/sheets/out), CMP_INSTALL=install|ci, CMP_WBS, plus the CMP_* bench knobs.
set -uo pipefail
: "${CMP_RUN:?}"
HERE=$(cd "$(dirname "$0")" && pwd)
ROOT=$(cd "$HERE/../../.." && pwd)
OUT="$HERE/out/$CMP_RUN"
mkdir -p "$OUT"
( while sleep 30; do echo .; done ) & HB=$!
trap 'kill $HB 2>/dev/null' EXIT
echo "host $(hostname) nproc $(nproc) load $(cut -d' ' -f1-3 /proc/loadavg) start $(date -u +%FT%TZ)" | tee "$OUT/machine.txt"
cd "$ROOT"
rm -rf "$HERE/dist"
bun run build:xlsx > "$OUT/build.log" 2>&1 || { echo BUILD-FAILED; tail -30 "$OUT/build.log"; exit 1; }
bun node_modules/vite/bin/vite.js build --config scripts/compare/sheets/bo/vite.config.mjs > "$OUT/site-bo.log" 2>&1 || { echo BO-SITE-FAILED; tail -30 "$OUT/site-bo.log"; exit 1; }
cd "$HERE/site"
npm "${CMP_INSTALL:-ci}" --no-audit --no-fund > "$OUT/npm.log" 2>&1 || { echo NPM-FAILED; tail -30 "$OUT/npm.log"; exit 1; }
cp package-lock.json "$OUT/package-lock.json"
if [ "${CMP_DTS:-0}" = 1 ]; then
  grep -rhoE "(scrollToCell|getActiveRange|setValue|getValue|activate|getCellRect|IEditorBridgeService|DOCS_NORMAL_EDITOR_UNIT_ID_KEY|initialFormulaComputing|CalculationMode)[^;{]{0,80}" \
    node_modules/@univerjs/*/lib/types/facade/*.d.ts node_modules/@univerjs/presets/lib/types/*.d.ts node_modules/@univerjs/preset-sheets-core/lib/types/*.d.ts 2>/dev/null | sort -u | head -60 > "$OUT/univer-dts.txt"
  ls node_modules/@univerjs/preset-sheets-core/lib node_modules/@univerjs/preset-sheets-core/lib/locales 2>/dev/null | head -30 >> "$OUT/univer-dts.txt"
fi
node node_modules/vite/bin/vite.js build > "$OUT/site-univer.log" 2>&1 || { echo UNIVER-SITE-FAILED; tail -40 "$OUT/site-univer.log"; exit 1; }
ls "$HERE/dist" | tee "$OUT/dist.txt"
cd "$HERE"
bun gen.mjs gen "${CMP_WBS:-w100k}" | tee "$OUT/gen.log" || { echo GEN-FAILED; exit 1; }
echo "build done $(date -u +%FT%TZ) load $(cut -d' ' -f1-3 /proc/loadavg)" | tee -a "$OUT/machine.txt"
[ "${CMP_BUILD_ONLY:-0}" = 1 ] && exit 0
CMP_OUT="out/$CMP_RUN" bun bench.mjs 2>&1 | grep --line-buffered -v '^\.$' | tee "$OUT/bench.log"
rc=${PIPESTATUS[0]}
echo "end $(date -u +%FT%TZ) load $(cut -d' ' -f1-3 /proc/loadavg)" | tee -a "$OUT/machine.txt"
exit $rc
