#!/usr/bin/env bash
# Builds the BetterOffice page (workspace packages) and the docx-editor.dev page (own lockfile) into scripts/compare/dist.
# Env: CMP_RUN (output name under scripts/compare/out), CMP_INSTALL=install|ci.
set -uo pipefail
: "${CMP_RUN:?}"
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
HERE="$ROOT/scripts/compare"
OUT="$HERE/out/$CMP_RUN"
mkdir -p "$OUT"
echo "host $(hostname) nproc $(nproc) load $(cut -d' ' -f1-3 /proc/loadavg) start $(date -u +%FT%TZ)" | tee "$OUT/machine.txt"
cd "$ROOT"
rm -rf "$HERE/dist"
{ bun run build:fonts && bun run build:docx; } > "$OUT/build.log" 2>&1 || { echo BUILD-FAILED; tail -30 "$OUT/build.log"; exit 1; }
bun node_modules/vite/bin/vite.js build --config scripts/compare/bo/vite.config.mjs > "$OUT/site-bo.log" 2>&1 || { echo BO-SITE-FAILED; tail -30 "$OUT/site-bo.log"; exit 1; }
cd "$HERE/site"
npm "${CMP_INSTALL:-ci}" --no-audit --no-fund > "$OUT/npm.log" 2>&1 || { echo NPM-FAILED; tail -30 "$OUT/npm.log"; exit 1; }
cp package-lock.json "$OUT/package-lock.json"
node node_modules/vite/bin/vite.js build > "$OUT/site-cmp.log" 2>&1 || { echo CMP-SITE-FAILED; tail -40 "$OUT/site-cmp.log"; exit 1; }
ls "$HERE/dist" | tee "$OUT/dist.txt"
echo "build done $(date -u +%FT%TZ) load $(cut -d' ' -f1-3 /proc/loadavg)" | tee -a "$OUT/machine.txt"
