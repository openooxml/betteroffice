#!/usr/bin/env bash
set -euo pipefail
: "${CMP_RUN:?}"
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
cd "$ROOT"
OUT="scripts/compare/out/$CMP_RUN"
mkdir -p "$OUT" scripts/compare/gen
sha256sum scripts/compare/bench-open.mjs scripts/compare/bo/bo-docx.tsx scripts/compare/site/dxe-docx.tsx scripts/compare/long_doc.py scripts/compare/fixtures.ts > "$OUT/harness-sha256.txt"
if [ "${CMP_REUSE_RUNTIME:-0}" != 1 ]; then
  CMP_INSTALL=install bash scripts/compare/build.sh
  tar -czf "$OUT/runtime-dist.tgz" -C scripts/compare dist
else
  tar -xzf "$CMP_RUNTIME_ARCHIVE" -C scripts/compare
  cp "$CMP_RUNTIME_LOCK" scripts/compare/site/package-lock.json
  ( cd scripts/compare/site
    npm ci --no-audit --no-fund > "$ROOT/$OUT/npm.log" 2>&1
    if [ "${CMP_PREPARE_FONTS:-0}" = 1 ]; then
      python3 ../font-families.py ../docs/public-800pp.docx > "$ROOT/$OUT/font-families.json"
      bun prepare-fonts.mjs "$ROOT/$OUT/font-families.json" "$ROOT/$OUT/font-config.json" > "$ROOT/$OUT/font-prep.log" 2>&1
    fi
    node node_modules/vite/bin/vite.js build > "$ROOT/$OUT/site-cmp.log" 2>&1
  )
  cp scripts/compare/site/package-lock.json "$OUT/package-lock.json"
fi
for n in 1500 3000; do
  python3 scripts/compare/long_doc.py --pages "$n" --out "scripts/compare/gen/long$n.docx" | tee -a "$OUT/fixtures.txt"
done
cd scripts/compare
CMP_OUT="out/$CMP_RUN" bun bench-open.mjs 2>&1 | tee "out/$CMP_RUN/open.log"
mkdir -p "out/$CMP_RUN/fixtures"
cp gen/long*.docx "out/$CMP_RUN/fixtures/"
[ ! -f gen/syn800.docx ] || cp gen/syn800.docx "out/$CMP_RUN/fixtures/"
