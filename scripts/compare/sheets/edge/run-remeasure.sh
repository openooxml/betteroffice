#!/usr/bin/env bash
# Formula fixtures on the pinned Univer presets vs BO: native recalc, Univer forced calc, paired browser captures.
set -euo pipefail
HERE=scripts/compare/sheets
OUT=$HERE/out/${EDGE_RUN:-remeasure-xlsx-formulas}
PROBES=isoweeknum,isoweeknum-20181231,isoweeknum-20191230,isoweeknum-20201231,isoweeknum-20210111,isoweeknum-20210118,isoweeknum-20210201,isoweeknum-20210607,isoweeknum-20211227,isoweeknum-20230102,mround-time,mround-time-0126,mround-time-0510,mround-time-0751,mround-time-0900,rankref-eq,rankref-legacy,rankref-control
mkdir -p "$OUT"
date -u +%FT%TZ > "$OUT/date.txt"
{ hostname; nproc; } > "$OUT/machine.txt"
bun run build:xlsx > "$OUT/build.log" 2>&1
bun node_modules/vite/bin/vite.js build --config "$HERE/bo/vite.config.mjs" > "$OUT/site-bo.log" 2>&1
( cd "$HERE/site" && npm install --no-audit --no-fund && node node_modules/vite/bin/vite.js build ) > "$OUT/site-univer.log" 2>&1
cp "$HERE/site/package-lock.json" "$OUT/package-lock.json"
( cd "$HERE/site" && npm ls --all --json ) > "$OUT/npm-ls.json" 2>&1 || true
node "$HERE/edge/check-packages.mjs" "$OUT" > "$OUT/import-api-scan.txt"
rm -rf "$OUT/work"
python3 "$HERE/edge/formula_probe.py" "$OUT/work" "$PROBES" > "$OUT/generator-sha256.txt"
node scripts/office-quality/build-format.mjs xlsx "$PWD" /tmp/remeasure-native > "$OUT/native.log" 2>&1
cp /tmp/remeasure-native/build.json "$OUT/native-build.json"
python3 "$HERE/acc/bo_calc.py" /tmp/remeasure-native/xlsx-native-bench "$OUT/work" 2 > "$OUT/bo-calc.log" 2>&1
CALC_WORKERS=2 node "$HERE/acc/calc.mjs" "$OUT/work" > "$OUT/univer-calc.log" 2>&1
python3 "$HERE/edge/audit.py" "$OUT/work" "$OUT" > "$OUT/audit.log"
CAPTURE_PROBES=$PROBES node "$HERE/edge/capture-probes.mjs" "$OUT" > "$OUT/capture.log" 2>&1
cat "$OUT/audit.log"
python3 - "$OUT" <<'PY'
import json, sys
from pathlib import Path
out = Path(sys.argv[1])
for r in json.loads((out / 'capture.json').read_text())['records']:
    print(r['id'], r['arm'], json.dumps(r.get('value')), r.get('error') or '', '|'.join(r.get('errors') or [])[:160])
PY
echo REMEASURE-FORMULAS-DONE
