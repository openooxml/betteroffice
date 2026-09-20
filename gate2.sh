#!/bin/bash
# usage: gate2.sh <docx|xlsx|pptx|opc> [--wasm] [--bindings]
set -o pipefail
export PATH="$HOME/.cargo/bin:$PATH"
export CARGO_TARGET_DIR=/private/tmp/bo-perf-land-20260920/target
cd /private/tmp/bo-perf-land-20260920
FMT=$1; shift; fail=0
case $FMT in
  docx) CRATES="-p betteroffice-docx-parse -p betteroffice-docx-edit -p betteroffice-docx-layout -p betteroffice-docx-raster -p betteroffice-docx";;
  xlsx) CRATES="-p betteroffice-xlsx-parse -p betteroffice-xlsx-calc -p betteroffice-xlsx-ops -p betteroffice-xlsx-render -p betteroffice-xlsx-raster -p betteroffice-xlsx";;
  pptx) CRATES="-p betteroffice-pptx-parse -p betteroffice-pptx-edit -p betteroffice-pptx-render -p betteroffice-pptx-raster -p betteroffice-pptx";;
  opc)  CRATES="-p betteroffice-ooxml-opc -p betteroffice-docx -p betteroffice-xlsx -p betteroffice-pptx -p betteroffice-vsdx";;
esac
echo "### fmt";   cargo fmt --check 2>&1 | head -5 || fail=1
echo "### clippy"; cargo clippy --all-targets $CRATES -- -D warnings > /tmp/g-clippy.log 2>&1 || { fail=1; grep -E "^error" /tmp/g-clippy.log | head -10; }
echo "### test";  cargo test $CRATES > /tmp/g-test.log 2>&1 || fail=1
grep -E "^test result" /tmp/g-test.log | awk -F'[ ;]' '{p+=$4; f+=$6} END{print "passed="p" failed="f}'
grep -E "panicked at|^test result: FAILED|^error\[|^error:" /tmp/g-test.log | head -10
for a in "$@"; do
  if [ "$a" = "--wasm" ]; then
    echo "### wasm+drift"; bun run build:$FMT-wasm > /tmp/g-wasm.log 2>&1 || fail=1; tail -1 /tmp/g-wasm.log
    git status --porcelain packages/$FMT/src/wasm/generated/; echo "(empty=no drift)"
    echo "### bun"; bun test packages/$FMT packages/$FMT-react > /tmp/g-bun.log 2>&1 || fail=1; grep -E "^ [0-9]+ (pass|fail)" /tmp/g-bun.log
    echo "### tsc"; bun run --filter "./packages/$FMT*" typecheck > /tmp/g-tsc.log 2>&1 || fail=1; grep -cE "Exited with code 0" /tmp/g-tsc.log | sed 's/^/ok-packages=/'
  fi
  if [ "$a" = "--bindings" ]; then
    echo "### bindings (separate workspace, --locked)"
    (cd bindings && cargo check --locked --all-targets > /tmp/g-bind.log 2>&1) || { fail=1; grep -E "^error" /tmp/g-bind.log | head -8; }
    echo "bindings-exit=$?"
  fi
done
echo "### GATE $([ $fail -eq 0 ] && echo GREEN || echo RED)"
