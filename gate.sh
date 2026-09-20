#!/bin/bash
# usage: gate.sh <docx|xlsx|pptx|opc> [--wasm]
set -o pipefail
export PATH="$HOME/.cargo/bin:$PATH"
export CARGO_TARGET_DIR=/private/tmp/bo-perf-land-20260920/target
cd /private/tmp/bo-perf-land-20260920
FMT=$1; WASM=$2; fail=0
case $FMT in
  docx) CRATES="-p betteroffice-docx-parse -p betteroffice-docx-edit -p betteroffice-docx-layout -p betteroffice-docx-raster -p betteroffice-docx";;
  xlsx) CRATES="-p betteroffice-xlsx-parse -p betteroffice-xlsx-calc -p betteroffice-xlsx-ops -p betteroffice-xlsx-render -p betteroffice-xlsx-raster -p betteroffice-xlsx";;
  pptx) CRATES="-p betteroffice-pptx-parse -p betteroffice-pptx-edit -p betteroffice-pptx-render -p betteroffice-pptx-raster -p betteroffice-pptx";;
  opc)  CRATES="-p betteroffice-ooxml-opc -p betteroffice-docx -p betteroffice-xlsx -p betteroffice-pptx -p betteroffice-vsdx";;
esac
echo "### fmt"; cargo fmt --check || fail=1
echo "### clippy"; cargo clippy --all-targets $CRATES -- -D warnings 2>&1 | grep -E "^error|^warning" | head -15; [ ${PIPESTATUS[0]} -eq 0 ] || fail=1
echo "### test"; cargo test $CRATES 2>&1 | grep -E "^test result: FAILED|^error|panicked at|failures:" | head -15
cargo test $CRATES 2>&1 | grep -c "^test result: ok" | sed 's/^/ok-suites=/'
cargo test $CRATES 2>&1 | grep -E "^test result" | awk -F'[ ;]' '{p+=$4} END{print "tests-passed="p}'
cargo test $CRATES >/dev/null 2>&1 || fail=1
if [ "$WASM" = "--wasm" ]; then
  echo "### wasm + drift"; bun run build:$FMT-wasm 2>&1 | tail -1
  git status --porcelain packages/$FMT/src/wasm/generated/; echo "(empty=no drift)"
  echo "### bun"; bun test packages/$FMT packages/$FMT-react 2>&1 | tail -4 || fail=1
  echo "### typecheck"; bun run --filter "./packages/$FMT*" typecheck 2>&1 | grep -E "Exited with code|error" | head -6 || fail=1
fi
echo "### GATE $([ $fail -eq 0 ] && echo GREEN || echo RED)"
