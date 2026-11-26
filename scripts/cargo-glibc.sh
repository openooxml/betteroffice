#!/usr/bin/env bash
set -euo pipefail
floor=$(node --input-type=module -e "import { GLIBC_FLOOR } from '$(cd "$(dirname "$0")" && pwd)/glibc-floor.mjs'; console.log(GLIBC_FLOOR)")
args=()
for arg in "$@"; do
  case "$arg" in
    *-unknown-linux-gnu) arg="$arg.$floor" ;;
  esac
  args+=("$arg")
done
unset CARGO
exec cargo "${args[@]}"
