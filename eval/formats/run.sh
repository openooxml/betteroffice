#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 2 ]]; then
  echo 'Usage: run.sh <agents-cli-command> <results-dir> [model]' >&2
  echo '       run.sh --arm generic --python <python> <results-dir> [model]' >&2
  exit 2
fi

suite_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec python3 "$suite_dir/run.py" "$@" </dev/null
