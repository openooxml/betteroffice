#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 2 ]]; then
  echo 'Usage: run.sh <agents-cli-command> <results-dir> [model]' >&2
  echo '       run.sh --arm generic|scripted-lib --python <python> <results-dir> [model]' >&2
  echo '       run.sh --arm scripted-bo <agents-cli-command> <results-dir>' >&2
  echo 'Either form takes --formats <list>, a comma-separated subset of xlsx,pptx,docx (default: all).' >&2
  echo 'Options: --suite v2|v3 --runs N --jobs K --tasks <ids> --cross-model --resume --agent codex|muse' >&2
  exit 2
fi

suite_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec python3 "$suite_dir/run.py" "$@" </dev/null
