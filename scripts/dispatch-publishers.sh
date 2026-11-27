#!/usr/bin/env bash
set -euo pipefail

if [ "$#" -lt 3 ]; then
  echo 'Usage: bash scripts/dispatch-publishers.sh <workflow-file> <watch-minutes> <label, with %s for the binding> [binding ...]' >&2
  exit 2
fi

workflow=$1
minutes=$2
label_format=$3
shift 3

status=0
runs=()
for binding in "$@"; do
  # shellcheck disable=SC2059 # the label format is the caller's
  label=$(printf "$label_format" "$binding")
  if ! bash "$CHECK_PUBLISHER_FILES"; then
    echo "::error::Dispatching the publish of $label failed."
    status=1
    break
  fi
  if run=$(gh api --method POST -H "X-GitHub-Api-Version: 2026-03-10" \
    "repos/$GH_REPO/actions/workflows/$workflow/dispatches" \
    -f ref=main -f "inputs[binding]=$binding" -f "inputs[sha]=$SHA" \
    -F "inputs[dry_run]=false" --jq '.workflow_run_id') && [ -n "$run" ]; then
    runs+=("$run:$label")
  else
    echo "::error::Dispatching the publish of $label failed."
    status=1
  fi
done

if [ -z "${runs[0]:-}" ]; then
  exit "$status"
fi

watchers=()
for entry in "${runs[@]}"; do
  run=${entry%%:*}
  label=${entry#*:}
  (
    timeout "${minutes}m" gh run watch "$run" --exit-status || {
      if [ "$?" -eq 124 ]; then
        gh run cancel "$run" || true
      fi
      echo "::error::$label did not publish: $GITHUB_SERVER_URL/$GH_REPO/actions/runs/$run"
      exit 1
    }
  ) &
  watchers+=("$!")
done
for watcher in "${watchers[@]}"; do
  wait "$watcher" || status=1
done
exit "$status"
