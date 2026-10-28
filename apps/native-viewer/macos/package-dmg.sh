#!/usr/bin/env bash

set -euo pipefail

if [[ $# -ne 4 ]]; then
    echo "usage: $0 DOCS_APP SHEETS_APP SLIDES_APP DMG_PATH" >&2
    exit 2
fi

apps=("$1" "$2" "$3")
expected=("BetterOffice Docs.app" "BetterOffice Sheets.app" "BetterOffice Slides.app")
dmg_path=$4

for index in "${!apps[@]}"; do
    if [[ ! -d ${apps[$index]} || ${apps[$index]##*/} != "${expected[$index]}" ]]; then
        echo "expected ${expected[$index]}" >&2
        exit 2
    fi
done

staging_dir=$(mktemp -d "${TMPDIR:-/tmp}/betteroffice-dmg.XXXXXX")
errors="$staging_dir.err"
trap 'rm -rf "$staging_dir" "$errors"' EXIT

mkdir -p "$(dirname "$dmg_path")"
for app in "${apps[@]}"; do
    ditto "$app" "$staging_dir/${app##*/}"
done
ln -s /Applications "$staging_dir/Applications"
for attempt in 1 2 3 4 5; do
    if [[ $attempt -gt 1 && -d /Volumes/BetterOffice ]]; then
        hdiutil detach /Volumes/BetterOffice -force >/dev/null 2>&1 || true
    fi
    status=0
    hdiutil create \
        -volname BetterOffice \
        -srcfolder "$staging_dir" \
        -format UDZO \
        -ov \
        "$dmg_path" 2>"$errors" || status=$?
    cat "$errors" >&2
    if [[ $status -eq 0 ]]; then
        exit 0
    fi
    if ! grep -q "Resource busy" "$errors" || [[ $attempt -eq 5 ]]; then
        exit 1
    fi
    delay=$((2 ** attempt))
    echo "hdiutil create: resource busy, retrying in ${delay}s (attempt $attempt of 5)" >&2
    sleep "$delay"
done
