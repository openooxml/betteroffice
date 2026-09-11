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
trap 'rm -rf "$staging_dir"' EXIT

mkdir -p "$(dirname "$dmg_path")"
for app in "${apps[@]}"; do
    ditto "$app" "$staging_dir/${app##*/}"
done
ln -s /Applications "$staging_dir/Applications"
hdiutil create \
    -volname BetterOffice \
    -srcfolder "$staging_dir" \
    -format UDZO \
    -ov \
    "$dmg_path"
