#!/usr/bin/env bash
set -euo pipefail
script_dir=$(cd "$(dirname "$0")" && pwd)
test_dir=$(mktemp -d "${TMPDIR:-/tmp}/betteroffice-desktop-tests.XXXXXX")
trap 'rm -rf "$test_dir"' EXIT
swiftc -parse-as-library \
    "$script_dir/DesktopServer.swift" "$script_dir/tests/DesktopServerTests.swift" \
    -o "$test_dir/server-tests"
"$test_dir/server-tests"
swiftc -D DEBUG -parse-as-library -target "$(uname -m)-apple-macosx13.3" \
    "$script_dir/DesktopServer.swift" "$script_dir/Desktop.swift" \
    -o "$test_dir/desktop"
if [[ ${1:-} == --ui ]]; then
    python3 "$script_dir/tests/editor-smoke.py" "$test_dir/desktop"
fi
