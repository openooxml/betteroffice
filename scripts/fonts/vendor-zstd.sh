#!/usr/bin/env bash
set -euo pipefail

ZSTD_TARBALL_SHA256=eb33e51f49a15e023950cd7825ca74a4a2b43db8354825ac24fc1b7ee09e6fa3
ZSTD_VERSION=1.5.7
ZSTD_URL="https://github.com/facebook/zstd/releases/download/v${ZSTD_VERSION}/zstd-${ZSTD_VERSION}.tar.gz"

if [[ ! "$ZSTD_TARBALL_SHA256" =~ ^[a-fA-F0-9]{64}$ ]]; then
  echo 'Fill ZSTD_TARBALL_SHA256 with the verified release tarball digest first.' >&2
  exit 1
fi

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
work_dir="$(mktemp -d)"
trap 'rm -rf "$work_dir"' EXIT
archive="$work_dir/zstd.tar.gz"
curl --fail --location --proto '=https' --tlsv1.2 "$ZSTD_URL" --output "$archive"
if command -v sha256sum >/dev/null 2>&1; then
  printf '%s  %s\n' "$ZSTD_TARBALL_SHA256" "$archive" | sha256sum --check --status
else
  printf '%s  %s\n' "$ZSTD_TARBALL_SHA256" "$archive" | shasum -a 256 --check --status
fi
tar -xzf "$archive" -C "$work_dir"
source_dir="$work_dir/zstd-${ZSTD_VERSION}"
(
  cd "$source_dir/build/single_file_libs"
  sh create_single_file_decoder.sh
)
destination="$repo_root/crates/betteroffice-fonts/vendor/zstd"
mkdir -p "$destination"
cp "$source_dir/build/single_file_libs/zstddeclib.c" "$destination/zstddeclib.c"
cp "$source_dir/LICENSE" "$destination/LICENSE"
cp "$source_dir/LICENSE" "$repo_root/packages/fonts/LICENSES/BSD-3-Clause-Zstandard.txt"
