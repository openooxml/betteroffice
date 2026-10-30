#!/usr/bin/env bash
set -euo pipefail

SANS_COMMIT=523d033d6cb47f4a80c58a35753646f5c3608a78
SERIF_COMMIT=9b0f1436e455d902de067a2501422e5dc71ad16b

if [[ ! "$SANS_COMMIT" =~ ^[a-fA-F0-9]{40}$ || ! "$SERIF_COMMIT" =~ ^[a-fA-F0-9]{40}$ ]]; then
  echo 'Fill SANS_COMMIT and SERIF_COMMIT with verified noto-cjk commits first.' >&2
  exit 1
fi
if [[ $# -lt 1 || $# -gt 2 ]]; then
  echo 'Usage: fetch-cjk.sh <output-directory> [local-source-directory]' >&2
  exit 1
fi

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
checksums="$script_dir/cjk-sources.sha256"
sources=(
  NotoSansJP-Regular.otf
  NotoSansKR-Regular.otf
  NotoSansSC-Regular.otf
  NotoSansTC-Regular.otf
  NotoSerifSC-Regular.otf
)
paths=(
  Sans/SubsetOTF/JP/NotoSansJP-Regular.otf
  Sans/SubsetOTF/KR/NotoSansKR-Regular.otf
  Sans/SubsetOTF/SC/NotoSansSC-Regular.otf
  Sans/SubsetOTF/TC/NotoSansTC-Regular.otf
  Serif/SubsetOTF/SC/NotoSerifSC-Regular.otf
)
work_dir="$(mktemp -d)"
trap 'rm -rf "$work_dir"' EXIT

seen=' '
while read -r digest file extra; do
  if [[ ! "$digest" =~ ^[a-fA-F0-9]{64}$ || -n "${extra:-}" ]]; then
    echo "Invalid or unfilled SHA-256 entry: $digest $file" >&2
    exit 1
  fi
  case " ${sources[*]} " in
    *" $file "*) ;;
    *) echo "Unexpected source: $file" >&2; exit 1 ;;
  esac
  case "$seen" in
    *" $file "*) echo "Duplicate source: $file" >&2; exit 1 ;;
  esac
  seen="$seen$file "
done < "$checksums"

for index in "${!sources[@]}"; do
  file="${sources[$index]}"
  case "$seen" in
    *" $file "*) ;;
    *) echo "Missing SHA-256 entry: $file" >&2; exit 1 ;;
  esac
  if [[ $# -eq 2 ]]; then
    if ! cp "$2/$file" "$work_dir/$file"; then
      echo "Failed to copy local CJK font: $file" >&2
      exit 1
    fi
  else
    path="${paths[$index]}"
    case "$path" in
      Sans/*) commit="$SANS_COMMIT" ;;
      Serif/*) commit="$SERIF_COMMIT" ;;
      *) echo "Unexpected source path: $path" >&2; exit 1 ;;
    esac
    url="https://raw.githubusercontent.com/notofonts/noto-cjk/$commit/$path"
    if ! curl --fail --location --retry 3 --proto '=https' --tlsv1.2 "$url" --output "$work_dir/$file"; then
      echo "Failed to download CJK font: $file" >&2
      exit 1
    fi
  fi
done

if ! (
  cd "$work_dir"
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum --check "$checksums"
  else
    shasum -a 256 --check "$checksums"
  fi
); then
  echo 'CJK source SHA-256 verification failed.' >&2
  exit 1
fi
mkdir -p "$1"
for file in "${sources[@]}"; do
  cp "$work_dir/$file" "$1/$file"
done
