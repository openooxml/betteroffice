#!/usr/bin/env bash
set -euo pipefail

PINNED_COMMIT=c1d3718e2a0eb25c2567a694ee017061fdaa5f51

if [[ ! "$PINNED_COMMIT" =~ ^[a-fA-F0-9]{40}$ ]]; then
  echo 'Fill PINNED_COMMIT with the verified notofonts.github.io commit first.' >&2
  exit 1
fi
if [[ $# -ne 1 ]]; then
  echo 'Usage: fetch-intl.sh <output-directory>' >&2
  exit 1
fi

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
checksums="$script_dir/intl-sources.sha256"
sources=(
  NotoSansArmenian-Bold.ttf
  NotoSansArmenian-Regular.ttf
  NotoSansBengali-Bold.ttf
  NotoSansBengali-Regular.ttf
  NotoSansCherokee-Bold.ttf
  NotoSansCherokee-Regular.ttf
  NotoSansDevanagari-Bold.ttf
  NotoSansDevanagari-Regular.ttf
  NotoSansEthiopic-Bold.ttf
  NotoSansEthiopic-Regular.ttf
  NotoSansGeorgian-Bold.ttf
  NotoSansGeorgian-Regular.ttf
  NotoSansGujarati-Bold.ttf
  NotoSansGujarati-Regular.ttf
  NotoSansGurmukhi-Bold.ttf
  NotoSansGurmukhi-Regular.ttf
  NotoSansKannada-Bold.ttf
  NotoSansKannada-Regular.ttf
  NotoSansKhmer-Bold.ttf
  NotoSansKhmer-Regular.ttf
  NotoSansLao-Bold.ttf
  NotoSansLao-Regular.ttf
  NotoSansLaoLooped-Bold.ttf
  NotoSansLaoLooped-Regular.ttf
  NotoSansMalayalam-Bold.ttf
  NotoSansMalayalam-Regular.ttf
  NotoSansMongolian-Regular.ttf
  NotoSansMyanmar-Bold.ttf
  NotoSansMyanmar-Regular.ttf
  NotoSansOriya-Bold.ttf
  NotoSansOriya-Regular.ttf
  NotoSansSinhala-Bold.ttf
  NotoSansSinhala-Regular.ttf
  NotoSansSyriac-Regular.ttf
  NotoSansTamil-Bold.ttf
  NotoSansTamil-Regular.ttf
  NotoSansTelugu-Bold.ttf
  NotoSansTelugu-Regular.ttf
  NotoSansThaana-Bold.ttf
  NotoSansThaana-Regular.ttf
  NotoSansThai-Bold.ttf
  NotoSansThai-Regular.ttf
  NotoSansTifinagh-Regular.ttf
  NotoSerifTibetan-Bold.ttf
  NotoSerifTibetan-Regular.ttf
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

for file in "${sources[@]}"; do
  case "$seen" in
    *" $file "*) ;;
    *) echo "Missing SHA-256 entry: $file" >&2; exit 1 ;;
  esac
  family="${file%-*}"
  url="https://raw.githubusercontent.com/notofonts/notofonts.github.io/$PINNED_COMMIT/fonts/$family/hinted/ttf/$file"
  if ! curl --fail --location --retry 3 --proto '=https' --tlsv1.2 "$url" --output "$work_dir/$file"; then
    echo "Failed to download hinted font: $file" >&2
    exit 1
  fi
done

(
  cd "$work_dir"
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum --check "$checksums"
  else
    shasum -a 256 --check "$checksums"
  fi
)
mkdir -p "$1"
for file in "${sources[@]}"; do
  cp "$work_dir/$file" "$1/$file"
done
