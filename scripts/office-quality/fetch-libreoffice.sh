#!/usr/bin/env bash
set -euo pipefail

source "$(dirname "${BASH_SOURCE[0]}")/libreoffice-pin.sh"

dest=${1:-$LO_ARCHIVE_PATH}
sha=${2:-$LO_SHA256}
if (($# > 2)); then mirrors=("${@:3}"); else mirrors=("${LO_MIRRORS[@]}"); fi
floor=${LO_FETCH_FLOOR:-524288}
floor_seconds=${LO_FETCH_FLOOR_SECONDS:-30}
budget=${LO_FETCH_BUDGET:-3600}
attempts=${LO_FETCH_ATTEMPTS:-3}
partial="$dest.part"
slow=''
failed=0

digest() {
  if command -v sha256sum >/dev/null; then sha256sum "$1"; else shasum -a 256 "$1"; fi | cut -d ' ' -f 1
}

matches() {
  [[ -s "$1" && "$(digest "$1")" == "$sha" ]]
}

size() {
  if [[ -e "$partial" ]]; then wc -c < "$partial" | tr -d ' '; else echo 0; fi
}

remaining() {
  echo $((budget - SECONDS))
}

fetch() {
  local url=$1 out limit
  shift
  offset=$(size)
  status=0
  limit=$(remaining)
  if ((limit < 1)); then limit=1; fi
  out=$(curl --location --fail --silent --show-error --connect-timeout 30 --max-time "$limit" \
    --continue-at - --output "$partial" --write-out '%{http_code} %{speed_download} %{size_download} %{time_connect}' "$@" "$url") || status=$?
  read -r code rate received connected <<< "$out" || true
  rate=${rate%.*}
  code=${code:-000}
  rate=${rate:-0}
  received=${received%.*}
  received=${received:-0}
  connected=${connected:-0}
}

timed_out() {
  if ((received == 0)) && [[ $connected =~ ^0(\.0*)?$ ]]; then
    echo "no data (connect timeout)"
  else
    echo "$1 after $received bytes at an average $rate B/s"
  fi
}

report() {
  echo "mirror $1: $2" >&2
}

try_mirror() {
  local url=$1 attempt
  outcome="no attempt inside the budget"
  for ((attempt = 1; attempt <= attempts; attempt++)); do
    if (($(remaining) <= 0)); then outcome="budget of ${budget}s exhausted"; return 1; fi
    if [[ -n ${floor_args:-} ]]; then
      fetch "$url" --speed-limit "$floor" --speed-time "$floor_seconds"
    else
      fetch "$url" --speed-limit 1 --speed-time 120
    fi
    if matches "$partial"; then outcome="ok at $rate B/s"; return 0; fi
    if ((status == 0)); then
      rm -f "$partial"
      if ((offset > 0)); then echo "retrying $url from byte 0 after a resumed sha256 mismatch" >&2; continue; fi
      outcome="sha256 mismatch, file deleted"
      return 1
    fi
    case $status in
      22)
        if [[ $code == 416 ]]; then rm -f "$partial"; continue; fi
        outcome="http $code"
        return 1
        ;;
      28)
        if (($(remaining) <= 0)); then outcome="budget of ${budget}s exhausted: $(timed_out "timed out"), $(size) bytes in total"; return 1; fi
        if [[ -n $floor_args ]]; then outcome="$(timed_out "below the floor of $floor B/s for ${floor_seconds}s")"; return 2; fi
        ;;
      33) rm -f "$partial" ;;
    esac
    outcome="curl exit $status (http $code) after $received bytes at an average $rate B/s"
    echo "retrying $url after curl exit $status at byte $(size)" >&2
    sleep "${LO_FETCH_PAUSE:-2}"
  done
  return 1
}

mkdir -p "$(dirname "$dest")"
if matches "$dest"; then
  echo "LibreOffice archive already present: $dest"
  exit 0
fi
rm -f "$dest"

floor_args=1
for url in "${mirrors[@]}"; do
  result=0
  try_mirror "$url" || result=$?
  report "$url" "$outcome"
  if ((result == 0)); then failed=0; break; fi
  failed=1
  if ((result == 2)); then slow=$url; fi
done

if ((failed)) && [[ -n $slow ]]; then
  floor_args=''
  attempts=1000000
  result=0
  try_mirror "$slow" || result=$?
  report "$slow" "last resort without a floor: $outcome"
  if ((result == 0)); then failed=0; fi
fi

if ((failed)); then
  rm -f "$partial"
  echo "::error::no mirror delivered the LibreOffice archive with sha256 $sha within ${budget}s"
  exit 1
fi
mv "$partial" "$dest"
echo "LibreOffice archive ready: $dest"
