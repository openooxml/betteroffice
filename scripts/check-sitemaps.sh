#!/usr/bin/env bash
# usage: check-sitemaps.sh <sitemap-url> [FROM=TO]
# Fetches a sitemap or sitemap index, follows child sitemaps, and requires every listed URL to answer 200 without a redirect.
# FROM=TO rewrites the origin of each listed URL (for checking a local build); robots meta and canonical are checked on HTML pages.
set -u
root=$1
rewrite=${2:-}
fail=0
fetch() { local url=$1; [ -n "$rewrite" ] && url=$(printf '%s' "$url" | sed "s|^${rewrite%%=*}|${rewrite#*=}|"); curl -s -o "$2" -w '%{http_code} %{redirect_url}' "$url"; }
locs() { grep -o '<loc>[^<]*</loc>' "$1" | sed 's/<loc>//;s/<\/loc>//'; }
walk() {
  local map=$1 tmp; tmp=$(mktemp)
  read -r code redirect <<<"$(fetch "$map" "$tmp")"
  if [ "$code" != 200 ]; then echo "FAIL $code $map $redirect"; fail=1; rm -f "$tmp"; return; fi
  if grep -q '<sitemapindex' "$tmp"; then
    echo "index $map"
    for child in $(locs "$tmp"); do walk "$child"; done
  else
    for url in $(locs "$tmp"); do
      page=$(mktemp)
      read -r code redirect <<<"$(fetch "$url" "$page")"
      if [ "$code" != 200 ]; then echo "FAIL $code $url $redirect"; fail=1
      elif grep -qi 'name="robots"[^>]*noindex' "$page"; then echo "FAIL noindex $url"; fail=1
      elif grep -q 'text/html' <<<"$(file -b --mime "$page" 2>/dev/null)" && ! grep -q 'rel="canonical"' "$page"; then echo "WARN no canonical $url"
      else echo "ok   $url"; fi
      rm -f "$page"
    done
  fi
  rm -f "$tmp"
}
walk "$root"
exit $fail
