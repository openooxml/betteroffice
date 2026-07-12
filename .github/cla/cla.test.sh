#!/usr/bin/env bash
# Hermetic tests for .github/cla/cla.sh — no network, no git, no gh calls.

set -uo pipefail
cd "$(dirname "$0")/../.."

# shellcheck source=cla.sh
source .github/cla/cla.sh

PASS=0
FAIL=0
TMPDIR=$(mktemp -d)
trap 'rm -rf "$TMPDIR"' EXIT

assert_eq() {
  local name="$1" expected="$2" actual="$3"
  if [ "$expected" = "$actual" ]; then
    printf '  \xe2\x9c\x93 %s\n' "$name"
    PASS=$((PASS + 1))
  else
    printf '  \xe2\x9c\x97 %s\n      expected: %q\n      got:      %q\n' "$name" "$expected" "$actual"
    FAIL=$((FAIL + 1))
  fi
}

assert_true() {
  local name="$1"; shift
  if "$@"; then
    printf '  \xe2\x9c\x93 %s\n' "$name"
    PASS=$((PASS + 1))
  else
    printf '  \xe2\x9c\x97 %s (command returned non-zero)\n' "$name"
    FAIL=$((FAIL + 1))
  fi
}

assert_false() {
  local name="$1"; shift
  if ! "$@"; then
    printf '  \xe2\x9c\x93 %s\n' "$name"
    PASS=$((PASS + 1))
  else
    printf '  \xe2\x9c\x97 %s (command returned zero)\n' "$name"
    FAIL=$((FAIL + 1))
  fi
}

ALLOW='dependabot[bot] github-actions[bot] renovate[bot] jedrazb'

echo "cla_allowlisted:"
assert_true  "bracketed bot name matches literally"      cla_allowlisted "dependabot[bot]" "$ALLOW"
assert_true  "plain user matches"                        cla_allowlisted "jedrazb" "$ALLOW"
assert_false "stranger does not match"                   cla_allowlisted "stranger" "$ALLOW"
assert_false "username prefix is not a partial match"    cla_allowlisted "depend" "$ALLOW"
assert_false "username substring is not a partial match" cla_allowlisted "abot" "$ALLOW"

echo
echo "cla_signed:"
SIGS="$TMPDIR/sigs.json"
echo '{"signedContributors":[{"name":"alice","id":111,"signed_at":"2026-05-09T10:00:00Z","pull_request_no":1}]}' > "$SIGS"
assert_true  "alice (id 111) is recognized" cla_signed 111 "$SIGS"
assert_false "id 222 not recognized"        cla_signed 222 "$SIGS"

echo
echo "cla_add_signature:"
echo '{"signedContributors":[]}' > "$SIGS"
cla_add_signature "bob" 222 "2026-05-09T10:00:00Z" 5 "$SIGS"
assert_eq "first add records bob" \
  '[{"name":"bob","id":222,"signed_at":"2026-05-09T10:00:00Z","pull_request_no":5}]' \
  "$(jq -c .signedContributors "$SIGS")"

cla_add_signature "carol" 333 "2026-05-09T10:01:00Z" 6 "$SIGS"
assert_eq "second add appends carol" \
  '2' \
  "$(jq '.signedContributors | length' "$SIGS")"

echo
echo "sign-once guarantee (the load-bearing test):"
echo '{"signedContributors":[{"name":"alice","id":111,"signed_at":"2026-05-09T10:00:00Z","pull_request_no":1}]}' > "$SIGS"
assert_true "returning signer recognized → caller skips append+push" \
  cla_signed 111 "$SIGS"

before_count=$(jq '.signedContributors | length' "$SIGS")
if cla_signed 111 "$SIGS"; then
  : # gate fired — no append
else
  cla_add_signature "alice" 111 "would-not-happen" 99 "$SIGS"
fi
after_count=$(jq '.signedContributors | length' "$SIGS")
assert_eq "duplicate sign comment is a no-op (count unchanged)" \
  "$before_count" "$after_count"

echo
echo "cla_org_member (mocked):"
# hermetic stub: only alice and carol are members; empty $org returns 1
cla_org_member() {
  [ -z "$2" ] && return 1
  case "$1" in
    alice|carol) return 0 ;;
    *) return 1 ;;
  esac
}
assert_true  "alice is a member of openooxml"   cla_org_member "alice" "openooxml"
assert_false "stranger is not a member"        cla_org_member "stranger" "openooxml"
assert_false "no org configured → not a match" cla_org_member "alice" ""

echo
echo "cla_should_skip (allowlist + org-member combined):"
assert_true  "literal allowlist match → skip"           cla_should_skip "dependabot[bot]" "$ALLOW" ""
assert_true  "named maintainer match → skip"            cla_should_skip "jedrazb"         "$ALLOW" ""
assert_false "non-allowlisted, no org configured → don't skip" \
                                                        cla_should_skip "stranger"        "$ALLOW" ""
assert_true  "org member match (mocked alice) → skip"   cla_should_skip "alice"           "$ALLOW" "openooxml"
assert_false "non-allowlisted, non-org-member → don't skip" \
                                                        cla_should_skip "stranger"        "$ALLOW" "openooxml"
assert_true  "literal allowlist still wins when org configured" \
                                                        cla_should_skip "jedrazb"         "$ALLOW" "openooxml"
assert_false "empty allowlist + empty org → no one skipped" \
                                                        cla_should_skip "anyone"          ""        ""

echo
echo "cla_render_unsigned_comment:"
body=$(cla_render_unsigned_comment "https://e.x/CLA.md" "I sign" "<!-- m -->" "octocat")
case "$body" in "@octocat "*) ok=1 ;; *) ok=0 ;; esac
assert_eq "starts with @-mention of the PR author (notifies them)" "1" "$ok"
case "$body" in *"https://e.x/CLA.md"*) ok=1 ;; *) ok=0 ;; esac
assert_eq "includes CLA URL" "1" "$ok"
case "$body" in *"I sign"*) ok=1 ;; *) ok=0 ;; esac
assert_eq "includes the sign phrase" "1" "$ok"
case "$body" in *'`!cla-check`'*) ok=1 ;; *) ok=0 ;; esac
assert_eq "mentions the !cla-check keyword in a code span" "1" "$ok"
case "$body" in *"<!-- m -->"*) ok=1 ;; *) ok=0 ;; esac
assert_eq "includes sticky-comment marker" "1" "$ok"

echo
echo "cla_render_signed_comment:"
body=$(cla_render_signed_comment "<!-- m -->")
case "$body" in *"All contributors have signed the CLA"*) ok=1 ;; *) ok=0 ;; esac
assert_eq "matches action's all-signed wording" "1" "$ok"
case "$body" in *"✍️"*"✅"*) ok=1 ;; *) ok=0 ;; esac
assert_eq "includes the celebratory emoji" "1" "$ok"
case "$body" in *"<!-- m -->"*) ok=1 ;; *) ok=0 ;; esac
assert_eq "includes sticky-comment marker" "1" "$ok"

echo
echo "GraphQL PR-author extraction (regression guard):"
# shaped like cla_main's query; a field rename would silently let unsigned PRs through
gql_normal='{"data":{"repository":{"pullRequest":{"headRefOid":"abc123","author":{"login":"leandrotcawork","databaseId":99999}}}}}'
assert_eq "headRefOid extracts" "abc123" \
  "$(echo "$gql_normal" | jq -r '.data.repository.pullRequest.headRefOid')"
assert_eq "author login extracts" "leandrotcawork" \
  "$(echo "$gql_normal" | jq -r '.data.repository.pullRequest.author.login // empty')"
assert_eq "author databaseId extracts" "99999" \
  "$(echo "$gql_normal" | jq -r '.data.repository.pullRequest.author.databaseId // empty')"

gql_deleted='{"data":{"repository":{"pullRequest":{"headRefOid":"def456","author":null}}}}'
assert_eq "deleted-author login → empty" "" \
  "$(echo "$gql_deleted" | jq -r '.data.repository.pullRequest.author.login // empty')"
assert_eq "deleted-author id → empty" "" \
  "$(echo "$gql_deleted" | jq -r '.data.repository.pullRequest.author.databaseId // empty')"

# GraphQL returns Bot logins without the "[bot]" suffix (verified on a real dependabot PR)
gql_bot='{"data":{"repository":{"pullRequest":{"headRefOid":"bot789","author":{"__typename":"Bot","login":"dependabot","databaseId":49699333}}}}}'
assert_eq "bot author __typename extracts" "Bot" \
  "$(echo "$gql_bot" | jq -r '.data.repository.pullRequest.author.__typename // empty')"
assert_eq "bot author bare login (pre-normalization) extracts" "dependabot" \
  "$(echo "$gql_bot" | jq -r '.data.repository.pullRequest.author.login // empty')"
assert_eq "bot author databaseId extracts" "49699333" \
  "$(echo "$gql_bot" | jq -r '.data.repository.pullRequest.author.databaseId // empty')"

normalize_bot_login() {
  local type="$1" login="$2"
  if [ "$type" = "Bot" ] && [[ "$login" != *"[bot]" ]]; then
    echo "${login}[bot]"
  else
    echo "$login"
  fi
}
assert_eq "Bot type without [bot] suffix → suffix appended" "dependabot[bot]" \
  "$(normalize_bot_login Bot dependabot)"
assert_eq "Bot type that already has [bot] → unchanged" "renovate[bot]" \
  "$(normalize_bot_login Bot 'renovate[bot]')"
assert_eq "User type → never modified" "octocat" \
  "$(normalize_bot_login User octocat)"

echo
echo "cla_init_signatures:"
NEW="$TMPDIR/sub/cla.json"
cla_init_signatures "$NEW"
[ -f "$NEW" ] && exists=1 || exists=0
assert_eq "creates parent directory + file" "1" "$exists"
assert_eq "initial structure is empty array" \
  '{"signedContributors":[]}' "$(jq -c . "$NEW")"

echo '{"signedContributors":[{"name":"x","id":1,"signed_at":"y","pull_request_no":1}]}' > "$NEW"
cla_init_signatures "$NEW"
assert_eq "existing signatures file is preserved" \
  "x" "$(jq -r '.signedContributors[0].name' "$NEW")"

echo
echo "Results: $PASS passed, $FAIL failed"
[ "$FAIL" -eq 0 ]
