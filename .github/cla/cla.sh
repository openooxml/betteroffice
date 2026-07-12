#!/usr/bin/env bash
# CLA workflow logic; sourced by cla.yml (cla_main) and cla.test.sh (pure fns).
# Side effects live only in cla_main.

set -euo pipefail

# 0 if $login is in space-separated $allowlist; the quoted expansion keeps
# "[bot]" names literal.
cla_allowlisted() {
  local login="$1" allowlist="$2"
  [[ " $allowlist " == *" $login "* ]]
}

# 0 if $user_id is already in $signatures_file's signedContributors.
cla_signed() {
  local user_id="$1" signatures_file="$2"
  local count
  count=$(jq --argjson id "$user_id" \
    '[.signedContributors[] | select(.id == $id)] | length' \
    "$signatures_file")
  [ "$count" != "0" ]
}

# Always appends; caller must gate on cla_signed.
cla_add_signature() {
  local name="$1" user_id="$2" ts="$3" pr="$4" signatures_file="$5"
  jq --arg name "$name" --argjson id "$user_id" --arg ts "$ts" --argjson pr "$pr" \
    '.signedContributors += [{name: $name, id: $id, signed_at: $ts, pull_request_no: $pr}]' \
    "$signatures_file" > "$signatures_file.tmp"
  mv "$signatures_file.tmp" "$signatures_file"
}

# 0 if $login is a public member of $org; fails closed on private members and
# API errors. Tests stub this.
cla_org_member() {
  local login="$1" org="$2"
  [ -z "$org" ] && return 1
  gh api "orgs/$org/members/$login" --silent 2>/dev/null
}

# 0 if $login is exempt: allowlisted or public org member.
cla_should_skip() {
  local login="$1" allowlist="$2" org="${3:-}"
  cla_allowlisted "$login" "$allowlist" && return 0
  [ -n "$org" ] && cla_org_member "$login" "$org" && return 0
  return 1
}

# The leading @-mention notifies once; sticky-comment edits don't re-notify.
cla_render_unsigned_comment() {
  local cla_url="$1" sign_phrase="$2" marker="$3" pr_author_login="$4"
  local ccla_url="${cla_url%CLA.md}CCLA.md"
  cat <<EOF
@${pr_author_login} thanks a lot for the contribution! Before we can merge it, we need your one-time signature on the [OpenOOXML Contributor License Agreement](${cla_url}). To sign, post a comment on this pull request with exactly:

\`\`\`
${sign_phrase}
\`\`\`

Contributing as part of your job? Then please also have your employer sign the [Corporate CLA](${ccla_url}) — your signature covers you, theirs covers them.

The check updates on its own once you have signed; you can also re-run it any time by commenting \`!cla-check\`.

${marker}
EOF
}

cla_render_signed_comment() {
  local marker="$1"
  cat <<EOF
All contributors have signed the CLA — thank you! ✍️ ✅

<sub>Posted by the CLA bot.</sub>

${marker}
EOF
}

cla_init_signatures() {
  local signatures_file="$1"
  mkdir -p "$(dirname "$signatures_file")"
  [ -f "$signatures_file" ] || echo '{"signedContributors":[]}' > "$signatures_file"
}

# Orchestrates the workflow; the only function with side effects. Env: REPO,
# PR_NUMBER, EVENT_NAME, ALLOWLIST, CLA_URL, SIGN_PHRASE (+ COMMENT_USER_*).
cla_main() {
  local signatures="signatures/v1/cla.json"
  local marker='<!-- cla-bot -->'

  cla_init_signatures "$signatures"

  if [ "${EVENT_NAME:-}" = "issue_comment" ]; then
    if cla_should_skip "$COMMENT_USER_LOGIN" "$ALLOWLIST" "${CLA_ORG:-}"; then
      : # exempt — no signature row needed
    elif ! cla_signed "$COMMENT_USER_ID" "$signatures"; then
      cla_add_signature "$COMMENT_USER_LOGIN" "$COMMENT_USER_ID" \
        "$(date -u +%Y-%m-%dT%H:%M:%SZ)" "$PR_NUMBER" "$signatures"
      git config user.name "$GIT_AUTHOR_NAME"
      git config user.email "$GIT_AUTHOR_EMAIL"
      git add "$signatures"
      git commit -m "Record CLA signature for @${COMMENT_USER_LOGIN} (PR #${PR_NUMBER})"
      # the push works because the App is on the main ruleset's bypass list
      git push origin main
    fi
  fi

  # the PR author is the signer of record; commit authors are not checked
  local pr_data pr_author_login pr_author_id head_sha
  # databaseId must be selected per concrete Actor type
  pr_data=$(gh api graphql \
    -F owner="${REPO%/*}" -F name="${REPO#*/}" -F number="$PR_NUMBER" \
    -f query='
      query($owner:String!, $name:String!, $number:Int!) {
        repository(owner:$owner, name:$name) {
          pullRequest(number:$number) {
            headRefOid
            author {
              __typename
              login
              ... on User { databaseId }
              ... on Bot { databaseId }
            }
          }
        }
      }')
  head_sha=$(echo "$pr_data" | jq -r '.data.repository.pullRequest.headRefOid')
  local pr_author_type
  pr_author_type=$(echo "$pr_data" | jq -r '.data.repository.pullRequest.author.__typename // empty')
  pr_author_login=$(echo "$pr_data" | jq -r '.data.repository.pullRequest.author.login // empty')
  pr_author_id=$(echo "$pr_data" | jq -r '.data.repository.pullRequest.author.databaseId // empty')

  # GraphQL returns Bot logins bare; the allowlist uses the "[bot]" suffix
  if [ "$pr_author_type" = "Bot" ] && [[ "$pr_author_login" != *"[bot]" ]]; then
    pr_author_login="${pr_author_login}[bot]"
  fi

  if [ -z "$pr_author_login" ]; then
    echo "ERROR: PR #${PR_NUMBER} has no identifiable GitHub author (deleted account?)" >&2
    exit 1
  fi

  local pr_author_signed=false
  if cla_should_skip "$pr_author_login" "$ALLOWLIST" "${CLA_ORG:-}"; then
    pr_author_signed=true
  elif [ -n "$pr_author_id" ] && cla_signed "$pr_author_id" "$signatures"; then
    pr_author_signed=true
  fi

  local body status_state status_desc
  if "$pr_author_signed"; then
    status_state="success"
    status_desc="PR author has signed the CLA"
    body=$(cla_render_signed_comment "$marker")
  else
    status_state="failure"
    status_desc="Awaiting CLA signature from PR author"
    body=$(cla_render_unsigned_comment "$CLA_URL" "$SIGN_PHRASE" "$marker" "$pr_author_login")
  fi

  # upsert the sticky comment; -s is load-bearing (--paginate emits one array per page)
  local existing
  existing=$(gh api "repos/${REPO}/issues/${PR_NUMBER}/comments" --paginate \
    | jq -rs --arg m "$marker" 'add | [.[] | select(.body | contains($m)) | .id] | first // empty')
  if [ -n "$existing" ]; then
    gh api -X PATCH "repos/${REPO}/issues/comments/${existing}" -f body="$body" > /dev/null
  else
    gh api -X POST "repos/${REPO}/issues/${PR_NUMBER}/comments" -f body="$body" > /dev/null
  fi

  gh api -X POST "repos/${REPO}/statuses/${head_sha}" \
    -f state="$status_state" \
    -f context="CLA" \
    -f description="$status_desc" \
    -f target_url="$CLA_URL" > /dev/null
}

if [ "${BASH_SOURCE[0]:-}" = "${0}" ]; then
  cla_main
fi
