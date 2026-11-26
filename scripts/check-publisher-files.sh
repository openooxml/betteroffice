#!/usr/bin/env bash
set -euo pipefail

if [ "$(git rev-parse --is-shallow-repository)" = true ]; then
  git fetch --no-tags --depth=1 origin main
else
  git fetch --no-tags origin main
fi

paths=(
  .github/workflows/native-dist.yml
  .github/workflows/publish-native.yml
  .github/workflows/python-dist.yml
  .github/workflows/publish-python-binding.yml
  .github/workflows/deploy.yml
  .github/workflows/wopi-image.yml
  .github/actions/
  package.json
  bun.lock
  apps/demo/scripts/build-collaboration-seeds.ts
  packages/docx/scripts/copy-assets.ts
  scripts/build-docx-wasm.ts
  scripts/build-fonts-decoder.ts
  scripts/build-pptx-wasm.ts
  scripts/build-stale-packages.ts
  scripts/build-vsdx-wasm.ts
  scripts/build-xlsx-wasm.ts
  scripts/cargo-glibc.sh
  scripts/check-packed-license.mjs
  scripts/create-demo-deck.ts
  scripts/create-demo-diagram.ts
  scripts/create-demo-workbook.ts
  scripts/glibc-floor.mjs
  scripts/native-allocator-smoke.mjs
  scripts/native-engine.ts
  scripts/native-raw-smoke.mjs
  scripts/package-boundary.ts
  scripts/python-binding-names.mjs
  scripts/python_font_packs.py
  scripts/rewrite-workspace-deps.ts
  scripts/wasm-cache-key.ts
  scripts/wasm.ts
  scripts/worker-stub.ts
)

if ! git diff --quiet "$GITHUB_SHA" FETCH_HEAD -- "${paths[@]}"; then
  git diff --name-only "$GITHUB_SHA" FETCH_HEAD -- "${paths[@]}" | while IFS= read -r path; do
    echo "::error::Publisher file differs from main: $path"
  done
  exit 1
fi
