#!/usr/bin/env bash
set -euo pipefail

script_dir=$(cd "$(dirname "$0")" && pwd)
repo_root=$(cd "$script_dir/../../.." && pwd)
cd "$repo_root"
bun install --frozen-lockfile
bun run build:docx-wasm
bun run build:xlsx-wasm
bun run build:pptx-wasm
bun run build:packages
bun run --filter @betteroffice/desktop build
