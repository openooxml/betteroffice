#!/usr/bin/env bun
/**
 * Copy non-JS assets into dist/ after tsup build.
 * tsup bundles JS/TS only; CSS and similar resources have to be copied
 * explicitly so they're reachable through the package's subpath exports.
 */
import { cp, mkdir } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

interface Asset {
  from: string;
  to: string;
}

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');

// the gitignored wasm binaries ship inside dist/ (dist/generated mirrors
// src/wasm/generated so the root-level loader chunks' relative URLs hold)
const assets: Asset[] = [
  { from: 'src/wasm/generated/opc/ooxml_opc_bg.wasm', to: 'dist/generated/opc/ooxml_opc_bg.wasm' },
  {
    from: 'src/wasm/generated/layout/docx_layout_bg.wasm',
    to: 'dist/generated/layout/docx_layout_bg.wasm',
  },
  { from: 'src/wasm/generated/edit/docx_edit_bg.wasm', to: 'dist/generated/edit/docx_edit_bg.wasm' },
  {
    from: 'src/wasm/generated/parse/docx_parse_bg.wasm',
    to: 'dist/generated/parse/docx_parse_bg.wasm',
  },
];

for (const { from, to } of assets) {
  const src = resolve(root, from);
  const dst = resolve(root, to);
  await mkdir(dirname(dst), { recursive: true });
  await cp(src, dst);
  console.log(`[copy-assets] ${from} → ${to}`);
}
