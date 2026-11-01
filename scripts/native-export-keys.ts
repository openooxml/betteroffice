#!/usr/bin/env bun
// usage: bun scripts/native-export-keys.ts > keys.json   (export names per entry of each wasm package a -native package mirrors)
import { join } from 'node:path';
import { FORMATS } from './api-surface';
import { NATIVE_ENGINES, NATIVE_ENTRIES } from './native-engine';

const keys: Record<string, Record<string, string[]>> = {};
for (const format of Object.keys(NATIVE_ENGINES)) {
  if (!FORMATS[`${format}-native`] && !NATIVE_ENTRIES[format]) continue;
  const { packageDir, entries } = NATIVE_ENTRIES[format]
    ? { packageDir: `packages/${format}`, entries: NATIVE_ENTRIES[format] }
    : FORMATS[format];
  keys[format] = {};
  for (const [entry, source] of Object.entries(entries)) {
    keys[format][entry] = Object.keys(await import(join(import.meta.dir, '..', packageDir, source))).sort();
  }
}
console.log(JSON.stringify(keys));
