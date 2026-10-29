import { readdirSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { test } from 'playwright/test';

const assets = resolve(import.meta.dirname, '../../apps/desktop/dist/assets');
let built: boolean | undefined;

function harnessBuilt(): boolean {
  built ??= readdirSync(assets).some((name) =>
    name.endsWith('.js') && readFileSync(join(assets, name), 'utf8').includes('__docxPeerlessHarness'));
  return built;
}

/** Skips unless DOCX_PEERLESS_HARNESS=1, and fails at once when the preview was built without it. */
export function requirePeerlessHarness(): void {
  test.skip(process.env.DOCX_PEERLESS_HARNESS !== '1', 'Needs DOCX_PEERLESS_HARNESS=1 for the desktop build and the run (e2e/README.md)');
  if (!harnessBuilt()) {
    throw new Error('apps/desktop/dist was built without DOCX_PEERLESS_HARNESS=1; rebuild it with ' +
      '`DOCX_PEERLESS_HARNESS=1 bun run --filter @betteroffice/desktop build` (e2e/README.md)');
  }
}
