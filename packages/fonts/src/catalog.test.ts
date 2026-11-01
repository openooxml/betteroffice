import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { exportCatalog } from '../scripts/export-catalog';

test('keeps the native catalog and resolution cases in sync', () => {
  for (const [file, expected] of Object.entries(exportCatalog())) {
    expect(readFileSync(new URL(`../../../crates/betteroffice-fonts/data/${file}`, import.meta.url), 'utf8'))
      .toBe(expected);
  }
});
