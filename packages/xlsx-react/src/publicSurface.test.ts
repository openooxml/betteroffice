import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { XlsxSheetVisibility } from './index';

const index = readFileSync(resolve(import.meta.dir, 'index.ts'), 'utf8');

test('sheet visibility type is exported from the root', () => {
  const visibility: XlsxSheetVisibility[] = ['visible', 'hidden', 'veryHidden', 'unknown'];
  expect(visibility).toHaveLength(4);
  expect(index).toMatch(/export\s+type\s*\{[^}]*\bXlsxSheetVisibility\b[^}]*\}\s*from\s*['"]@betteroffice\/xlsx['"]/);
});
