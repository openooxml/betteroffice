import { describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx/parser';
import { renderEnvironment } from '../docx/structuredExport';
import { buildResidentRegionLayoutRequest } from './computeLayout';

const ROOT = resolve(import.meta.dir, '../../../..');
const FIXTURES = resolve(ROOT, 'crates/docx-edit/tests/fixtures/layout-request');
const CASES: Record<string, string> = JSON.parse(
  readFileSync(resolve(FIXTURES, 'cases.json'), 'utf8')
);

describe('DOCX layout request parity', () => {
  test('every case has a golden request', () => {
    expect(readdirSync(FIXTURES).filter((name) => name.endsWith('.request.json')).sort()).toEqual(
      Object.keys(CASES).map((name) => `${name}.request.json`).sort()
    );
  });

  for (const [name, source] of Object.entries(CASES)) {
    test(`${name} matches the Rust request`, async () => {
      const bytes = new Uint8Array(readFileSync(resolve(ROOT, source)));
      const document = await parseDocx(bytes.buffer, { preloadFonts: false });
      const request = buildResidentRegionLayoutRequest(
        document,
        24,
        renderEnvironment(document, undefined)
      );
      const expected = JSON.parse(readFileSync(resolve(FIXTURES, `${name}.request.json`), 'utf8'));
      expect(JSON.parse(JSON.stringify(request))).toEqual(expected);
    }, 30_000);
  }
});
