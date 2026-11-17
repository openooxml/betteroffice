import { afterAll, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { initWasm, openWorkbook, renderXlsxMarkdown } from '../index';
import type { WorkbookHandle } from '../index';
import { engineKind } from './engine';

const fixture = new Uint8Array(readFileSync(resolve(import.meta.dir, '../../test-fixtures/sample.xlsx')));
const LEAN_WASM = resolve(import.meta.dir, './generated/xlsx_wasm_lean_bg.wasm');

const scope = { scope: [{ sheet: 0, range: 'A1:C5' }] };

describe.skipIf(engineKind === 'native')('the lean core', () => {
  let handle: WorkbookHandle;

  beforeAll(async () => {
    await initWasm(new Uint8Array(readFileSync(LEAN_WASM)), { png: false });
    handle = openWorkbook(fixture.slice());
  }, 30_000);

  afterAll(() => handle.dispose());

  it('rejects rendering standalone structured content', async () => {
    const exported = handle.exportStructured(scope);
    if (!exported.ok) throw new Error('the export should succeed');
    await expect(renderXlsxMarkdown(exported.content)).rejects.toThrow(/not in this build/);
  });

  it('still exports Markdown from an open workbook', () => {
    const exported = handle.exportMarkdown(scope);
    if (!exported.ok) throw new Error('the export should succeed');
    expect(exported.content.markdown.length).toBeGreaterThan(0);
  });
});
