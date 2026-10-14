import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  exportXlsxMarkdown,
  exportXlsxStructured,
  initWasm,
  openWorkbook,
  renderXlsxMarkdown,
} from '../index';
import type { XlsxExportOptions, XlsxRangeTarget, XlsxStructuredContent } from '../index';

const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/sample.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

function sampleBytes(): Uint8Array {
  return new Uint8Array(readFileSync(FIXTURE));
}

function cell(content: XlsxStructuredContent, sheet: number, a1: string) {
  const found = content.sheets[sheet].cells.find(
    (candidate) => candidate.anchor.kind === 'cell' && candidate.anchor.a1 === a1
  );
  if (!found) throw new Error(`no exported cell ${a1}`);
  return found;
}

function markers(markdown: string): string[] {
  return [...markdown.matchAll(/<!-- xlsx-export:\d+ -->/g)].map((match) => match[0]);
}

describe('xlsx structured export', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('exports the committed workbook with its version and changes nothing', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      let updates = 0;
      handle.onUpdate(() => {
        updates += 1;
      });
      const version = handle.version();
      const history = handle.historyState();
      const saved = handle.save();
      const result = handle.exportStructured();
      if (!result.ok) throw new Error(result.failure.message);
      expect(result.version).toBe(version);
      const content = result.content;
      expect(content.schemaVersion).toBe(1);
      expect(content.anchorScope).toBe('session');
      expect(content.calculation).toEqual({ policy: 'asStored', freshness: 'unverified' });
      expect(content.sheets.map((sheet) => sheet.anchor)).toEqual([
        { kind: 'sheet', sheet: { sheetId: 'sheet:0', index: 0, name: 'Budget' } },
        { kind: 'sheet', sheet: { sheetId: 'sheet:1', index: 1, name: 'Summary' } },
        { kind: 'sheet', sheet: { sheetId: 'sheet:2', index: 2, name: 'Styled' } },
      ]);
      expect(cell(content, 0, 'D3')).toMatchObject({
        id: 's0!D3',
        value: { kind: 'number', value: 157 },
        formula: 'B3+C3',
        displayText: '157',
        formulaResult: 'unverified',
      });
      expect(
        content.diagnostics.some((diagnostic) => diagnostic.code === 'formula-cache-unverified')
      ).toBe(true);

      const markdown = handle.exportMarkdown({ scope: [{ sheet: 0, range: 'A1:D4' }] });
      if (!markdown.ok) throw new Error(markdown.failure.message);
      expect(markdown.version).toBe(version);
      expect(markdown.content.markdown).toContain('## Budget');
      expect(markdown.content.markdown).toContain(
        '<tr><th>1</th><td rowspan="1" colspan="4" data-merge="A1:D1">Quarterly Budget Report</td></tr>'
      );
      expect(markdown.content.markdown).toContain(
        '<tr><th>3</th><td>Line item 1</td><td>100</td><td>57</td><td>157</td></tr>'
      );
      expect(markers(markdown.content.markdown)).toEqual(
        markdown.content.anchors.map((anchor) => anchor.marker)
      );

      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(handle.save()).toEqual(saved);
      expect(updates).toBe(0);
    } finally {
      handle.dispose();
    }
  });

  it('refuses unusable options as data and throws on malformed ones', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const refused = handle.exportStructured({ scope: [{ sheet: 99 }] });
      expect(refused.ok).toBe(false);
      if (!refused.ok) {
        expect(refused.version).toBe(handle.version());
        expect(refused.failure).toEqual({
          code: 'invalid-scope',
          target: null,
          message: 'sheet 99 does not exist; the workbook has 3 sheets',
        });
      }
      const reversed = handle.exportStructured({ scope: [{ sheet: 1, range: 'C3:A1' }] });
      expect(reversed.ok).toBe(false);
      if (!reversed.ok) {
        expect(reversed.failure.target).toEqual({
          kind: 'sheet',
          sheet: { sheetId: 'sheet:1', index: 1, name: 'Summary' },
        });
      }
      const markdown = handle.exportMarkdown({}, { maxRows: 0 });
      expect(markdown.ok).toBe(false);
      if (!markdown.ok) expect(markdown.failure.code).toBe('invalid-options');
      expect(() =>
        handle.exportStructured({ maxCells: 'many' } as unknown as XlsxExportOptions)
      ).toThrow('malformed request');
    } finally {
      handle.dispose();
    }
  });

  it('exports cell anchors that address the same cells as batch targets', () => {
    for (const handle of [
      openWorkbook(sampleBytes()),
      openWorkbook(sampleBytes(), { collaborative: true, clientId: 7201 }),
    ]) {
      try {
        const exported = handle.exportStructured({ scope: [{ sheet: 1 }] });
        if (!exported.ok) throw new Error(exported.failure.message);
        const { anchor } = exported.content.sheets[0].cells[0];
        if (anchor.kind !== 'cell') throw new Error('not a cell anchor');
        const target: XlsxRangeTarget = {
          sheetId: anchor.sheet.sheetId,
          range: { kind: 'a1', a1: anchor.a1 },
        };
        const catalog = handle.readCells({ ranges: [] });
        if (!catalog.ok) throw new Error(catalog.failure.message);
        expect(target.sheetId).toBe(catalog.sheets[1].sheetId);
        const applied = handle.applyEdits({
          expectVersion: exported.version,
          steps: [{ op: 'setCellInputs', target, inputs: [['edited']] }],
        });
        expect(applied.ok).toBe(true);
        const read = handle.readCells({ ranges: [target] });
        if (!read.ok) throw new Error(read.failure.message);
        expect(read.ranges[0].cells[0][0].value).toEqual({ kind: 'text', value: 'edited' });
      } finally {
        handle.dispose();
      }
    }
  });

  it('reads no clock from bytes, while opening a session recalculates volatile cells', async () => {
    const author = openWorkbook(sampleBytes());
    let bytes: Uint8Array;
    try {
      author.editCell(0, 0, 7, '=NOW()');
      bytes = author.save();
    } finally {
      author.dispose();
    }
    const now = Date.now;
    const at = (iso: string) => () => Date.parse(iso);
    const liveValue = () => {
      const handle = openWorkbook(bytes);
      try {
        const result = handle.exportStructured({ scope: [{ sheet: 0, range: 'H1' }] });
        if (!result.ok) throw new Error(result.failure.message);
        return result.content.sheets[0].cells[0].value;
      } finally {
        handle.dispose();
      }
    };
    try {
      Date.now = at('2030-01-01T00:00:00Z');
      const first = await exportXlsxStructured(bytes);
      const firstLive = liveValue();
      Date.now = at('2031-06-15T12:00:00Z');
      const second = await exportXlsxStructured(bytes);
      const secondLive = liveValue();
      expect(second).toEqual(first);
      expect(secondLive).not.toEqual(firstLive);
    } finally {
      Date.now = now;
    }
  });

  it('exports bytes deterministically without a session', async () => {
    const bytes = sampleBytes();
    const first = await exportXlsxStructured(bytes, { maxCells: 10 });
    const second = await exportXlsxStructured(bytes, { maxCells: 10 });
    expect(second).toEqual(first);
    expect(first.anchorScope).toBe('snapshot');
    expect('version' in first).toBe(false);
    expect(first.truncated).toBe(true);
    expect(first.sheets[0].cells).toHaveLength(10);
    expect(first.diagnostics[first.diagnostics.length - 1].code).toBe('truncated');

    const full = await exportXlsxStructured(bytes);
    expect(full.sheets[1].anchor).toEqual({
      kind: 'sheet',
      sheet: { sheetId: 'sheet:1', index: 1, name: 'Summary' },
    });
    const rendered = await renderXlsxMarkdown(full, { maxRows: 5 });
    expect(await exportXlsxMarkdown(bytes, {}, { maxRows: 5 })).toEqual(rendered);
    expect(rendered.markdown.startsWith('<!-- xlsx-export:0 -->\n## Budget')).toBe(true);

    await expect(exportXlsxStructured(new Uint8Array([1, 2, 3]))).rejects.toThrow();
    await expect(exportXlsxStructured(bytes, { maxBytes: 10 })).rejects.toThrow('maxBytes');
    await expect(
      renderXlsxMarkdown({ ...full, schemaVersion: 2 } as unknown as XlsxStructuredContent)
    ).rejects.toThrow('schemaVersion');
  });
});
