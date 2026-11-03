import { readFileSync, readdirSync } from 'node:fs';
import { describe, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import type { WorkbookCommand, WorkbookCommandKind, WorkbookCommandPayload } from '../../packages/xlsx/src/session/commandProtocol';
import type { WorkbookSession } from '../../packages/xlsx/src/session/client';
import type { Proposal } from '../../packages/xlsx/src/wasm/loader';

const manifest = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));
const wasmManifest = JSON.parse(readFileSync(new URL('../../packages/xlsx/package.json', import.meta.url), 'utf8'));
const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url));
const dist = new URL('./dist/', import.meta.url);
const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'B2' } };

async function entry(name: string) {
  return import(new URL(manifest.exports[name].import, import.meta.url).href);
}

describe('@betteroffice/xlsx-native', () => {
  test('exposes the entries of @betteroffice/xlsx and refuses browsers', () => {
    expect(Object.keys(manifest.exports)).toEqual(Object.keys(wasmManifest.exports));
    for (const [name, conditions] of Object.entries(wasmManifest.exports) as [string, Record<string, string>][]) {
      expect(manifest.exports[name]).toEqual({
        types: conditions.types,
        browser: './dist/browser.js',
        import: conditions.import,
      });
    }
    expect(manifest.dependencies).toBeUndefined();
    expect(JSON.parse(readFileSync(new URL('package.json', dist), 'utf8'))).toEqual({
      type: 'module',
      imports: {
        '#session-worker': { node: './sessionWorker.node.js', default: './sessionWorker.browser.js' },
        '#session-scope': { node: './sessionScope.node.js', default: './sessionScope.browser.js' },
      },
    });
  });

  test('the browser entry throws', async () => {
    await expect(import(new URL('browser.js', dist).href)).rejects.toThrow(
      'use @betteroffice/xlsx in browsers'
    );
  });

  test('ships no wasm engine', () => {
    for (const file of readdirSync(dist)) {
      expect(file.endsWith('.wasm')).toBe(false);
      const text = readFileSync(new URL(file, dist), 'utf8');
      expect(text).not.toContain('__wbindgen');
      expect(text).not.toContain('_bg.wasm');
    }
  });

  test('opens, recalculates, saves and reopens on the native engine', async () => {
    const instantiate = spyOn(WebAssembly, 'instantiate');
    const compile = spyOn(WebAssembly, 'compile');
    try {
      const xlsx = await entry('.');
      await xlsx.initWasm();
      expect(xlsx.isPngExportAvailable()).toBe(true);
      const workbook = xlsx.openWorkbook(fixture);
      workbook.editCell(0, 1, 1, '=1+2');
      const read = workbook.readCells({ ranges: [target] });
      expect(read.ok).toBe(true);
      expect(read.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 3 });
      const reopened = xlsx.openWorkbook(workbook.save());
      expect(reopened.cell(0, 1, 1).input).toBe('=1+2');
      expect(reopened.readCells({ ranges: [target] }).ranges[0].cells[0][0].displayText).toBe('3');
      workbook.dispose();
      reopened.dispose();
      expect(instantiate).not.toHaveBeenCalled();
      expect(compile).not.toHaveBeenCalled();
    } finally {
      instantiate.mockRestore();
      compile.mockRestore();
    }
  });

  test('runs a worker session on the native engine', async () => {
    const { openWorkbookSession } = await entry('./worker');
    const session = await openWorkbookSession(fixture);
    try {
      const read = await session.call.readCells({ ranges: [{ ...target, sheetId: session.state.sheets[0].id }] });
      expect(read.ok).toBe(true);
      expect(new Uint8Array(await session.call.save()).subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
    } finally {
      await session.dispose();
    }
  });

  test('native sessions admit read-only host proposals through the shared command API', async () => {
    const zip = new JSZip();
    const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
    zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rels}"><sheets><sheet name="Review" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`);
    const { openWorkbookSession, workbookSessionRecovery } = await entry('./worker');
    const session: WorkbookSession = await openWorkbookSession(await zip.generateAsync({ type: 'uint8array' }));
    let seq = 0;
    const send = <K extends WorkbookCommandKind>(kind: K, payload: WorkbookCommandPayload<K>) => session.submit({
      ...session.commandIdentity, basis: session.confirmedBasis, seq: ++seq, kind, payload,
    } as WorkbookCommand);
    try {
      await send('setReadOnly', { readOnly: true, allowHostProposals: true });
      const staged = await send('propose', { args: ['host', null, [{ sheet: 0, row: 0, col: 0, input: '7' }]] });
      expect(staged).toMatchObject({ status: 'applied', committed: false });
      const recovery = workbookSessionRecovery.get(session)!;
      expect(await recovery.capture(() => true)).toBe(true);
      expect(recovery.throughSeq).toBe(2);
      expect(recovery.retainedCommands).toBe(0);
      expect(await send('editCell', { args: [0, 0, 0, '99'] }))
        .toMatchObject({ status: 'refused', reason: 'read-only' });
      expect(await send('acceptProposal', { args: [(staged.result as Proposal).id] }))
        .toMatchObject({ status: 'applied', committed: true });
      const read = await session.call.cellInputs(0, 'A1');
      expect(read.cells[0]![0]!.input).toBe('7');
      expect(session.editState.readOnly).toBe(true);
      const rejected = await send('propose', { args: ['host', null, [{ sheet: 0, row: 0, col: 0, input: '9' }]] });
      expect(rejected).toMatchObject({ status: 'applied', committed: false });
      expect(await send('rejectProposal', { args: [(rejected.result as Proposal).id] }))
        .toMatchObject({ status: 'applied', committed: false });
      const bytes = await recovery.save(() => {});
      expect(bytes).toEqual(await session.save());
      const { openWorkbook } = await entry('.');
      const restored = openWorkbook(bytes);
      try { expect(restored.cell(0, 0, 0).input).toBe('7'); }
      finally { restored.dispose(); }
    } finally { await session.dispose(); }
  });
});
