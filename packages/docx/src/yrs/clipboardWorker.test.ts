import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from './wasm/index';
import { preloadOpcWasm } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import type { DocxCommandEnvelope, ResidentCommandInitRequest } from './residentCommandProtocol';
import type { DocxClipboardFragment } from './clipboardFragment';
import type { YrsCellLoc } from './index';

const TIMEOUT = 20_000;
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});

function seedDoc(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  ));
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>' +
    '<w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function intent(seq: number, version: number, kind: string, payload: DocxCommandEnvelope['payload']): DocxCommandEnvelope {
  return { session: 'tclip-worker', incarnation: 1, seq, basis: { epoch: 1, version }, kind, payload };
}

function cellLoc(story: string, tableIndex: number, row: number, column: number): YrsCellLoc {
  return { story, tableIndex, row, column };
}

function bodyTables(engine: { storySegments(story: string): { kind: string; embedKind?: string }[] }): number {
  return engine.storySegments('body').filter((segment) => segment.kind === 'embed' && segment.embedKind === 'table').length;
}

async function openTableWorker() {
  const worker = startWorker(8401);
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await client.open(seedDoc(), { generation: 'tclip-table' });
    const engine = worker.sessions[0]!;
    const anchor = engine.paragraphs('body')[0]!;
    const receipt = engine.insertTable({ story: 'body', paraId: anchor.paraId, offset: 0 }, 2, 2);
    const payload = engine.tablePayload('body', 0);
    const rows = (payload?.rows ?? []) as Array<{ cells: Array<{ story: string }> }>;
    const texts = ['a', 'b', 'c', 'd'];
    rows.forEach((row, rowIndex) =>
      row.cells.forEach((cell, columnIndex) => {
        const paraId = engine.paragraphs(cell.story)[0]!.paraId;
        engine.insertText({ story: cell.story, paraId, offset: 0 }, texts[rowIndex * 2 + columnIndex]!);
      }),
    );
    const point = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 0 };
    const fonts = [new Uint8Array(readFileSync(resolve(
      process.cwd(), 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
    )))];
    const fontId = engine.registerFont(fonts[0]!);
    const measurement = {
      defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
      fontChains: { 'liberation sans|0|0': [fontId] },
    };
    const layoutInput = JSON.stringify({ ...JSON.parse(LAYOUT), measurement });
    const extras = JSON.stringify({ fontChains: { ...measurement.fontChains, 'calibri|0|0': [fontId] } });
    const init: ResidentCommandInitRequest = {
      session: 'tclip-worker', incarnation: 1, epoch: 1, readOnly: false,
      selection: { anchor: point, focus: point },
      fonts, layoutInput, layoutWithRegions: true, layoutRevision: 1, extras,
    };
    const initial = await client.commandInit(init);
    return { worker, client, engine, table: receipt.table, initial };
  } catch (error) {
    client.destroy();
    throw error;
  }
}

test('clipboard slice reads a multi-cell selection as text and fragment', async () => {
  const { client, engine } = await openTableWorker();
  try {
    engine.setCellSelection({
      anchor: cellLoc('body', 0, 0, 0),
      head: cellLoc('body', 0, 1, 1),
    });
    const payload = engine.tablePayload('body', 0);
    const cellStory = ((payload?.rows ?? []) as Array<{ cells: Array<{ story: string }> }>)[0]?.cells[0]?.story;
    if (!cellStory) throw new Error('cell story is missing');
    const paraId = engine.paragraphs(cellStory)[0]!.paraId;
    const focus = { story: cellStory, paraId, offset: 1 };
    const { value } = await client.documentRead({
      kind: 'clipboardSlice',
      selection: { anchor: focus, focus },
    });
    expect(value.text).toBe('a\tb\nc\td');
    expect(value.fragment?.paragraphs).toHaveLength(1);
    expect(value.fragment?.paragraphs[0]?.tables?.[0]?.rows).toHaveLength(2);
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('clipboard slice reads a body range spanning a table', async () => {
  const { client, engine } = await openTableWorker();
  try {
    const paragraphs = engine.paragraphs('body');
    const last = paragraphs[paragraphs.length - 1]!;
    const { value } = await client.documentRead({
      kind: 'clipboardSlice',
      selection: {
        anchor: { story: 'body', paraId: paragraphs[0]!.paraId, offset: 0 },
        focus: { story: 'body', paraId: last.paraId, offset: last.text.length },
      },
    });
    expect(value.text).toContain('a\tb\nc\td\n');
    const tables = value.fragment?.paragraphs.flatMap((paragraph) => paragraph.tables ?? []);
    expect(tables).toHaveLength(1);
    expect(tables?.[0]?.rows).toHaveLength(2);
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('paste creates an identical table and one undo removes it', async () => {
  const { client, engine } = await openTableWorker();
  try {
    let version = 0;
    let seq = 0;
    const paragraphs = engine.paragraphs('body');
    const last = paragraphs[paragraphs.length - 1]!;
    const copy = await client.documentRead({
      kind: 'clipboardSlice',
      selection: {
        anchor: { story: 'body', paraId: paragraphs[0]!.paraId, offset: 0 },
        focus: { story: 'body', paraId: last.paraId, offset: last.text.length },
      },
    });
    expect(copy.value.fragment).not.toBeNull();
    const span = engine.geometryReader.locateParagraph('body', last.paraId);
    const split = engine.splitParagraph({ story: 'body', paraId: last.paraId, offset: span.end - span.start });
    const target = { story: 'body', paraId: split.secondParaId, offset: 0 };
    const pasted = await client.edit(intent(++seq, version, 'pasteFragment', {
      selection: { anchor: target, focus: target },
      fragment: copy.value.fragment!,
    }));
    expect(pasted.outcomes[0]?.status).toBe('applied');
    version = pasted.version;
    expect(bodyTables(engine)).toBe(2);
    const first = engine.tablePayload('body', 0);
    const second = engine.tablePayload('body', 1);
    expect(second?.tblPr).toEqual(first?.tblPr);
    expect(second?.grid).toEqual(first?.grid);
    const textsOf = (index: number): string[] => {
      const rows = (engine.tablePayload('body', index)?.rows ?? []) as Array<{ cells: Array<{ story: string }> }>;
      return rows.flatMap((row) =>
        row.cells.map((cell) =>
          engine.storySegments(cell.story).filter((segment) => segment.kind === 'text').map((segment) =>
            segment.kind === 'text' ? segment.text : '').join(''),
        ),
      );
    };
    expect(textsOf(1)).toEqual(textsOf(0));
    expect(pasted.selection).toMatchObject({ anchor: { story: 'body' } });
    const undone = await client.edit(intent(++seq, version, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect(bodyTables(engine)).toBe(1);
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('paste replays text, tables and splits across paragraphs', async () => {
  const { client, engine } = await openTableWorker();
  try {
    let version = 0;
    let seq = 0;
    const paragraphs = engine.paragraphs('body');
    const last = paragraphs[paragraphs.length - 1]!;
    const copy = await client.documentRead({
      kind: 'clipboardSlice',
      selection: {
        anchor: { story: 'body', paraId: paragraphs[0]!.paraId, offset: 0 },
        focus: { story: 'body', paraId: last.paraId, offset: last.text.length },
      },
    });
    const table = copy.value.fragment?.paragraphs[0]?.tables?.[0];
    expect(table).toBeDefined();
    const fragment: DocxClipboardFragment = {
      paragraphs: [{ text: 'pre' }, { tables: [table!], text: 'mid' }, { text: 'post' }],
    };
    const refreshed = engine.paragraphs('body');
    const atPara = refreshed[refreshed.length - 1]!;
    const atSpan = engine.geometryReader.locateParagraph('body', atPara.paraId);
    const at = { story: 'body', paraId: atPara.paraId, offset: atSpan.end - atSpan.start };
    const pasted = await client.edit(intent(++seq, version, 'pasteFragment', {
      selection: { anchor: at, focus: at },
      fragment,
    }));
    expect(pasted.outcomes[0]?.status).toBe('applied');
    version = pasted.version;
    expect(bodyTables(engine)).toBe(2);
    expect(engine.paragraphs('body').map((paragraph) => paragraph.text)).toContain('Seedpre');
    expect(engine.paragraphs('body').map((paragraph) => paragraph.text)).toContain('mid');
    expect(engine.paragraphs('body').map((paragraph) => paragraph.text)).toContain('post');
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('paste into a table cell nests the table', async () => {
  const { client, engine } = await openTableWorker();
  try {
    let version = 0;
    let seq = 0;
    const payload = engine.tablePayload('body', 0);
    const story = ((payload?.rows ?? []) as Array<{ cells: Array<{ story: string }> }>)[1]?.cells[1]?.story;
    if (!story) throw new Error('cell story is missing');
    const paraId = engine.paragraphs(story)[0]!.paraId;
    const at = { story, paraId, offset: 1 };
    const fragment: DocxClipboardFragment = {
      paragraphs: [{ tables: [{
        tblPr: {},
        grid: [],
        rows: [{ trPr: {}, cells: [{
          tcPr: {},
          content: [
            { kind: 'text', text: 'nested', attributes: {} },
            { kind: 'pilcrow', properties: {}, attributes: {} },
          ],
        }] }],
      }], text: '' }],
    };
    const pasted = await client.edit(intent(++seq, version, 'pasteFragment', {
      selection: { anchor: at, focus: at },
      fragment,
    }));
    expect(pasted.outcomes[0]?.status).toBe('applied');
    const nested = engine.storySegments(story).filter((segment) => segment.kind === 'embed' && segment.embedKind === 'table');
    expect(nested).toHaveLength(1);
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('paste refuses invalid payloads and read-only sessions without moving the version', async () => {
  const { client, engine } = await openTableWorker();
  try {
    let version = 0;
    let seq = 0;
    const before = engine.encodeStateVector();
    const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 0 };
    const invalid = await client.edit(intent(++seq, version, 'pasteFragment', {
      selection: { anchor: at, focus: at },
      fragment: { paragraphs: [] },
    }));
    expect(invalid.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    expect(invalid.version).toBe(version);
    expect(engine.encodeStateVector()).toEqual(before);
    expect(bodyTables(engine)).toBe(1);
    const locked = await client.edit(intent(++seq, version, 'setReadOnly', { readOnly: true }));
    expect(locked.outcomes[0]?.status).toBe('noop');
    const fragment: DocxClipboardFragment = {
      paragraphs: [{ text: 'x' }],
    };
    const refused = await client.edit(intent(++seq, version, 'pasteFragment', {
      selection: { anchor: at, focus: at },
      fragment,
    }));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'read-only' });
    expect(engine.encodeStateVector()).toEqual(before);
  } finally {
    client.destroy();
  }
}, TIMEOUT);
