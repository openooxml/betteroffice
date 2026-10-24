import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from './wasm/index';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import type { DocxCommandEnvelope, DocxCommandInitRequest } from './residentCommandProtocol';
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

async function openTableWorker(cellObject?: { kind: string; depth: number; controlKind: string | null }) {
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
    if (cellObject) {
      let story = rows[0]!.cells[0]!.story;
      for (let level = 0; level < cellObject.depth; level++) {
        engine.insertTable({ story, paraId: engine.paragraphs(story)[0]!.paraId, offset: 0 }, 1, 1);
        const nested = engine.tablePayload(story, 0)?.rows as Array<{ cells: Array<{ story: string }> }>;
        story = nested[0]!.cells[0]!.story;
      }
      const imagePayload = { src: 'data:image/png;base64,AA==' };
      const payload = cellObject.kind === 'chart' ? { chartJson: '{}', width: 8, height: 8 } : imagePayload;
      const object = { op: 'insertEmbed' as const, index: 0, kind: cellObject.kind, payload };
      if (cellObject.controlKind === 'sdt') {
        engine.applyRawOps(story, [{ ...object, kind: 'sdt', payload: { content: [{ kind: 'text', text: 'kept' }, { kind: cellObject.kind, payload }] } }]);
      } else if (cellObject.controlKind === 'blockSdt') {
        engine.insertTable({ story, paraId: engine.paragraphs(story)[0]!.paraId, offset: 0 }, 1, 1);
        const controlRows = engine.tablePayload(story, 0)?.rows as Array<{ cells: Array<{ story: string }> }>;
        const controlStory = controlRows[0]!.cells[0]!.story;
        engine.applyRawOps(controlStory, [{ op: 'insert', index: 0, text: 'kept' }, { ...object, index: 4 }]);
        engine.applyRawOps(story, [{ op: 'delete', index: 0, len: 1 }, { ...object, kind: 'blockSdt', payload: { story: controlStory } }]);
      } else engine.applyRawOps(story, [object]);
      engine.applyRawOps('body', [{ ...object, index: 1, kind: 'image', payload: imagePayload }]);
    }
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
    const init: DocxCommandInitRequest = {
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

for (const kind of ['chart', 'shape', 'image']) {
  for (const depth of [0, 2]) {
    for (const controlKind of [null, 'sdt', 'blockSdt']) {
      test(`worker clipboard refuses a picture with ${kind} in a table at depth ${depth} through ${controlKind ?? 'no control'}`, async () => {
        const { client, engine } = await openTableWorker({ kind, depth, controlKind });
        try {
          const paragraphs = engine.paragraphs('body');
          const last = paragraphs.at(-1)!;
          const span = engine.geometryReader.locateParagraph('body', last.paraId);
          const before = engine.encodeStateVector();
          const segments = engine.storySegments('body');
          await expect(client.documentRead({ kind: 'clipboardSlice', selection: {
            anchor: { story: 'body', paraId: paragraphs[0]!.paraId, offset: 0 },
            focus: { story: 'body', paraId: last.paraId, offset: span.end - span.start },
          } })).rejects.toThrow('Picture selection contains unsupported clipboard objects');
          expect(engine.encodeStateVector()).toEqual(before);
          expect(engine.storySegments('body')).toEqual(segments);
        } finally {
          client.destroy();
        }
      }, TIMEOUT);
    }
  }
}

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


for (const another of [false, true]) {
  test(`mixed picture paste into ${another ? 'another' : 'the same'} worker creates independent media and one undo step`, async () => {
    const source = await openTableWorker();
    const target = another ? await openTableWorker() : source;
    const data = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aD1sAAAAASUVORK5CYII=';
    const images = (engine: typeof source.engine) => engine.storySegments('body').filter(segment => segment.kind === 'embed' && segment.embedKind === 'image');
    try {
      const freshParagraph = (engine: typeof source.engine) => {
        const last = engine.paragraphs('body').at(-1)!;
        const lastSpan = engine.geometryReader.locateParagraph('body', last.paraId);
        const split = engine.splitParagraph({ story: 'body', paraId: last.paraId, offset: lastSpan.end - lastSpan.start });
        return { story: 'body', paraId: split.secondParaId, offset: 0 };
      };
      const at = freshParagraph(source.engine);
      const typed = await source.client.edit(intent(1, 0, 'insertText', {
        selection: { anchor: at, focus: at }, text: 'abcd',
      }));
      expect(typed.outcomes[0]?.status).toBe('applied');
      const imageAt = { ...at, offset: 2 };
      const pictured = await source.client.edit(intent(2, typed.version, 'insertImage', {
        selection: { anchor: imageAt, focus: imageAt },
        image: { src: `data:image/png;base64,${data}`, alt: 'picture', width: 120, height: 60 },
      }));
      expect(pictured.outcomes[0]?.status).toBe('applied');
      let seq = another ? 0 : 2;
      const version = another ? 0 : pictured.version;
      const copy = await source.client.documentRead({ kind: 'clipboardSlice', selection: { anchor: at, focus: { ...at, offset: 5 } } });
      expect(copy.value.text).toBe('abcd');
      expect(copy.value.fragment).not.toBeNull();
      const destination = freshParagraph(target.engine);
      const destinationPara = target.engine.paragraphs('body').at(-1)!;
      const beforeImages = images(target.engine).length;
      const pasted = await target.client.edit(intent(++seq, version, 'pasteFragment', {
        selection: { anchor: destination, focus: destination }, fragment: copy.value.fragment!,
      }));
      expect(pasted.outcomes[0]?.status).toBe('applied');
      expect(images(target.engine)).toHaveLength(beforeImages + 1);
      const saved = await target.client.save({ comments: [] });
      const parts = unzipContainer(new Uint8Array(saved.bytes));
      const xml = new TextDecoder().decode(parts['word/document.xml']);
      const rels = new TextDecoder().decode(parts['word/_rels/document.xml.rels']);
      const types = new TextDecoder().decode(parts['[Content_Types].xml']);
      const ids = [...xml.matchAll(/<wp:docPr[^>]* id="([^"]+)"/g)].map(match => match[1]);
      const embeds = [...xml.matchAll(/r:embed="([^"]+)"/g)].map(match => match[1]!);
      expect(ids).toHaveLength(beforeImages + 1);
      expect(new Set(ids).size).toBe(ids.length);
      expect(embeds).toHaveLength(beforeImages + 1);
      expect(new Set(embeds).size).toBe(embeds.length);
      const mediaPaths = embeds.map(id => {
        const relationship = [...rels.matchAll(/<Relationship\b[^>]+>/g)].find(match => match[0].includes(`Id="${id}"`))?.[0];
        const path = /Target="([^"]+)"/.exec(relationship ?? '')?.[1];
        if (!path) throw new Error('Missing picture relationship');
        return path.startsWith('/') ? path.slice(1) : `word/${path}`;
      });
      expect(new Set(mediaPaths).size).toBe(mediaPaths.length);
      for (const path of mediaPaths) {
        expect(parts[path]).toEqual(new Uint8Array(Buffer.from(data, 'base64')));
        expect(types).toContain('image/png');
      }
      const reopenedWorker = startWorker(8402);
      const reopened = new ResidentEngineWorkerClient(reopenedWorker);
      try {
        await reopened.open(new Uint8Array(saved.bytes), { generation: 'clipboard-reopened' });
        const reopenedAt = { story: 'body', paraId: reopenedWorker.sessions[0]!.paragraphs('body').at(-1)!.paraId, offset: 0 };
        const read = await reopened.documentRead({ kind: 'clipboardSlice', selection: {
          anchor: reopenedAt, focus: { ...reopenedAt, offset: 5 },
        } });
        expect(read.value.text).toBe('abcd');
        expect(read.value.fragment?.paragraphs[0]?.inlines?.[1]).toMatchObject({ image: { src: `data:image/png;base64,${data}` } });
      } finally { reopened.destroy(); }
      const undone = await target.client.edit(intent(++seq, pasted.version, 'undo', {}));
      expect(undone.outcomes[0]?.status).toBe('applied');
      expect(images(target.engine)).toHaveLength(beforeImages);
      expect(target.engine.paragraphs('body').at(-1)!.text).toBe(destinationPara.text);
    } finally {
      source.client.destroy();
      if (another) target.client.destroy();
    }
  }, TIMEOUT);
}
