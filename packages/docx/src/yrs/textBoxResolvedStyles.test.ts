import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { createYrsSession, documentToYrs, saveYrsDocx, yrsToDocument, type YrsSession } from './index';

const sessions: YrsSession[] = [];
const KEYS = ['resolvedParagraphStyle', 'resolvedRunStyle'];
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"`;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
  )));
  await preloadOpcWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')
  )));
}, 30_000);
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

async function session(clientId = 97800 + sessions.length): Promise<YrsSession> {
  const result = await createYrsSession({ clientId });
  sessions.push(result);
  return result;
}

function fixture(extraBody = ''): Uint8Array<ArrayBuffer> {
  const paragraph = '<w:p><w:pPr><w:pStyle w:val="Box"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Accent"/></w:rPr><w:t>Styled box</w:t></w:r></w:p>';
  const nested = '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>' + paragraph + '</w:tc></w:tr></w:tbl>' +
    '<w:sdt><w:sdtPr/><w:sdtContent>' + paragraph + '</w:sdtContent></w:sdt>';
  const drawing = `<w:p><w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="1371600"/><wp:wrapNone/><wp:docPr id="1" name="Drawing box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>${paragraph}${nested}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
  const vml = `<w:p><w:r><w:pict><v:shape id="VML box" style="position:absolute;width:216pt;height:108pt" filled="f" stroked="f"><v:textbox><w:txbxContent>${paragraph}${nested}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  add('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Type="${R}/header" Target="header1.xml"/><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`);
  add('word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Box"><w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:line="300"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="30"/></w:rPr></w:style><w:style w:type="character" w:styleId="Accent"><w:rPr><w:b/><w:lang w:val="de-DE"/></w:rPr></w:style><w:style w:type="character" w:styleId="Later"><w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:sz w:val="42"/></w:rPr></w:style><w:style w:type="character" w:styleId="Secret"><w:rPr><w:vanish/></w:rPr></w:style></w:styles>`);
  add('word/document.xml', `<w:document ${NS}><w:body>${extraBody}${drawing}${vml}<w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr></w:body></w:document>`);
  add('word/header1.xml', `<w:hdr ${NS}>${drawing}${vml}</w:hdr>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function updateEncoder() {
  const bytes: number[] = [];
  const uint = (value: number): void => {
    while (value >= 128) {
      bytes.push((value % 128) + 128);
      value = Math.floor(value / 128);
    }
    bytes.push(value);
  };
  const string = (value: string): void => {
    const encoded = new TextEncoder().encode(value);
    uint(encoded.length);
    bytes.push(...encoded);
  };
  return { uint, string, finish: () => Uint8Array.from(bytes) };
}

function mapStringUpdate(client: number, root: string, key: string, value: string): Uint8Array {
  const encoder = updateEncoder();
  for (const value of [1, 1, client, 0, 40, 1]) encoder.uint(value);
  encoder.string(root);
  encoder.string(key);
  encoder.uint(1);
  encoder.uint(119);
  encoder.string(value);
  encoder.uint(0);
  return encoder.finish();
}

function invalidParentUpdate(): Uint8Array {
  const encoder = updateEncoder();
  const client = 111;
  for (const value of [1, 2, client, 0, 40, 1]) encoder.uint(value);
  encoder.string('invalidParent');
  encoder.string('value');
  encoder.uint(1);
  encoder.uint(119);
  encoder.string('scalar');
  for (const value of [4, 0, client, 0]) encoder.uint(value);
  encoder.string('child');
  encoder.uint(0);
  return encoder.finish();
}

function assertAuthored(value: unknown): void {
  if (value instanceof Map) {
    for (const child of value.values()) assertAuthored(child);
  } else if (Array.isArray(value)) {
    value.forEach(assertAuthored);
  } else if (value && typeof value === 'object' && !ArrayBuffer.isView(value) && !(value instanceof ArrayBuffer)) {
    for (const [key, child] of Object.entries(value)) {
      expect(KEYS).not.toContain(key);
      assertAuthored(child);
    }
  }
}

function payloads(session: YrsSession, story: string): string[] {
  return session.storySegments(story).flatMap((segment) =>
    segment.kind === 'embed'
      ? [segment.payload.shapeJson, segment.payload.textBoxJson].filter((value): value is string => typeof value === 'string')
      : []
  );
}

function styledRuns(session: YrsSession, story: string, text = 'Styled box'): Record<string, unknown>[] {
  const runs: Record<string, unknown>[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) value.forEach(visit);
    else if (value && typeof value === 'object') {
      const node = value as Record<string, unknown>;
      if (node.kind === 'text' && node.text === text) runs.push(node);
      Object.values(node).forEach(visit);
    }
  };
  visit(session.yrsBlocksForStory(story, {}));
  return runs;
}

test('shared text-box styles match state-only and shared-bootstrap peers in body and headers', async () => {
  const bytes = fixture();
  const source = await session();
  const host = source.openDocx(bytes, true, { generation: 'text-box-styles' });
  assertAuthored(source.storySegments('body'));
  for (const story of source.storyIds()) {
    const runs = styledRuns(source, story);
    expect(runs.length).toBeGreaterThanOrEqual(6);
    for (const run of runs) expect(run).toMatchObject({ fontFamily: 'Courier New', fontSize: 15, bold: true });
  }
  const repeat = await session(source.clientId);
  repeat.openDocx(bytes, true, { generation: 'text-box-styles' });
  expect(repeat.encodeState()).toEqual(source.encodeState());
  const state = source.encodeState();
  const cold = await session();
  cold.loadState(state);
  expect(cold.materializeDocx()).toBeNull();
  const peer = await session();
  peer.bootstrapPeer(state, source.encodePeerMetadata(), bytes, host);
  for (const other of [cold, peer]) {
    expect(other.encodeState()).toEqual(state);
    for (const story of source.storyIds()) {
      expect(other.yrsBlocksForStory(story, {})).toEqual(source.yrsBlocksForStory(story, {}));
    }
  }
  assertAuthored(source.materializeDocx());
  assertAuthored(peer.materializeDocx());
  for (const other of [source, cold, peer]) {
    const exported = other.exportStructured({ revisionView: 'accepted', includeFormatting: true });
    expect(exported.ok).toBe(true);
    assertAuthored(exported);
    assertAuthored(yrsToDocument(other, source.materializeDocx()!));
  }
  expect((await saveYrsDocx(peer, { updateModifiedDate: false })).bytes)
    .toEqual((await saveYrsDocx(source, { updateModifiedDate: false })).bytes);
});

test('shared styles preserve payload strings and original bytes on both seed paths', async () => {
  const bytes = fixture();
  const source = await session();
  source.openDocx(bytes, true);
  const legacy = await session();
  legacy.openDocx(bytes, false);
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  documentToYrs(legacy, parsed);
  for (const story of legacy.storyIds()) {
    expect(JSON.stringify(legacy.storySegments(story))).not.toContain('resolvedParagraphStyle');
    expect(JSON.stringify(legacy.storySegments(story))).not.toContain('resolvedRunStyle');
  }
  for (const story of source.storyIds()) expect(payloads(source, story)).toEqual(payloads(legacy, story));
  const saved = await saveYrsDocx(source, { updateModifiedDate: false });
  const before = await saveYrsDocx(legacy, { updateModifiedDate: false });
  expect(saved.bytes).toEqual(before.bytes);
  expect(saved.bytes).toEqual(bytes);
  expect(before.bytes).toEqual(bytes);
  const older = await session();
  older.openDocx(bytes, false);
  older.loadState(legacy.encodeState());
  expect(older.yrsBlocksForStory('body', {})).toEqual(legacy.yrsBlocksForStory('body', {}));
  expect(styledRuns(older, 'body').every((run) => run.fontFamily !== 'Courier New')).toBe(true);
  expect((await saveYrsDocx(older, { updateModifiedDate: false })).bytes).toEqual(bytes);
});

test('edited text-box saves and Document projections contain only authored values', async () => {
  const bytes = fixture();
  const source = await session();
  source.openDocx(bytes, true);
  assertAuthored(source.storySegments('body'));
  let offset = 0;
  let edited = false;
  for (const segment of source.storySegments('body')) {
    if (segment.kind === 'text') {
      offset += segment.text.length;
    } else if (segment.kind === 'embed') {
      if (typeof segment.payload.shapeJson === 'string') {
        const shape = JSON.parse(segment.payload.shapeJson);
        const run = shape.textBody.content[0].content[0];
        run.content[0].text = 'Edited box';
        run.formatting = { ...run.formatting, fontSize: 36, fontFamily: { ascii: 'Verdana', hAnsi: 'Verdana' }, bold: false };
        shape.textBody.content[0].content.push({ type: 'run', formatting: { styleId: 'Accent', italic: true }, content: [{ type: 'text', text: 'New box run' }] });
        source.applyRawOps('body', [{ op: 'setEmbedAttr', index: offset, key: 'shapeJson', value: JSON.stringify(shape) }]);
        edited = true;
        break;
      }
      offset += 1;
    } else {
      offset += 1;
    }
  }
  expect(edited).toBe(true);
  for (const run of styledRuns(source, 'body')) expect(run).toMatchObject({ fontFamily: 'Courier New', fontSize: 15, bold: true });
  const editedRuns = styledRuns(source, 'body', 'Edited box');
  expect(editedRuns).toEqual([expect.objectContaining({ fontFamily: 'Verdana', fontSize: 18 })]);
  expect(editedRuns[0]!.bold).not.toBe(true);
  expect(styledRuns(source, 'body', 'New box run')).toEqual([
    expect.objectContaining({ fontFamily: 'Courier New', fontSize: 15, bold: true, italic: true }),
  ]);
  const cold = await session();
  cold.loadState(source.encodeState());
  expect(cold.yrsBlocksForStory('body', {})).toEqual(source.yrsBlocksForStory('body', {}));
  assertAuthored(yrsToDocument(source, source.materializeDocx()!));
  const saved = await saveYrsDocx(source, { updateModifiedDate: false });
  const parts = unzipContainer(saved.bytes);
  const xml = new TextDecoder().decode(parts['word/document.xml']);
  expect(xml).toContain('Edited box');
  expect(xml).toContain('Verdana');
  for (const [name, part] of Object.entries(parts)) {
    if (!name.endsWith('.xml')) continue;
    for (const key of KEYS) expect(new TextDecoder().decode(part)).not.toContain(key);
  }
  assertAuthored(await parseDocx(saved.bytes.buffer as ArrayBuffer, { preloadFonts: false }));
});

for (const styleId of ['Later', 'Secret']) {
  test(`host payload edits resolve an initially unused ${styleId} style before and after saving`, async () => {
    const bytes = fixture();
    const source = await session();
    source.openDocx(bytes, true);
    source.yrsBlocksForStory('body', {});
    let offset = 0;
    let edited = false;
    for (const segment of source.storySegments('body')) {
      if (segment.kind === 'text') offset += segment.text.length;
      else {
        if (segment.kind === 'embed' && typeof segment.payload.shapeJson === 'string') {
          const shape = JSON.parse(segment.payload.shapeJson);
          const run = shape.textBody.content[0].content[0];
          run.content[0].text = 'Reassigned run';
          run.formatting = { styleId };
          source.applyRawOps('body', [{ op: 'setEmbedAttr', index: offset, key: 'shapeJson', value: JSON.stringify(shape) }]);
          edited = true;
          break;
        }
        offset += 1;
      }
    }
    expect(edited).toBe(true);
    const live = styledRuns(source, 'body', 'Reassigned run');
    if (styleId === 'Secret') expect(live).toEqual([]);
    else expect(live).toEqual([expect.objectContaining({ fontFamily: 'Verdana', fontSize: 21 })]);
    const saved = await saveYrsDocx(source, { updateModifiedDate: false });
    const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
    expect(xml).toContain('Reassigned run');
    expect(xml).toContain(`w:rStyle w:val="${styleId}"`);
    const reopened = await session();
    reopened.openDocx(saved.bytes, true);
    expect(styledRuns(reopened, 'body', 'Reassigned run')).toEqual(live);
  });
}

test('a map-only remote update preserves original saved package bytes', async () => {
  const bytes = fixture();
  const source = await session();
  source.openDocx(bytes, true);
  const before = await saveYrsDocx(source, { updateModifiedDate: false });
  const resolver = JSON.stringify({
    enabled: true,
    styles: {},
    doc_defaults: { rPr: { fontFamily: { ascii: 'Verdana', hAnsi: 'Verdana' }, fontSize: 42 } },
    default_paragraph: null,
    default_table: null,
    default_character: null,
    table_paragraph_formatting: null,
  });
  source.applyUpdate(mapStringUpdate(Number.MAX_SAFE_INTEGER, 'shapeStyles', 'resolver', resolver));
  for (const run of styledRuns(source, 'body')) expect(run).toMatchObject({ fontFamily: 'Verdana', fontSize: 21 });
  expect(styledRuns(source, 'body').length).toBeGreaterThan(0);
  const after = await saveYrsDocx(source, { updateModifiedDate: false });
  expect(after.bytes).toEqual(before.bytes);
  expect(after.bytes).toEqual(bytes);
});

test('a chunk integration error still invalidates projections for its committed authored prefix', async () => {
  const bytes = fixture('<w:p><w:r><w:t>Body</w:t></w:r></w:p>');
  const target = await session();
  target.openDocx(bytes, true);
  const peer = await session();
  peer.loadState(target.encodeState());
  const vector = target.encodeStateVector();
  const paraId = peer.paragraphs('body')[0]!.paraId;
  peer.insertText({ story: 'body', paraId, offset: 0 }, 'Committed ');
  const base = yrsToDocument(target, target.materializeDocx()!);
  expect(() => target.applyUpdateChunks([peer.encodeStateAsUpdate(vector), invalidParentUpdate()])).toThrow();
  expect(target.paragraphs('body')[0]!.text).toBe('Committed Body');
  expect(JSON.stringify(yrsToDocument(target, base))).toContain('Committed ');
  const saved = await saveYrsDocx(target, { updateModifiedDate: false });
  expect(new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']).replace(/<[^>]*>/g, '')).toContain('Committed Body');
});
