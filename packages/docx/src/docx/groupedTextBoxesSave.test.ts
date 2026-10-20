import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '.';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, documentToYrs, saveYrsDocx, yrsToDocument, type YrsSession } from '../yrs';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { unzipContainer } from './wasm';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const sessions: YrsSession[] = [];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

function fixture(): Uint8Array {
  const box = (id: number, text: string) =>
    `<wps:wsp><wps:cNvPr id="${id}"/><wps:cNvSpPr txBox="1"/><wps:spPr>` +
    `<a:xfrm><a:off x="${id * 95250}" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>` +
    '<a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>' +
    `<w:p w14:paraId="0000000${id}"><w:r><w:t>${text}</w:t></w:r></w:p>` +
    '</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>';
  const group = '<w:drawing><wp:inline><wp:extent cx="3810000" cy="1905000"/>' +
    '<wp:docPr id="1" name="Group"/><a:graphic><a:graphicData><wpg:wgp>' +
    '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="1905000" cy="952500"/></a:xfrm></wpg:grpSpPr>' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="4" name="Sibling"/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="image"/></pic:blipFill><pic:spPr><a:xfrm>' +
    '<a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm></pic:spPr></pic:pic>' +
    box(2, 'Alpha') + box(3, 'Beta') +
    '</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>';
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="image" Type="${R}/image" Target="media/sibling.png"/></Relationships>`);
  add('word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" ` +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>' +
    `<w:p w14:paraId="00000001"><w:r><w:t>Host</w:t></w:r><w:r>${group}</w:r></w:p>` +
    '<w:p w14:paraId="00000005"><w:r><w:t>Other</w:t></w:r></w:p><w:sectPr/></w:body></w:document>');
  parts.set('word/media/sibling.png', new Uint8Array([1, 2, 3, 4]));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: sessions.length + 1 });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

function xml(bytes: Uint8Array): string {
  return new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
}

function assertGroup(document: Document): void {
  const host = document.package.document.content[0]!;
  if (host.type !== 'paragraph') throw new Error('host paragraph');
  const shapes = host.content.flatMap((inline) => inline.type === 'run'
    ? inline.content.filter((content) => content.type === 'shape') : []);
  expect(shapes).toHaveLength(1);
  const group = shapes[0]!.shape;
  expect(group.children).toHaveLength(2);
  expect(JSON.stringify(group.children![0]!.textBody)).toContain('Alpha');
  expect(JSON.stringify(group.children![1]!.textBody)).toContain('Beta');
  expect(group.scene!.root!.children!.map((child) => child.kind)).toEqual(['picture', 'shape', 'shape']);
}

test('grouped stories stay display-only through seed, save and reopen', async () => {
  const bytes = fixture();
  const session = await open(bytes);
  expect(session.storyIds()).toEqual(['body']);
  expect(session.searchText('Alpha')).toEqual([]);
  expect(session.searchText('Beta')).toEqual([]);
  expect(session.paragraphs('body').map(({ text }) => text).join('')).not.toContain('Alpha');
  const unedited = await saveYrsDocx(session, { updateModifiedDate: false });
  expect(unedited.bytes).toEqual(bytes);
  assertGroup(await parseDocx(unedited.bytes.buffer as ArrayBuffer, { preloadFonts: false }));

  const other = session.paragraphs('body')[1]!;
  session.insertText({ story: 'body', paraId: other.paraId, offset: 0 }, 'Edited ');
  const saved = await saveYrsDocx(session, { updateModifiedDate: false });
  expect(xml(saved.bytes)).toContain('Edited Other');
  expect(xml(saved.bytes).match(/<wpg:wgp>[\s\S]*?<\/wpg:wgp>/)?.[0])
    .toBe(xml(bytes).match(/<wpg:wgp>[\s\S]*?<\/wpg:wgp>/)?.[0]);
  const reopened = await open(saved.bytes);
  const document = await parseDocx(saved.bytes.buffer as ArrayBuffer, { preloadFonts: false });
  assertGroup(yrsToDocument(reopened, document));
  expect(unzipContainer(saved.bytes)['word/media/sibling.png'])
    .toEqual(unzipContainer(bytes)['word/media/sibling.png']);
});

function placeholderStoryContent(value: unknown): void {
  if (Array.isArray(value)) {
    for (const child of value) placeholderStoryContent(child);
  } else if (value && typeof value === 'object') {
    const object = value as Record<string, unknown>;
    if (object.textBody && typeof object.textBody === 'object') {
      const body = object.textBody as Record<string, unknown>;
      if (Array.isArray(body.content)) {
        body.content = body.content.map(() => ({ type: 'paragraph', formatting: {}, content: [] }));
      }
    }
    for (const child of Object.values(object)) placeholderStoryContent(child);
  }
}

test('host edits serialize the same with enriched and placeholder grouped stories', async () => {
  const bytes = fixture();
  const enriched = await open(bytes);
  const legacy = await createYrsSession({ clientId: 100 });
  sessions.push(legacy);
  legacy.openDocx(bytes, false);
  const baseline = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  placeholderStoryContent(baseline.package.document.content);
  documentToYrs(legacy, baseline);
  const outputs: Uint8Array[] = [];
  for (const session of [enriched, legacy]) {
    const host = session.paragraphs('body')[0]!;
    session.insertText({ story: 'body', paraId: host.paraId, offset: 0 }, 'Edited ');
    outputs.push((await saveYrsDocx(session, { updateModifiedDate: false })).bytes);
  }
  expect(xml(outputs[0]!)).toBe(xml(outputs[1]!));
  const reopened = await open(outputs[0]!);
  expect(reopened.paragraphs('body')[0]!.text).toContain('Edited Host');
  expect(xml(outputs[0]!)).not.toContain('<wpg:wgp>');
  expect(xml(outputs[0]!)).toContain('<wps:wsp>');
});
