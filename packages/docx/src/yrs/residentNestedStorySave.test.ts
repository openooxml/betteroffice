import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { BlockContent } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import type { ResidentSaveRecord } from './residentSave';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const CELL = 'body:t0:r0c0';
const OTHER_CELL = 'body:t1:r0c0';
const CONTROL = 'body:sdt0';

const owned: Array<YrsSession | ResidentEngineSession> = [];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
});

afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});

function paragraph(paraId: string, text: string): string {
  return `<w:p w14:paraId="${paraId}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function table(paraId: string, text: string): string {
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${paragraph(paraId, text)}</w:tc></w:tr></w:tbl>`;
}

function contentControl(paraId: string, text: string): string {
  return `<w:sdt><w:sdtPr><w:id w:val="7"/><w:tag w:val="cc"/></w:sdtPr><w:sdtContent>${paragraph(paraId, text)}</w:sdtContent></w:sdt>`;
}

const SOURCE = (() => {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>
</Types>`));
  parts.set('_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/>
</Relationships>`));
  parts.set('word/document.xml', toBytes(`<w:document ${NS}><w:body>${
    table('0000C001', 'First cell') + paragraph('0000B001', 'Between tables') + table('0000C002', 'Second cell') +
    contentControl('0000D001', 'Controlled text') + paragraph('0000B002', 'Closing paragraph')
  }<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
})();

interface Opened {
  session: ResidentEngineSession;
  record: ResidentSaveRecord;
  save(): Promise<Uint8Array>;
}

const opens: Record<string, () => Promise<Opened>> = {
  'a parsed document': async () => {
    const session = await createResidentEngineSession();
    owned.push(session);
    session.openDocument(await parseDocx(SOURCE.slice().buffer, { preloadFonts: false }), {});
    const record: ResidentSaveRecord = { full: false };
    return { session, record, save: async () => new Uint8Array(await session.saveDocument(undefined, [], record)) };
  },
  'a document buffer': async () => {
    const session = await createResidentEngineSession();
    owned.push(session);
    const hostJson = session.openSourceDocx(SOURCE.slice(), {});
    const host = decodeDocxHostJson(hostJson, SOURCE).document;
    const record: ResidentSaveRecord = { full: false };
    return { session, record, save: async () => new Uint8Array(await session.save(SOURCE.slice(), hostJson, host, [], record)) };
  },
};

function at(session: ResidentEngineSession, story: string) {
  expect(session.geometryReader.storyIds()).toContain(story);
  return { story, paraId: session.geometryReader.paragraphs(story)[0]!.paraId, offset: 0 };
}

function edit(session: ResidentEngineSession, story: string, text = 'Edited '): void {
  session.insertText(at(session, story), text);
  session.addUndoBoundary();
}

async function peerEdit(session: ResidentEngineSession, story: string): Promise<void> {
  const peer = await createYrsSession();
  owned.push(peer);
  peer.openDocx(SOURCE.slice(), false);
  peer.loadState(session.encodeState());
  peer.insertText(at(session, story), 'Peer ');
  session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
}

async function storyText(bytes: Uint8Array, story: string): Promise<string[]> {
  const reopened = await createYrsSession({ clientId: 4242 });
  owned.push(reopened);
  reopened.openDocx(bytes, true);
  return reopened.paragraphs(story).map((entry) => entry.text);
}

function tables(record: ResidentSaveRecord): BlockContent[] {
  return record.base!.package.document.content.filter((block) => block.type === 'table');
}

describe.each(Object.keys(opens))('a resident session opened from %s', (from) => {
  describe.each([
    ['table cell', CELL, 'First cell'],
    ['block content control', CONTROL, 'Controlled text'],
  ])('keeps a %s edit', (_, story, text) => {
    test('made after the first save', async () => {
      const opened = await opens[from]!();
      expect(await storyText(await opened.save(), story)).toEqual([text]);
      edit(opened.session, story);
      const second = await opened.save();
      expect(await storyText(second, story)).toEqual([`Edited ${text}`]);
      expect(await storyText(second, OTHER_CELL)).toEqual(['Second cell']);
    }, 30_000);

    test('made on top of an edit the first save wrote', async () => {
      const opened = await opens[from]!();
      edit(opened.session, story, 'One ');
      expect(await storyText(await opened.save(), story)).toEqual([`One ${text}`]);
      edit(opened.session, story, 'Two ');
      expect(await storyText(await opened.save(), story)).toEqual([`Two One ${text}`]);
    }, 30_000);

    test('undone and redone after the first save', async () => {
      const opened = await opens[from]!();
      edit(opened.session, story);
      expect(await storyText(await opened.save(), story)).toEqual([`Edited ${text}`]);
      expect(opened.session.undo()).toBe(true);
      expect(await storyText(await opened.save(), story)).toEqual([text]);
      expect(opened.session.redo()).toBe(true);
      expect(await storyText(await opened.save(), story)).toEqual([`Edited ${text}`]);
    }, 30_000);

    test('applied from a peer update after the first save', async () => {
      const opened = await opens[from]!();
      await opened.save();
      await peerEdit(opened.session, story);
      expect(await storyText(await opened.save(), story)).toEqual([`Peer ${text}`]);
    }, 30_000);
  });
});

test('a parsed-document save reuses the projection of tables nobody edited', async () => {
  const opened = await opens['a parsed document']!();
  await opened.save();
  await opened.save();
  const [first, second] = tables(opened.record);
  edit(opened.session, CELL);
  await opened.save();
  const [firstAfter, secondAfter] = tables(opened.record);
  expect(firstAfter).not.toBe(first);
  expect(secondAfter).toBe(second);
}, 30_000);
