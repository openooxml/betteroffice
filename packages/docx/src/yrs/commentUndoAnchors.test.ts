import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { documentToYrs } from './documentToYrs';
import { DirtyProjectionStories, hostSaveMetadata, saveEditorDocument } from './editorSave';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import type { DocxCommandEnvelope, DocxCommandInitRequest } from './residentCommandProtocol';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { yrsToDocument } from './yrsToDocument';

const TIMEOUT = 60_000;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const ORIGINAL = 'Head middle tail';
const COVERED = 'middle';
const BODY =
  '<w:p w14:paraId="0A000001" w14:textId="0A000001">' +
  '<w:r><w:t xml:space="preserve">Head </w:t></w:r>' +
  '<w:commentRangeStart w:id="1"/><w:r><w:t>middle</w:t></w:r><w:commentRangeEnd w:id="1"/>' +
  '<w:r><w:commentReference w:id="1"/></w:r>' +
  '<w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>';

function fixture(): Uint8Array {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
    )],
    ['word/_rels/document.xml.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdComments" Type="${REL}/comments" Target="comments.xml"/></Relationships>`
    )],
    ['word/styles.xml', toBytes(
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
      '<w:name w:val="Normal"/></w:style></w:styles>'
    )],
    ['word/comments.xml', toBytes(
      `<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
      '<w:comment w:id="1" w:author="Grace" w:date="2025-03-04T05:06:07Z" w:initials="G">' +
      '<w:p w14:paraId="0D000001" w14:textId="77777777"><w:r><w:t>Note 1</w:t></w:r></w:p></w:comment></w:comments>'
    )],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${W}" xmlns:r="${REL}" xmlns:w14="${W14}"><w:body>${BODY}<w:sectPr/></w:body></w:document>`
    )],
  ])));
}

type Reader = Pick<YrsSession, 'paragraphs' | 'storySegments' | 'listComments' | 'resolveComment'>;

/** The first body paragraph: its id, text, each character's story offset, and its length to the pilcrow. */
function layoutOf(reader: Reader) {
  const { paraId, text } = reader.paragraphs('body')[0]!;
  const positions: number[] = [];
  let length = 0;
  for (const segment of reader.storySegments('body')) {
    if (segment.kind === 'pilcrow') break;
    if (segment.kind === 'text') {
      for (let at = 0; at < segment.text.length; at += 1) positions.push(length + at);
      length += segment.text.length;
    } else length += 1;
  }
  return { paraId, text, positions, length };
}

function storyText(reader: Reader, story: string, start: number, end: number): string {
  let offset = 0;
  let text = '';
  for (const segment of reader.storySegments(story)) {
    if (segment.kind === 'text') {
      text += segment.text.slice(Math.max(0, start - offset), Math.max(0, end - offset));
      offset += segment.text.length;
    } else offset += 1;
  }
  return text;
}

function liveRangeText(reader: Reader): string {
  const comments = reader.listComments();
  expect(comments).toHaveLength(1);
  return reader.resolveComment(comments[0]!.id).map(anchor => storyText(reader, anchor.story, anchor.start, anchor.end)).join(' + ');
}

/** The saved document's text between comment 1's markers, and its markers in order. */
function savedRange(bytes: Uint8Array): { text: string | null; markers: string[] } {
  const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  let text = '';
  let start: number | null = null;
  let end: number | null = null;
  const markers: string[] = [];
  for (const match of xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:comment(RangeStart|RangeEnd|Reference)\b[^>]*?\bw:id="(\d+)"/g)) {
    if (match[1] !== undefined) {
      text += match[1];
      continue;
    }
    if (match[3] !== '1') continue;
    markers.push(match[2]!);
    if (match[2] === 'RangeStart') start = text.length;
    if (match[2] === 'RangeEnd') end = text.length;
  }
  return { text: start !== null && end !== null ? text.slice(start, end) : null, markers };
}

interface Route {
  reader: Reader;
  deleteRange(start: number, end: number): Promise<void>;
  undo(): Promise<void>;
  typeAtEnd(): Promise<void>;
  save(): Promise<Uint8Array>;
}

const owned: Array<{ destroy(): void }> = [];
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
let clientId = 91_700;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
  startWorker = await residentWorkerFactory();
}, 30_000);
afterEach(() => { owned.splice(0).forEach(entry => entry.destroy()); });

async function openWorker(bytes: Uint8Array): Promise<Route> {
  const worker = startWorker((clientId += 1));
  const client = new ResidentEngineWorkerClient(worker);
  owned.push(client);
  const { hostJson } = await client.open(bytes.slice(), { generation: 'comment-undo' });
  const host = decodeDocxHostJson(hostJson, bytes.slice()).document;
  const engine = worker.sessions[0]!;
  const reader = engine as unknown as Reader;
  const paraId = reader.paragraphs('body')[0]!.paraId;
  const point = { story: 'body', paraId, offset: 0 };
  const fonts = [new Uint8Array(readFileSync(resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')))];
  const fontId = engine.registerFont(fonts[0]!);
  const measurement = { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [fontId] } };
  const init: DocxCommandInitRequest = {
    session: 'comment-undo', incarnation: 1, epoch: 1, readOnly: false,
    selection: { anchor: point, focus: point },
    fonts,
    layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] }, measurement, renderEnv: {} }),
    layoutWithRegions: true,
    layoutRevision: 1,
    extras: JSON.stringify({ fontChains: { ...measurement.fontChains, 'calibri|0|0': [fontId] } }),
  };
  let version = (await client.commandInit(init)).version;
  let seq = 0;
  const send = async (kind: string, payload: DocxCommandEnvelope['payload']) => {
    const ack = await client.edit({ session: 'comment-undo', incarnation: 1, seq: (seq += 1), basis: { epoch: 1, version }, kind, payload });
    version = ack.version;
    return ack.outcomes[0]?.status;
  };
  const select = (start: number, end: number) => send('select', {
    selection: { anchor: { story: 'body', paraId, offset: start }, focus: { story: 'body', paraId, offset: end } },
  });
  return {
    reader,
    async deleteRange(start, end) {
      await select(start, end);
      expect(await send('deleteBackward', {})).toBe('applied');
    },
    async undo() {
      expect(await send('undo', {})).toBe('applied');
    },
    async typeAtEnd() {
      const { length } = layoutOf(reader);
      await select(length, length);
      expect(await send('insertText', { text: 'X' })).toBe('applied');
    },
    async save() {
      const saved = await client.save({ comments: host.package.document.comments ?? [], host: hostSaveMetadata(host) });
      return new Uint8Array(saved.bytes);
    },
  };
}

async function openTypeScript(bytes: Uint8Array): Promise<Route> {
  const session = await createYrsSession({ clientId: (clientId += 1) });
  owned.push(session);
  const parsed = await parseDocx(bytes.slice().buffer as ArrayBuffer, { preloadFonts: false });
  documentToYrs(session, parsed);
  const base: Document = { ...parsed, originalBuffer: parsed.originalBuffer ?? (bytes.slice().buffer as ArrayBuffer) };
  const comments = parsed.package.document.comments ?? [];
  const paraId = session.paragraphs('body')[0]!.paraId;
  return {
    reader: session,
    async deleteRange(start, end) {
      session.beginUndoCapture();
      session.addUndoBoundary();
      session.deleteRange({ story: 'body', start: { paraId, offset: start }, end: { paraId, offset: end } });
      session.addUndoBoundary();
    },
    async undo() {
      expect(session.undo()).toBe(true);
      session.addUndoBoundary();
    },
    async typeAtEnd() {
      session.beginUndoCapture();
      session.insertText({ story: 'body', paraId, offset: layoutOf(session).length }, 'X');
      session.addUndoBoundary();
    },
    async save() {
      const dirty = new DirtyProjectionStories();
      dirty.add('body');
      const projected = yrsToDocument(session, base, dirty.projectionOptions());
      return new Uint8Array(await saveEditorDocument(session, projected, comments));
    },
  };
}

const ROUTES: Array<[string, (bytes: Uint8Array) => Promise<Route>]> = [
  ['the worker route', openWorker],
  ['the TypeScript loader route', openTypeScript],
];

const characters = (needle: string): [number, number] => {
  const at = ORIGINAL.indexOf(needle);
  return [at, at + needle.length];
};

const CASES: Array<{ name: string; characters: [number, number] | null; typeAfter?: boolean }> = [
  { name: "deleting 'd mid' across the range start", characters: characters('d mid') },
  { name: "deleting 'dle t' across the range end", characters: characters('dle t') },
  { name: "deleting 'iddl' inside the range", characters: characters('iddl') },
  { name: 'deleting the whole paragraph text', characters: null },
  { name: "deleting 'd mid' and typing at the paragraph end after the undo", characters: characters('d mid'), typeAfter: true },
];

describe("a comment over 'middle' in 'Head middle tail' after a deletion is undone", () => {
  for (const [routeName, open] of ROUTES) {
    for (const { name, characters: chars, typeAfter } of CASES) {
      it(`${routeName}: ${name} keeps the live and the saved range on 'middle'`, async () => {
        const route = await open(fixture());
        const before = layoutOf(route.reader);
        expect(before.text).toBe(ORIGINAL);
        expect(liveRangeText(route.reader)).toBe(COVERED);
        const range = chars
          ? { start: before.positions[chars[0]]!, end: before.positions[chars[1] - 1]! + 1 }
          : { start: 0, end: before.length };
        await route.deleteRange(range.start, range.end);
        expect(layoutOf(route.reader).text).not.toBe(ORIGINAL);
        await route.undo();
        expect(layoutOf(route.reader).text).toBe(ORIGINAL);
        if (typeAfter) {
          await route.typeAtEnd();
          expect(layoutOf(route.reader).text).toBe(`${ORIGINAL}X`);
        }
        expect(liveRangeText(route.reader)).toBe(COVERED);
        const saved = savedRange(await route.save());
        expect(saved).toEqual({ text: COVERED, markers: ['RangeStart', 'RangeEnd', 'Reference'] });
      }, TIMEOUT);
    }
  }
});
