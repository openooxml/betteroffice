import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { createEditSession, preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { DirtyProjectionStories } from './editorSave';
import type { YrsSession } from './index';
import { createdListKey } from './listDefinitionSave';
import { saveResidentDocument, type ResidentSaveRecord } from './residentSave';
import { wrapOpenedEditSession } from './yrsSessionFacade';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const LISTS_NUMBERING = readFileSync(resolve(import.meta.dir, '__fixtures__/seed-parity/lists/word/numbering.xml'), 'utf8');
const CLIENT = 411;

function makeDocx(document: string, numbering: string): Uint8Array {
  const parts = new Map([
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/></Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', toBytes(document)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`)],
    ['word/numbering.xml', toBytes(numbering)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const paragraph = (value: string, numId?: number) =>
  `<w:p>${numId === undefined ? '' : `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`}<w:r><w:t>${value}</w:t></w:r></w:p>`;
const DOCUMENT = `<w:document xmlns:w="${W}"><w:body>${[
  paragraph('Lead'), paragraph('Alpha'), paragraph('Gap'), paragraph('Beta'), paragraph('Tail'),
  paragraph('Decimal source', 1), paragraph('Bullet source', 2),
].join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;

const text = (bytes: Uint8Array | undefined) => (bytes === undefined ? '' : new TextDecoder().decode(bytes));
const elements = (xml: string, name: string) => xml.match(new RegExp(`<w:${name}\\b[\\s\\S]*?</w:${name}>`, 'g')) ?? [];
const attribute = (xml: string | undefined, pattern: RegExp) => {
  const value = xml?.match(pattern)?.[1];
  return value === undefined ? undefined : Number(value);
};
const NUM_ID = /w:numId="(\d+)"/;
const ABSTRACT_ID = /w:abstractNumId="(\d+)"/;
const POINTER = /<w:abstractNumId w:val="(\d+)"/;

function savedNumId(document: string, value: string): number | undefined {
  const saved = (document.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find(entry => entry.includes(`>${value}<`));
  return attribute(saved, /<w:numId w:val="(\d+)"/);
}

const owned: Array<{ destroy(): void }> = [];
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});

/**
 * A raw engine session opened on `source` as the document worker opens it, a facade over it to
 * issue edits through, and a save that goes through `saveResidentDocument` with one record.
 */
function worker(source: Uint8Array) {
  const raw = createEditSession(CLIENT);
  const hostJson = raw.open_docx(source, true);
  const opened = wrapOpenedEditSession(raw, CLIENT, source, hostJson);
  owned.push(opened.session);
  const record: ResidentSaveRecord = { full: false };
  const dirty = new DirtyProjectionStories();
  const save = async (): Promise<Uint8Array> => {
    dirty.add('body');
    return new Uint8Array(await saveResidentDocument(
      raw, CLIENT, new Uint8Array(source.buffer), hostJson, opened.host.document, [], record, dirty
    ));
  };
  return { session: opened.session, save };
}

function createList(target: YrsSession, value: string, kind: 'bullet' | 'numbered'): void {
  const found = target.paragraphs('body').find(entry => entry.text === value);
  expect(found).toBeDefined();
  const paraId = found!.paraId;
  const attrs = target.ensureListDefinition('body', paraId, kind);
  target.setParagraphAttrs({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, attrs);
}

function takeOff(target: YrsSession, value: string): void {
  const paraId = target.paragraphs('body').find(entry => entry.text === value)!.paraId;
  target.setParagraphAttrs({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, {
    numPr: null, listIsBullet: null, listNumFmt: null, listMarker: null, other: { listDef: null },
  });
}

/** The created-list key the paragraph `value` of `target` is on now, or null. */
const liveKey = (target: YrsSession, value: string): string | null =>
  createdListKey(target.paragraphs('body').find(entry => entry.text === value)!.properties) ?? null;

/**
 * That each paragraph of `expected` is saved in `bytes` on the created list of its key, or on no
 * list for null: its numId names one `w:num`, that points at one `w:abstractNum` carrying the key,
 * and paragraphs on different keys share neither id. The numIds read and the carriers of each key.
 */
function consistent(bytes: Uint8Array, expected: Record<string, string | null>) {
  const parts = unzipContainer(bytes);
  const document = text(parts['word/document.xml']);
  const numbering = text(parts['word/numbering.xml']);
  const nums = elements(numbering, 'num');
  const abstracts = elements(numbering, 'abstractNum');
  const numIds: Record<string, number | null> = {};
  const abstractIds: Record<string, number | undefined> = {};
  for (const [value, key] of Object.entries(expected)) {
    const numId = savedNumId(document, value);
    numIds[value] = numId ?? null;
    if (key === null) {
      expect(numId).toBeUndefined();
      continue;
    }
    expect(numId).toBeDefined();
    const matching = nums.filter(num => attribute(num, NUM_ID) === numId);
    expect(matching).toHaveLength(1);
    const pointer = attribute(matching[0], POINTER);
    const definitions = abstracts.filter(abstract => attribute(abstract, ABSTRACT_ID) === pointer);
    expect(definitions).toHaveLength(1);
    expect(definitions[0]).toContain(`<w:name w:val="betteroffice-list:${key}"/>`);
    abstractIds[value] = pointer;
  }
  const listed = Object.entries(expected).filter((entry): entry is [string, string] => entry[1] !== null);
  for (const [left, leftKey] of listed) {
    for (const [right, rightKey] of listed) {
      if (leftKey === rightKey) continue;
      expect(numIds[left]).not.toBe(numIds[right]);
      expect(abstractIds[left]).not.toBe(abstractIds[right]);
    }
  }
  const carriersPerKey = Object.fromEntries([...new Set(listed.map(([, key]) => key))].map(key => [
    key, abstracts.filter(abstract => abstract.includes(`<w:name w:val="betteroffice-list:${key}"/>`)).length,
  ]));
  return { numIds, carriersPerKey };
}

describe('created lists saved again and again through the document worker', () => {
  test('a bullet list saved three times stays on one consistent definition', async () => {
    const { session, save } = worker(makeDocx(DOCUMENT, LISTS_NUMBERING));
    createList(session, 'Alpha', 'bullet');
    const alpha = liveKey(session, 'Alpha');
    expect(alpha).not.toBeNull();
    const ids: Array<number | null> = [];
    const carriersPerKey: Array<Record<string, number>> = [];
    for (let index = 0; index < 3; index += 1) {
      const checked = consistent(await save(), { Alpha: alpha, Beta: null });
      ids.push(checked.numIds.Alpha ?? null);
      carriersPerKey.push(checked.carriersPerKey);
    }
    console.log('WORKER-SAVE-IDS', JSON.stringify({ sequence: 'W1', ids, carriersPerKey }));
  });

  test('a list taken off, another created and the first brought back by undo stay consistent on every save', async () => {
    const { session, save } = worker(makeDocx(DOCUMENT, LISTS_NUMBERING));
    session.beginUndoCapture();
    createList(session, 'Alpha', 'bullet');
    const alpha = liveKey(session, 'Alpha');
    expect(alpha).not.toBeNull();
    const ids: Array<Record<string, number | null>> = [];
    const carriersPerKey: Array<Record<string, number>> = [];
    const check = async (expected: Record<string, string | null>) => {
      const checked = consistent(await save(), expected);
      ids.push(checked.numIds);
      carriersPerKey.push(checked.carriersPerKey);
    };
    await check({ Alpha: alpha, Beta: null });
    session.addUndoBoundary();
    createList(session, 'Beta', 'numbered');
    const beta = liveKey(session, 'Beta');
    expect(beta).not.toBeNull();
    session.addUndoBoundary();
    takeOff(session, 'Alpha');
    session.addUndoBoundary();
    await check({ Alpha: null, Beta: beta });
    const undone = session.undo();
    const after = { Alpha: liveKey(session, 'Alpha'), Beta: liveKey(session, 'Beta') };
    await check(after);
    console.log('WORKER-SAVE-IDS', JSON.stringify({ sequence: 'W2', undone, after, ids, carriersPerKey }));
  });
});
