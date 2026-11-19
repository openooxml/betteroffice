import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { saveEditorDocument } from './editorSave';
import { createYrsSession, saveYrsDocx, yrsToDocument, type YrsSession } from './index';
import { savedListDefinitionIds } from './listDefinitionSave';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const FIXTURE = resolve(import.meta.dir, '__fixtures__/seed-parity/lists/word');
const LISTS_DOCUMENT = readFileSync(resolve(FIXTURE, 'document.xml'), 'utf8');
const LISTS_NUMBERING = readFileSync(resolve(FIXTURE, 'numbering.xml'), 'utf8');
const STYLES = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style></w:styles>`;
const LEVELS = [0, 1, 2, 3, 4, 5, 6, 7, 8];

function makeDocx(document: string, numbering?: string): Uint8Array {
  const overrides = [
    `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>`,
    `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>`,
    ...(numbering === undefined ? [] : [`<Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/>`]),
  ];
  const relationships = [
    `<Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/>`,
    ...(numbering === undefined ? [] : [`<Relationship Id="rId2" Type="${R}/numbering" Target="numbering.xml"/>`]),
  ];
  const parts = new Map([
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.join('')}</Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', toBytes(document)],
    ['word/styles.xml', toBytes(STYLES)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join('')}</Relationships>`)],
  ]);
  if (numbering !== undefined) parts.set('word/numbering.xml', toBytes(numbering));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const paragraph = (value: string, numId?: number) =>
  `<w:p>${numId === undefined ? '' : `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`}<w:r><w:t>${value}</w:t></w:r></w:p>`;
const documentOf = (...paragraphs: string[]) =>
  `<w:document xmlns:w="${W}"><w:body>${paragraphs.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;

const text = (bytes: Uint8Array | undefined) => (bytes === undefined ? undefined : new TextDecoder().decode(bytes));
const elements = (xml: string, name: string) => xml.match(new RegExp(`<w:${name}\\b[\\s\\S]*?</w:${name}>`, 'g')) ?? [];
const attribute = (xml: string | undefined, pattern: RegExp) => {
  const value = xml?.match(pattern)?.[1];
  return value === undefined ? undefined : Number(value);
};
const NUM_ID = /w:numId="(\d+)"/;
const ABSTRACT_ID = /w:abstractNumId="(\d+)"/;
const POINTER = /<w:abstractNumId w:val="(\d+)"/;
const levels = (abstractNum: string) => [...abstractNum.matchAll(/<w:lvl\b[^>]*\bw:ilvl="(\d+)"/g)].map(match => Number(match[1]));
const formats = (abstractNum: string) => [...abstractNum.matchAll(/<w:numFmt w:val="([^"]*)"/g)].map(match => match[1]);

function savedNumId(document: string, value: string): number | undefined {
  const saved = (document.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find(entry => entry.includes(`>${value}<`));
  return attribute(saved, /<w:numId w:val="(\d+)"/);
}

type Save = (session: YrsSession) => Promise<Uint8Array>;
const SAVES: Record<string, Save> = {
  saveYrsDocx: async session => (await saveYrsDocx(session, { updateModifiedDate: false })).bytes,
  editor: async (session) => {
    const projected = yrsToDocument(session, session.materializeDocx()!);
    return new Uint8Array(await saveEditorDocument(session, projected, projected.package.document.comments ?? []));
  },
};

const open: YrsSession[] = [];
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { open.splice(0).forEach(session => session.destroy()); });

async function seeded(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(bytes);
  return session;
}

/** Puts the paragraph `value` on a list of `kind` as a command would, writing the attributes the engine hands back. */
function createList(session: YrsSession, value: string, kind: 'bullet' | 'numbered'): void {
  const target = session.paragraphs('body').find(entry => entry.text === value);
  expect(target).toBeDefined();
  const paraId = target!.paraId;
  const attrs = session.ensureListDefinition('body', paraId, kind);
  session.setParagraphAttrs({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, attrs);
}

describe('savedListDefinitionIds', () => {
  const source = { numIds: [1, 2, 3], abstractNumIds: [0, 1] };

  test('created definitions take distinct ids above the source ids', () => {
    const ids = savedListDefinitionIds(source, ['401:7', '402:3']);
    expect(ids.size).toBe(2);
    const values = [...ids.values()];
    expect(new Set(values.map(id => id.numId)).size).toBe(2);
    expect(new Set(values.map(id => id.abstractNumId)).size).toBe(2);
    for (const id of values) {
      expect(id.numId).toBeGreaterThan(3);
      expect(id.abstractNumId).toBeGreaterThan(1);
    }
  });

  test('the same definitions take the same ids whatever order they are listed in', () => {
    const ids = savedListDefinitionIds(source, ['401:7', '402:3']);
    expect(ids.size).toBe(2);
    expect(savedListDefinitionIds(source, ['402:3', '401:7'])).toEqual(ids);
  });

  test('a definition saved before keeps its ids when another one is added', () => {
    const before = savedListDefinitionIds(source, ['402:3']);
    expect(before.size).toBe(1);
    const after = savedListDefinitionIds(source, ['401:7', '402:3'], before);
    expect(after.get('402:3')).toEqual(before.get('402:3'));
    const added = after.get('401:7');
    expect(added?.numId).toBeGreaterThan(3);
    expect(added?.abstractNumId).toBeGreaterThan(1);
    expect(added?.numId).not.toBe(before.get('402:3')?.numId);
    expect(added?.abstractNumId).not.toBe(before.get('402:3')?.abstractNumId);
  });
});

for (const [saveName, save] of Object.entries(SAVES)) {
  test(`${saveName}: created lists take their own ids above the source's, the same on every save`, async () => {
    const session = await seeded(makeDocx(documentOf(
      paragraph('Lead'), paragraph('Alpha'), paragraph('Gap'), paragraph('Beta'), paragraph('Tail'),
      paragraph('Decimal source', 1), paragraph('Bullet source', 2),
    ), LISTS_NUMBERING));
    createList(session, 'Alpha', 'bullet');
    createList(session, 'Beta', 'numbered');
    const first = unzipContainer(await save(session));
    const second = unzipContainer(await save(session));
    const document = text(first['word/document.xml']) ?? '';
    const numbering = text(first['word/numbering.xml']) ?? '';
    const alpha = savedNumId(document, 'Alpha');
    const beta = savedNumId(document, 'Beta');
    expect(alpha).toBeGreaterThan(3);
    expect(beta).toBeGreaterThan(3);
    expect(alpha).not.toBe(beta);
    const created = elements(numbering, 'num').filter(num => (attribute(num, NUM_ID) ?? 0) > 3);
    expect(new Set(created.map(num => attribute(num, NUM_ID)))).toEqual(new Set([alpha, beta]));
    expect(new Set(created.map(num => attribute(num, POINTER))).size).toBe(2);
    const abstracts = elements(numbering, 'abstractNum');
    const kinds = new Map(created.map(num => {
      const abstractId = attribute(num, POINTER);
      const abstract = abstracts.find(entry => attribute(entry, ABSTRACT_ID) === abstractId) ?? '';
      expect(abstractId).toBeGreaterThan(1);
      expect(levels(abstract)).toEqual(LEVELS);
      return [attribute(num, NUM_ID), formats(abstract)[0]] as const;
    }));
    expect(kinds.get(alpha)).toBe('bullet');
    expect(kinds.get(beta)).toBe('decimal');
    expect(text(second['word/document.xml'])).toBe(document);
    expect(text(second['word/numbering.xml'])).toBe(numbering);
  });

  test(`${saveName}: a created list is appended to the numbering part the document has`, async () => {
    const session = await seeded(makeDocx(LISTS_DOCUMENT, LISTS_NUMBERING));
    createList(session, 'Outline', 'bullet');
    const parts = unzipContainer(await save(session));
    const numbering = text(parts['word/numbering.xml']) ?? '';
    const sourceAbstracts = elements(LISTS_NUMBERING, 'abstractNum');
    const sourceNums = elements(LISTS_NUMBERING, 'num');
    const abstracts = elements(numbering, 'abstractNum');
    const nums = elements(numbering, 'num');
    expect(abstracts).toHaveLength(sourceAbstracts.length + 1);
    expect(nums).toHaveLength(sourceNums.length + 1);
    expect(abstracts.slice(0, -1)).toEqual(sourceAbstracts);
    expect(nums.slice(0, -1)).toEqual(sourceNums);
    expect(numbering.startsWith(LISTS_NUMBERING.slice(0, LISTS_NUMBERING.indexOf('<w:abstractNum ')))).toBe(true);
    const created = abstracts.at(-1) ?? '';
    expect(numbering.indexOf(created)).toBeLessThan(numbering.indexOf(nums[0] ?? ''));
    const abstractId = attribute(created, ABSTRACT_ID);
    expect(abstractId).toBeGreaterThan(1);
    expect(levels(created)).toEqual(LEVELS);
    expect(formats(created)).toEqual(LEVELS.map(() => 'bullet'));
    const num = nums.at(-1);
    expect(attribute(num, POINTER)).toBe(abstractId);
    const numId = attribute(num, NUM_ID);
    expect(numId).toBeGreaterThan(3);
    expect(savedNumId(text(parts['word/document.xml']) ?? '', 'Outline')).toBe(numId);
  });

  test(`${saveName}: a created list in a document without numbering adds the part, its relationship and its content type`, async () => {
    const session = await seeded(makeDocx(documentOf(paragraph('First'), paragraph('Second'))));
    createList(session, 'Second', 'numbered');
    const parts = unzipContainer(await save(session));
    const numbering = text(parts['word/numbering.xml']);
    expect(numbering).toBeDefined();
    const abstracts = elements(numbering ?? '', 'abstractNum');
    const nums = elements(numbering ?? '', 'num');
    expect(abstracts).toHaveLength(1);
    expect(nums).toHaveLength(1);
    expect(levels(abstracts[0] ?? '')).toEqual(LEVELS);
    expect(formats(abstracts[0] ?? '')[0]).toBe('decimal');
    expect(attribute(nums[0], POINTER)).toBe(attribute(abstracts[0], ABSTRACT_ID));
    const relationships = text(parts['word/_rels/document.xml.rels']) ?? '';
    const relationship = (relationships.match(/<Relationship\b[^>]*>/g) ?? []).find(entry => entry.includes(`Type="${R}/numbering"`));
    expect(relationship).toMatch(/Target="(\/word\/)?numbering\.xml"/);
    const types = text(parts['[Content_Types].xml']) ?? '';
    const override = (types.match(/<Override\b[^>]*>/g) ?? []).find(entry => entry.includes('PartName="/word/numbering.xml"'));
    expect(override).toContain(`ContentType="${OFFICE}.numbering+xml"`);
    expect(savedNumId(text(parts['word/document.xml']) ?? '', 'Second')).toBe(attribute(nums[0], NUM_ID));
  });

  test(`${saveName}: a document where no list was created keeps its numbering, relationships and content types byte for byte`, async () => {
    const bytes = makeDocx(LISTS_DOCUMENT, LISTS_NUMBERING);
    const source = unzipContainer(bytes);
    const session = await seeded(bytes);
    const parts = unzipContainer(await save(session));
    for (const part of ['word/document.xml', 'word/numbering.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml']) {
      expect(text(parts[part])).toBe(text(source[part]));
    }
  });
}
