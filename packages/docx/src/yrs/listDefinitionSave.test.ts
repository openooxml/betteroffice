import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { saveEditorDocument } from './editorSave';
import { createYrsSession, saveYrsDocx, yrsToDocument, type YrsSession } from './index';
import {
  numberingWithListDefinitions, resolvedRelationshipTarget, savedListDefinitionIds, writeListDefinitionParts,
  type DocxSavedListDefinition,
} from './listDefinitionSave';
import { captureSessionSave, writeSessionSave } from './saveYrsDocx';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const FIXTURE = resolve(import.meta.dir, '__fixtures__/seed-parity/lists/word');
const LISTS_DOCUMENT = readFileSync(resolve(FIXTURE, 'document.xml'), 'utf8');
const LISTS_NUMBERING = readFileSync(resolve(FIXTURE, 'numbering.xml'), 'utf8');
const STYLES = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style></w:styles>`;
const LEVELS = [0, 1, 2, 3, 4, 5, 6, 7, 8];

function makeDocx(document: string, numbering?: string, styles = STYLES): Uint8Array {
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
    ['word/styles.xml', toBytes(styles)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join('')}</Relationships>`)],
  ]);
  if (numbering !== undefined) parts.set('word/numbering.xml', toBytes(numbering));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const PACKAGE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** `bytes` with the parts `replaced` written over. */
function repackaged(bytes: Uint8Array, replaced: Record<string, string>): Uint8Array {
  const parts = new Map(Object.entries(unzipContainer(bytes)));
  for (const [path, xml] of Object.entries(replaced)) parts.set(path, toBytes(xml));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** `promise`'s value, or the error it rejects with. */
async function settled<T>(promise: Promise<T>): Promise<{ value?: T; error?: unknown }> {
  try {
    return { value: await promise };
  } catch (error) {
    return { error };
  }
}

const relationshipEntries = (xml: string) => xml.match(/<(?:[\w.-]+:)?Relationship\b[^>]*>/g) ?? [];
const numberingRelationships = (xml: string) =>
  relationshipEntries(xml).filter(entry => /\sType\s*=\s*["'][^"']*\/numbering["']/.test(entry));

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

  test(`${saveName}: a created list takes a numId above one a source paragraph names without a definition`, async () => {
    const session = await seeded(makeDocx(documentOf(paragraph('Dangling', 1), paragraph('Separator'), paragraph('Fresh'))));
    createList(session, 'Fresh', 'numbered');
    const parts = unzipContainer(await save(session));
    const document = text(parts['word/document.xml']) ?? '';
    const fresh = savedNumId(document, 'Fresh');
    expect(fresh).toBeGreaterThan(1);
    expect(savedNumId(document, 'Dangling')).toBe(1);
    expect(elements(text(parts['word/numbering.xml']) ?? '', 'num').map(num => attribute(num, NUM_ID))).toEqual([fresh]);
  });

  test(`${saveName}: a created list takes a numId above one a style names without a definition`, async () => {
    const styles = STYLES.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="Listed"><w:name w:val="Listed"/><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style></w:styles>');
    const session = await seeded(makeDocx(documentOf(paragraph('First'), paragraph('Second')), undefined, styles));
    createList(session, 'Second', 'bullet');
    const parts = unzipContainer(await save(session));
    expect(savedNumId(text(parts['word/document.xml']) ?? '', 'Second')).toBeGreaterThan(5);
  });

  test(`${saveName}: a created list takes a numId above one a paragraph of the session names without a created list`, async () => {
    const session = await seeded(makeDocx(documentOf(paragraph('Bare'), paragraph('Separator'), paragraph('Fresh'))));
    const bare = session.paragraphs('body').find(entry => entry.text === 'Bare')!.paraId;
    session.setParagraphAttrs({ story: 'body', start: { paraId: bare, offset: 0 }, end: { paraId: bare, offset: 0 } }, {
      numPr: { numId: 1, ilvl: 0 },
    });
    createList(session, 'Fresh', 'bullet');
    const document = text(unzipContainer(await save(session))['word/document.xml']) ?? '';
    expect(savedNumId(document, 'Fresh')).toBeGreaterThan(1);
    expect(savedNumId(document, 'Bare')).toBe(1);
  });

  test(`${saveName}: relationships written with a prefix gain one numbering relationship in that prefix`, async () => {
    const session = await seeded(repackaged(makeDocx(documentOf(paragraph('First'), paragraph('Second'))), {
      'word/_rels/document.xml.rels': `<r:Relationships xmlns:r="${PACKAGE_RELATIONSHIPS}"><r:Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/></r:Relationships>`,
    }));
    createList(session, 'Second', 'numbered');
    const saved = await settled(save(session));
    expect(saved.error).toBeUndefined();
    const parts = unzipContainer(saved.value!);
    const relationships = text(parts['word/_rels/document.xml.rels']) ?? '';
    const numbering = numberingRelationships(relationships);
    expect(numbering).toHaveLength(1);
    expect(numbering[0]).toMatch(/^<r:Relationship\s/);
    expect(relationships).not.toMatch(/<\/?Relationships?\b/);
    expect(elements(text(parts['word/numbering.xml']) ?? '', 'num')).toHaveLength(1);
    expect(text(parts['[Content_Types].xml'])).toContain('PartName="/word/numbering.xml"');
  });

  test(`${saveName}: relationships written with a prefix that name the numbering part keep one numbering relationship`, async () => {
    const session = await seeded(repackaged(makeDocx(documentOf(paragraph('First'), paragraph('Second')), LISTS_NUMBERING), {
      'word/_rels/document.xml.rels': `<r:Relationships xmlns:r="${PACKAGE_RELATIONSHIPS}"><r:Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><r:Relationship Id="rId2" Type="${R}/numbering" Target="numbering.xml"/></r:Relationships>`,
    }));
    createList(session, 'Second', 'numbered');
    const saved = await settled(save(session));
    expect(saved.error).toBeUndefined();
    const parts = unzipContainer(saved.value!);
    expect(numberingRelationships(text(parts['word/_rels/document.xml.rels']) ?? '')).toHaveLength(1);
    const numbering = text(parts['word/numbering.xml']) ?? '';
    expect(elements(numbering, 'abstractNum')).toHaveLength(elements(LISTS_NUMBERING, 'abstractNum').length + 1);
    expect(elements(numbering, 'num')).toHaveLength(elements(LISTS_NUMBERING, 'num').length + 1);
  });

  test(`${saveName}: a relationship id written in single quotes is not given again`, async () => {
    const session = await seeded(repackaged(makeDocx(documentOf(paragraph('First'), paragraph('Second'))), {
      'word/_rels/document.xml.rels': `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id='rId1' Type="${R}/styles" Target="styles.xml"/></Relationships>`,
    }));
    createList(session, 'Second', 'numbered');
    const relationships = text(unzipContainer(await save(session))['word/_rels/document.xml.rels']) ?? '';
    const ids = relationshipEntries(relationships).map(entry => /\sId\s*=\s*["']([^"']*)["']/.exec(entry)?.[1]);
    expect(ids).toHaveLength(2);
    expect(new Set(ids).size).toBe(2);
    expect(numberingRelationships(relationships)).toHaveLength(1);
  });

  test(`${saveName}: a numbering relationship whose type is written in single quotes is not added again`, async () => {
    const session = await seeded(repackaged(makeDocx(documentOf(paragraph('First'), paragraph('Second')), LISTS_NUMBERING), {
      'word/_rels/document.xml.rels': `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type='${R}/numbering' Target="numbering.xml"/></Relationships>`,
    }));
    createList(session, 'Second', 'numbered');
    const parts = unzipContainer(await save(session));
    expect(numberingRelationships(text(parts['word/_rels/document.xml.rels']) ?? '')).toHaveLength(1);
    expect(elements(text(parts['word/numbering.xml']) ?? '', 'num')).toHaveLength(elements(LISTS_NUMBERING, 'num').length + 1);
  });

  test(`${saveName}: a document where no list was created keeps its numbering, relationships and content types byte for byte`, async () => {
    const bytes = makeDocx(LISTS_DOCUMENT, LISTS_NUMBERING);
    const source = unzipContainer(bytes);
    const session = await seeded(bytes);
    const parts = unzipContainer(await save(session));
    for (const part of ['word/document.xml', 'word/numbering.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml']) {
      expect(parts[part]).toEqual(source[part]);
    }
  });
}

/** A session opened from `bytes` by one loader, and how it saves. */
type Loaded = { session: YrsSession; save: () => Promise<Uint8Array> };
const LOADERS: Record<string, (bytes: Uint8Array) => Promise<Loaded>> = {
  'engine seed, saveYrsDocx': async (bytes) => {
    const session = await seeded(bytes);
    return { session, save: () => SAVES.saveYrsDocx!(session) };
  },
  'engine seed, editor save': async (bytes) => {
    const session = await seeded(bytes);
    return { session, save: () => SAVES.editor!(session) };
  },
  'TypeScript loader, session save': async (bytes) => {
    const source = bytes.slice().buffer as ArrayBuffer;
    const parsed = await parseDocx(source, { preloadFonts: false });
    const session = await createYrsSession();
    open.push(session);
    documentToYrs(session, parsed);
    return {
      session,
      save: async () => (await writeSessionSave(
        session, yrsToDocument(session, parsed), captureSessionSave(session), parsed.originalBuffer ?? source, {}, () => false
      )).bytes,
    };
  },
};

for (const [loader, load] of Object.entries(LOADERS)) {
  for (const [spelling, numPr] of [
    ['with its level', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'],
    ['without a level', '<w:numPr><w:numId w:val="1"/></w:numPr>'],
  ] as const) {
    test(`${loader}: bullets on a paragraph numbered by its style ${spelling} save a direct created list and leave the next on the style's list`, async () => {
      const styles = STYLES.replace(
        '</w:styles>',
        `<w:style w:type="paragraph" w:styleId="Listed"><w:name w:val="Listed"/><w:pPr>${numPr}</w:pPr></w:style></w:styles>`
      );
      const styled = (value: string) => `<w:p><w:pPr><w:pStyle w:val="Listed"/></w:pPr><w:r><w:t>${value}</w:t></w:r></w:p>`;
      const { session, save } = await load(makeDocx(documentOf(styled('Styled A'), styled('Styled B')), LISTS_NUMBERING, styles));
      createList(session, 'Styled A', 'bullet');
      const bytes = await save();
      const parts = unzipContainer(bytes);
      const document = text(parts['word/document.xml']) ?? '';
      const ownProperties = (value: string) => {
        const saved = (document.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find(entry => entry.includes(`>${value}<`)) ?? '';
        return /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(saved)?.[0] ?? '';
      };
      const created = attribute(ownProperties('Styled A'), /<w:numId w:val="(\d+)"/);
      expect(created).toBeGreaterThan(3);
      expect(ownProperties('Styled B')).not.toContain('numPr');
      const numbering = text(parts['word/numbering.xml']) ?? '';
      const num = elements(numbering, 'num').find(entry => attribute(entry, NUM_ID) === created);
      const abstract = elements(numbering, 'abstractNum').find(entry => attribute(entry, ABSTRACT_ID) === attribute(num, POINTER));
      expect(formats(abstract ?? '')[0]).toBe('bullet');
      const reopened = await seeded(bytes);
      const reopenedProperties = (value: string) => reopened.paragraphs('body').find(entry => entry.text === value)?.properties;
      expect(reopenedProperties('Styled A')).toMatchObject({ numPr: { numId: created }, listIsBullet: true });
      expect(reopenedProperties('Styled B')).toMatchObject({ numPr: { numId: 1 }, listNumFmt: 'decimal' });
    });
  }
}

describe('the numbering part a save adds, its content type and relationship', () => {
  const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const NUMBERING_TYPE = `${OFFICE}.numbering+xml`;
  const ROOT_RELATIONSHIPS = `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`;
  const DOCUMENT_RELATIONSHIPS = `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/></Relationships>`;
  const DEFINITION: DocxSavedListDefinition = {
    key: '1:0',
    kind: 'bullet',
    levelTexts: LEVELS.map(() => 'b'),
    levelFormats: LEVELS.map(() => 'bullet'),
    numId: 1,
    abstractNumId: 0,
  };
  const written = (types: string, relationships = DOCUMENT_RELATIONSHIPS) => {
    const parts: Record<string, Uint8Array> = {
      '[Content_Types].xml': toBytes(types),
      '_rels/.rels': toBytes(ROOT_RELATIONSHIPS),
      'word/_rels/document.xml.rels': toBytes(relationships),
    };
    expect(writeListDefinitionParts(parts, [DEFINITION])).toBe(true);
    expect(parts['word/numbering.xml']).toBeDefined();
    return { types: text(parts['[Content_Types].xml']) ?? '', relationships: text(parts['word/_rels/document.xml.rels']) ?? '' };
  };
  /** That `result` is `source` with exactly one `inserted` at `at`. */
  const insertedOnce = (source: string, result: string, at: number, inserted: RegExp) => {
    expect(result.slice(0, at)).toBe(source.slice(0, at));
    expect(result.endsWith(source.slice(at))).toBe(true);
    expect(result.slice(at, result.length - (source.length - at))).toMatch(inserted);
  };

  const present: Array<[string, string]> = [
    ['an override in double quotes', `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING_TYPE}"/>`],
    ['an override in single quotes', `<Override PartName='/word/numbering.xml' ContentType='${NUMBERING_TYPE}'/>`],
    ['an override with its attributes in another order', `<Override ContentType="${NUMBERING_TYPE}" PartName="/word/numbering.xml"/>`],
    ['an override with white space and line breaks in its start tag', `<Override\n  PartName = "/word/numbering.xml"\r\n\tContentType="${NUMBERING_TYPE}"\n/>`],
    ['an override naming the part in another case', `<Override PartName="/WORD/Numbering.XML" ContentType="${NUMBERING_TYPE}"/>`],
  ];
  for (const [name, override] of present) {
    test(`content types with ${name} are kept as they are`, () => {
      const types = `<Types xmlns="${TYPES}"><Default Extension="xml" ContentType="application/xml"/>${override}</Types>`;
      expect(written(types).types).toBe(types);
    });
  }

  test('content types with a prefixed root that names the part are kept as they are', () => {
    const types = `<ct:Types xmlns:ct="${TYPES}"><ct:Default Extension="xml" ContentType="application/xml"/><ct:Override PartName="/word/numbering.xml" ContentType="${NUMBERING_TYPE}"/></ct:Types>`;
    expect(written(types).types).toBe(types);
  });

  test('content types with a prefixed root gain one override in that prefix', () => {
    const types = `<ct:Types xmlns:ct="${TYPES}"><ct:Default Extension="xml" ContentType="application/xml"/></ct:Types>`;
    const result = written(types).types;
    insertedOnce(types, result, types.lastIndexOf('</ct:Types>'), /^<ct:Override PartName="\/word\/numbering\.xml" ContentType="[^"]+"\/>$/);
  });

  test('content types with a self-closing root gain one override inside it', () => {
    const types = `<?xml version="1.0"?>\n<Types xmlns="${TYPES}"/>`;
    const result = written(types).types;
    expect(result).toBe(`<?xml version="1.0"?>\n<Types xmlns="${TYPES}"><Override PartName="/word/numbering.xml" ContentType="${NUMBERING_TYPE}"/></Types>`);
  });

  test('content types without the override gain exactly one before the root end', () => {
    const types = `<Types xmlns="${TYPES}">\n  <Default Extension="xml" ContentType="application/xml"/>\n</Types >\n`;
    const result = written(types).types;
    insertedOnce(types, result, types.lastIndexOf('</Types'), /^<Override PartName="\/word\/numbering\.xml" ContentType="[^"]+"\/>$/);
  });

  test('relationships with a prefixed root gain one numbering relationship in that prefix', () => {
    const relationships = `<r:Relationships xmlns:r="${PACKAGE_RELATIONSHIPS}"><r:Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/></r:Relationships>`;
    const result = written(`<Types xmlns="${TYPES}"/>`, relationships).relationships;
    insertedOnce(relationships, result, relationships.lastIndexOf('</r:Relationships>'),
      /^<r:Relationship Id="rId2" Type="[^"]+\/numbering" Target="numbering\.xml"\/>$/);
  });

  test('relationships with a self-closing root gain one numbering relationship inside it', () => {
    const relationships = `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"/>`;
    const result = written(`<Types xmlns="${TYPES}"/>`, relationships).relationships;
    expect(result).toBe(`<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`);
  });
});

const decimalLevel = (prefix: string) =>
  `<${prefix}lvl ${prefix}ilvl="0"><${prefix}start ${prefix}val="1"/><${prefix}numFmt ${prefix}val="decimal"/><${prefix}lvlText ${prefix}val="%1."/></${prefix}lvl>`;

/** A schema-valid numbering part: the main namespace as the default for elements, attributes qualified with `w:`. */
const QUALIFIED_DEFAULT = `<numbering xmlns="${W}" xmlns:w="${W}"><abstractNum w:abstractNumId="0"><lvl w:ilvl="0"><start w:val="1"/>` +
  `<numFmt w:val="decimal"/><lvlText w:val="%1."/></lvl></abstractNum><num w:numId="1"><abstractNumId w:val="0"/></num></numbering>`;

/** Numbering parts whose list elements are not all spelled with the root's prefix, and how a created list is spelled in each. */
const SPELLINGS: Array<{ name: string; xml: string; abstractStart: string; numStart: string }> = [
  {
    // Mirrors a schema-invalid input as found: unqualified attributes under the default namespace.
    name: 'with the main namespace as its default namespace',
    xml: `<numbering xmlns="${W}"><abstractNum abstractNumId="0">${decimalLevel('')}</abstractNum><num numId="1"><abstractNumId val="0"/></num></numbering>`,
    abstractStart: '<abstractNum abstractNumId="2">',
    numStart: '<num numId="3">',
  },
  {
    name: 'whose lists sit under a second prefix bound on the root',
    xml: `<w:numbering xmlns:w="${W}" xmlns:x="${W}"><x:abstractNum x:abstractNumId="0">${decimalLevel('x:')}</x:abstractNum><x:num x:numId="1"><x:abstractNumId x:val="0"/></x:num></w:numbering>`,
    abstractStart: '<w:abstractNum w:abstractNumId="2">',
    numStart: '<w:num w:numId="3">',
  },
  {
    name: 'with lists that bind the main namespace themselves',
    xml: `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${decimalLevel('w:')}</w:abstractNum>` +
      `<w2:abstractNum xmlns:w2="${W}" w2:abstractNumId="1">${decimalLevel('w2:')}</w2:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w2:num xmlns:w2="${W}" w2:numId="2"><w2:abstractNumId w2:val="1"/></w2:num></w:numbering>`,
    abstractStart: '<w:abstractNum w:abstractNumId="2">',
    numStart: '<w:num w:numId="3">',
  },
  {
    name: 'with the main namespace as its default namespace and its attributes qualified',
    xml: QUALIFIED_DEFAULT,
    abstractStart: '<abstractNum w:abstractNumId="2">',
    numStart: '<num w:numId="3">',
  },
];

/** Where the list elements of `numbering` start, in any spelling, and the ids their start tags name. */
function listStarts(numbering: string, local: 'abstractNum' | 'num', id: 'abstractNumId' | 'numId') {
  return [...numbering.matchAll(new RegExp(`<(?:[\\w.-]+:)?${local}(?=[\\s/>])[^>]*>`, 'g'))].map(match => ({
    index: match.index ?? 0,
    id: Number(new RegExp(`\\s(?:[\\w.-]+:)?${id}\\s*=\\s*["'](\\d+)["']`).exec(match[0])?.[1]),
  }));
}

/** That the list `numId` and the definition it points at stand in schema order among the lists of `numbering`. */
function expectInSchemaOrder(numbering: string, numId: number): void {
  const nums = listStarts(numbering, 'num', 'numId');
  const abstracts = listStarts(numbering, 'abstractNum', 'abstractNumId');
  const num = nums.find(entry => entry.id === numId);
  expect(num).toBeDefined();
  const pointer = Number(/<(?:[\w.-]+:)?abstractNumId\s[^>]*\bval\s*=\s*["'](\d+)["']/.exec(numbering.slice(num!.index))?.[1]);
  const abstract = abstracts.find(entry => entry.id === pointer);
  expect(abstract).toBeDefined();
  expect(abstracts.every(entry => entry.index <= abstract!.index)).toBe(true);
  expect(nums.every(entry => entry.index > abstract!.index)).toBe(true);
  expect(nums.every(entry => entry.index <= num!.index)).toBe(true);
}

describe('created lists in numbering parts spelled with other prefixes', () => {
  const CREATED: DocxSavedListDefinition = {
    key: '9:1',
    kind: 'numbered',
    levelTexts: LEVELS.map(level => `%${level + 1}.`),
    levelFormats: LEVELS.map(() => 'decimal'),
    numId: 3,
    abstractNumId: 2,
  };
  for (const { name, xml, abstractStart, numStart } of SPELLINGS) {
    test(`a numbering part ${name} takes a created list in its own spelling and schema order, and recognises it again`, () => {
      const result = numberingWithListDefinitions(xml, [CREATED]);
      expect(result.indexOf(abstractStart)).toBeGreaterThan(-1);
      expect(result.indexOf(numStart)).toBeGreaterThan(-1);
      expectInSchemaOrder(result, CREATED.numId);
      expect(numberingWithListDefinitions(result, [CREATED])).toBe(result);
    });
  }
});

for (const [saveName, save] of Object.entries(SAVES)) {
  for (const { name, xml } of SPELLINGS) {
    test(`${saveName}: a created list saved into a numbering part ${name} stands in schema order above its ids`, async () => {
      const session = await seeded(makeDocx(documentOf(paragraph('First'), paragraph('Second')), xml));
      createList(session, 'Second', 'numbered');
      const parts = unzipContainer(await save(session));
      const numId = savedNumId(text(parts['word/document.xml']) ?? '', 'Second');
      const numbering = text(parts['word/numbering.xml']) ?? '';
      expect(numId).toBeGreaterThan(Math.max(...listStarts(xml, 'num', 'numId').map(entry => entry.id)));
      expectInSchemaOrder(numbering, numId!);
    });
  }
}

describe('a numbering part with the main namespace as default and qualified attributes', () => {
  const CREATED: DocxSavedListDefinition = {
    key: '9:1',
    kind: 'numbered',
    levelTexts: LEVELS.map(level => `%${level + 1}.`),
    levelFormats: LEVELS.map(() => 'decimal'),
    numId: 3,
    abstractNumId: 2,
  };

  test('takes a created list with an unprefixed element and a qualified id, and writes it once', () => {
    const result = numberingWithListDefinitions(QUALIFIED_DEFAULT, [CREATED]);
    expect(result).toContain('<abstractNum w:abstractNumId="2">');
    expect(result).not.toMatch(/<abstractNum\s+abstractNumId\s*=/);
    expect(result).toContain('<num w:numId="3">');
    expect(listStarts(numberingWithListDefinitions(result, [CREATED]), 'abstractNum', 'abstractNumId')).toHaveLength(2);
  });

  test('saveYrsDocx: a save over the first save writes the created list no second time', async () => {
    const session = await seeded(makeDocx(documentOf(paragraph('First'), paragraph('Second')), QUALIFIED_DEFAULT));
    createList(session, 'Second', 'numbered');
    const first = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    const second = (await saveYrsDocx(session, {
      updateModifiedDate: false,
      commentBaseline: { bytes: first.slice(), comments: session.listComments() },
      commentDeletion: true,
    })).bytes;
    const numbering = (bytes: Uint8Array) => text(unzipContainer(bytes)['word/numbering.xml']) ?? '';
    expect(numbering(first)).not.toMatch(/<abstractNum\s+abstractNumId\s*=/);
    expect(listStarts(numbering(first), 'abstractNum', 'abstractNumId')).toHaveLength(2);
    expect(listStarts(numbering(second), 'abstractNum', 'abstractNumId')).toHaveLength(2);
    expect(listStarts(numbering(second), 'num', 'numId')).toHaveLength(2);
    expect(savedNumId(text(unzipContainer(second)['word/document.xml']) ?? '', 'Second'))
      .toBe(savedNumId(text(unzipContainer(first)['word/document.xml']) ?? '', 'Second'));
  });
});

describe('a created list saved from a parsed document over a file its client saved before', () => {
  test('a key drawn again through the TypeScript loader names a new definition, kept on a save over its own', async () => {
    const first = await createYrsSession({ clientId: 78 });
    open.push(first);
    first.seedFromDocx(makeDocx(documentOf(
      paragraph('Lead'), paragraph('Alpha'), paragraph('Gap'), paragraph('Beta'), paragraph('Tail'),
      paragraph('Decimal source', 1), paragraph('Bullet source', 2),
    ), LISTS_NUMBERING));
    createList(first, 'Alpha', 'bullet');
    const earlier = (await saveYrsDocx(first, { updateModifiedDate: false })).bytes;
    const source = earlier.slice().buffer as ArrayBuffer;
    const parsed = await parseDocx(source, { preloadFonts: false });
    const reopened = await createYrsSession({ clientId: 78 });
    open.push(reopened);
    documentToYrs(reopened, parsed);
    createList(reopened, 'Beta', 'bullet');
    const keys = (target: YrsSession) => target.createdListDefinitions().map(definition => definition.key);
    expect(keys(reopened)).toEqual(keys(first));
    const write = async (base: ArrayBuffer) => (await writeSessionSave(
      reopened, yrsToDocument(reopened, { ...parsed, originalBuffer: base }), captureSessionSave(reopened), base, {}, () => false
    )).bytes;
    const saved = await write(parsed.originalBuffer ?? source);
    const again = await write(saved.slice().buffer as ArrayBuffer);
    const numIdOf = (bytes: Uint8Array, value: string) => savedNumId(text(unzipContainer(bytes)['word/document.xml']) ?? '', value);
    const counted = (bytes: Uint8Array) => {
      const numbering = text(unzipContainer(bytes)['word/numbering.xml']) ?? '';
      return { abstractNums: elements(numbering, 'abstractNum').length, nums: elements(numbering, 'num').length };
    };
    const beta = numIdOf(saved, 'Beta');
    expect(beta).not.toBe(numIdOf(saved, 'Alpha'));
    expect(counted(saved)).toEqual({ abstractNums: counted(earlier).abstractNums + 1, nums: counted(earlier).nums + 1 });
    expect(numIdOf(again, 'Beta')).toBe(beta);
    expect(counted(again)).toEqual(counted(saved));
    expect(beta).toBeGreaterThan(3);
  });
});

describe('a numbering part or main document named through a dot segment', () => {
  const NUMBERED = () => makeDocx(documentOf(paragraph('First'), paragraph('Second')), LISTS_NUMBERING);
  const dotted = (parts: Record<string, Uint8Array>) =>
    Object.keys(parts).filter(path => path.split('/').some(segment => segment === '.' || segment === '..'));

  for (const [saveName, save] of Object.entries(SAVES)) {
    for (const target of ['./numbering.xml', '../word/numbering.xml']) {
      test(`${saveName}: a numbering relationship that names ${target} has the created list written into word/numbering.xml`, async () => {
        const session = await seeded(repackaged(NUMBERED(), {
          'word/_rels/document.xml.rels': `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R}/numbering" Target="${target}"/></Relationships>`,
        }));
        createList(session, 'Second', 'numbered');
        const saved = await settled(save(session));
        expect(saved.error).toBeUndefined();
        const parts = unzipContainer(saved.value!);
        expect(dotted(parts)).toEqual([]);
        const numbering = text(parts['word/numbering.xml']) ?? '';
        expect(elements(numbering, 'abstractNum')).toHaveLength(elements(LISTS_NUMBERING, 'abstractNum').length + 1);
        expect(elements(numbering, 'num')).toHaveLength(elements(LISTS_NUMBERING, 'num').length + 1);
        const types = text(parts['[Content_Types].xml']) ?? '';
        expect(types).not.toContain('word/./numbering.xml');
        expect(types).not.toContain('word/../word/numbering.xml');
        expect(numberingRelationships(text(parts['word/_rels/document.xml.rels']) ?? '')).toHaveLength(1);
      });
    }

    test(`${saveName}: a package relationship that names ./word/document.xml has the created list written into word/numbering.xml`, async () => {
      const opened = await settled(seeded(repackaged(NUMBERED(), {
        '_rels/.rels': `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="doc" Type="${R}/officeDocument" Target="./word/document.xml"/></Relationships>`,
      })));
      expect(opened.error).toBeUndefined();
      const session = opened.value!;
      createList(session, 'Second', 'numbered');
      const saved = await settled(save(session));
      expect(saved.error).toBeUndefined();
      const parts = unzipContainer(saved.value!);
      expect(dotted(parts)).toEqual([]);
      const numbering = text(parts['word/numbering.xml']) ?? '';
      expect(elements(numbering, 'abstractNum')).toHaveLength(elements(LISTS_NUMBERING, 'abstractNum').length + 1);
    });
  }
});

describe('resolvedRelationshipTarget', () => {
  const DOCUMENT_RELS = 'word/_rels/document.xml.rels';
  const TABLE: Array<[owner: string, target: string, expected: string]> = [
    [DOCUMENT_RELS, 'numbering.xml', 'word/numbering.xml'],
    [DOCUMENT_RELS, './numbering.xml', 'word/numbering.xml'],
    [DOCUMENT_RELS, '../numbering.xml', 'numbering.xml'],
    [DOCUMENT_RELS, '../word/numbering.xml', 'word/numbering.xml'],
    [DOCUMENT_RELS, '/word/numbering.xml', 'word/numbering.xml'],
    [DOCUMENT_RELS, 'lists/parts/numbering.xml', 'word/lists/parts/numbering.xml'],
    [DOCUMENT_RELS, '../../../numbering.xml', 'numbering.xml'],
    [DOCUMENT_RELS, 'lists\\numbering.xml', 'word/lists/numbering.xml'],
    [DOCUMENT_RELS, 'lists//./numbering.xml', 'word/lists/numbering.xml'],
    ['word/document.xml', 'numbering.xml', 'word/numbering.xml'],
    ['docs/main/_rels/document.xml.rels', '../numbering.xml', 'docs/numbering.xml'],
    ['_rels/.rels', 'word/document.xml', 'word/document.xml'],
    ['_rels/.rels', './word/document.xml', 'word/document.xml'],
    ['_rels/.rels', '/word/document.xml', 'word/document.xml'],
  ];

  test('resolves each target against the part that owns the relationships', () => {
    expect(TABLE.map(([owner, target]) => resolvedRelationshipTarget(owner, target))).toEqual(TABLE.map(([, , expected]) => expected));
  });
});
