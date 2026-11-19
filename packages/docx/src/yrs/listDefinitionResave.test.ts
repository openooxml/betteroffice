import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { BlockContent, NumberingDefinitions, Paragraph } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { saveEditorDocument } from './editorSave';
import { createYrsSession, saveYrsDocx, yrsToDocument, type YrsSession } from './index';
import { createdListKey, numberingWithListDefinitions, savedListDefinitions } from './listDefinitionSave';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const LISTS_NUMBERING = readFileSync(resolve(import.meta.dir, '__fixtures__/seed-parity/lists/word/numbering.xml'), 'utf8');
const SHAPES = JSON.parse(readFileSync(resolve(import.meta.dir, '__fixtures__/created-list-shapes.json'), 'utf8')) as Array<{
  name: string;
  attrs: Record<string, unknown>;
  key: string | null;
}>;
const LEVELS = [0, 1, 2, 3, 4, 5, 6, 7, 8];

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
const formats = (abstractNum: string) => [...abstractNum.matchAll(/<w:numFmt w:val="([^"]*)"/g)].map(match => match[1]);
const levels = (abstractNum: string) => [...abstractNum.matchAll(/<w:lvl\b[^>]*\bw:ilvl="(\d+)"/g)].map(match => Number(match[1]));

function savedNumId(bytes: Uint8Array, value: string): number | undefined {
  const document = text(unzipContainer(bytes)['word/document.xml']);
  const saved = (document.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find(entry => entry.includes(`>${value}<`));
  return attribute(saved, /<w:numId w:val="(\d+)"/);
}

/** The format of the first level of the definition `numId` points at in `bytes`. */
function savedFormat(bytes: Uint8Array, numId: number | undefined): string | undefined {
  const numbering = text(unzipContainer(bytes)['word/numbering.xml']);
  const num = elements(numbering, 'num').find(entry => attribute(entry, NUM_ID) === numId);
  const abstract = elements(numbering, 'abstractNum').find(entry => attribute(entry, ABSTRACT_ID) === attribute(num, POINTER));
  return abstract === undefined ? undefined : formats(abstract)[0];
}

function counts(bytes: Uint8Array) {
  const numbering = text(unzipContainer(bytes)['word/numbering.xml']);
  return { abstractNums: elements(numbering, 'abstractNum').length, nums: elements(numbering, 'num').length };
}

function highestNumId(bytes: Uint8Array): number {
  const numbering = text(unzipContainer(bytes)['word/numbering.xml']);
  return Math.max(0, ...elements(numbering, 'num').map(num => attribute(num, NUM_ID) ?? 0));
}

const SOURCE = {
  abstractNums: elements(LISTS_NUMBERING, 'abstractNum').length,
  nums: elements(LISTS_NUMBERING, 'num').length,
};

const open: YrsSession[] = [];
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { open.splice(0).forEach(session => session.destroy()); });

async function session(clientId?: number): Promise<YrsSession> {
  const created = await createYrsSession(clientId === undefined ? undefined : { clientId });
  open.push(created);
  return created;
}

function createList(target: YrsSession, value: string, kind: 'bullet' | 'numbered'): void {
  const found = target.paragraphs('body').find(entry => entry.text === value);
  expect(found).toBeDefined();
  const paraId = found!.paraId;
  const attrs = target.ensureListDefinition('body', paraId, kind);
  target.setParagraphAttrs({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, attrs);
}

/** `promise`'s value, or the error it rejects with. */
async function settled<T>(promise: Promise<T>): Promise<{ value?: T; error?: unknown }> {
  try {
    return { value: await promise };
  } catch (error) {
    return { error };
  }
}

function projectedParagraph(blocks: readonly BlockContent[], value: string): Paragraph | undefined {
  return blocks.find((block): block is Paragraph =>
    block.type === 'paragraph' && block.content.some(child => child.type === 'run' && child.content.some(entry => entry.type === 'text' && entry.text === value)));
}

describe('saving into a package an earlier save of the document wrote', () => {
  test('the same session saving over its own earlier save keeps each created definition once with its ids', async () => {
    const first = await session(5);
    first.seedFromDocx(makeDocx(DOCUMENT, LISTS_NUMBERING));
    createList(first, 'Alpha', 'bullet');
    const earlier = (await saveYrsDocx(first, { updateModifiedDate: false })).bytes;
    const alpha = savedNumId(earlier, 'Alpha');
    expect(alpha).toBeGreaterThan(3);
    const again = await settled(saveYrsDocx(first, {
      updateModifiedDate: false,
      commentBaseline: { bytes: earlier.slice(), comments: first.listComments() },
      commentDeletion: true,
    }));
    expect(again.error).toBeUndefined();
    const bytes = again.value!.bytes;
    expect(counts(bytes)).toEqual({ abstractNums: SOURCE.abstractNums + 1, nums: SOURCE.nums + 1 });
    expect(savedNumId(bytes, 'Alpha')).toBe(alpha);
    expect(savedFormat(bytes, alpha)).toBe('bullet');
  });

  test('a fresh session with a created key that sorts first keeps the ids the earlier save gave', async () => {
    const first = await session(5);
    first.seedFromDocx(makeDocx(DOCUMENT, LISTS_NUMBERING));
    createList(first, 'Alpha', 'bullet');
    const earlier = (await saveYrsDocx(first, { updateModifiedDate: false })).bytes;
    const alpha = savedNumId(earlier, 'Alpha');
    const draft = await session(41);
    draft.openDocx(new Uint8Array(first.materializeDocx()!.originalBuffer!), false);
    draft.loadState(first.encodeState());
    createList(draft, 'Beta', 'numbered');
    const base = draft.materializeDocx()!;
    const projected = yrsToDocument(draft, { ...base, originalBuffer: earlier.slice().buffer });
    const saved = await settled(saveEditorDocument(draft, projected, []));
    expect(saved.error).toBeUndefined();
    const bytes = new Uint8Array(saved.value!);
    expect(counts(bytes)).toEqual({ abstractNums: SOURCE.abstractNums + 2, nums: SOURCE.nums + 2 });
    expect(savedNumId(bytes, 'Alpha')).toBe(alpha);
    expect(savedFormat(bytes, alpha)).toBe('bullet');
    const beta = savedNumId(bytes, 'Beta');
    expect(beta).toBeGreaterThan(highestNumId(earlier));
    expect(savedFormat(bytes, beta)).toBe('decimal');
  });
});

describe('the table a save gives created lists', () => {
  test('a projection that is not a save keeps the numId a created list writes for earlier versions', async () => {
    const created = await session();
    created.seedFromDocx(makeDocx(DOCUMENT, LISTS_NUMBERING));
    createList(created, 'Alpha', 'bullet');
    createList(created, 'Beta', 'numbered');
    const projected = yrsToDocument(created, created.materializeDocx()!);
    const blocks = projected.package.document.content;
    expect(projectedParagraph(blocks, 'Alpha')?.formatting?.numPr?.numId).toBe(1);
    expect(projectedParagraph(blocks, 'Beta')?.formatting?.numPr?.numId).toBe(2);
  });

  test('a save reads the created definitions once, whatever the number of paragraphs on created lists', async () => {
    const created = await session();
    created.seedFromDocx(makeDocx(DOCUMENT, LISTS_NUMBERING));
    createList(created, 'Alpha', 'bullet');
    createList(created, 'Beta', 'numbered');
    createList(created, 'Tail', 'numbered');
    const read = created.createdListDefinitions.bind(created);
    let calls = 0;
    created.createdListDefinitions = () => {
      calls += 1;
      return read();
    };
    const bytes = (await saveYrsDocx(created, { updateModifiedDate: false })).bytes;
    expect(calls).toBe(1);
    expect(savedNumId(bytes, 'Tail')).toBe(savedNumId(bytes, 'Beta'));
  });
});

test('the created-list rule matches the engine over odd attribute shapes', () => {
  const keys = SHAPES.map(shape => [shape.name, createdListKey(shape.attrs) ?? null]);
  expect(keys).toEqual(SHAPES.map(shape => [shape.name, shape.key]));
});

describe('definitions written from peer text', () => {
  test('at most nine levels, and only characters XML allows', () => {
    const levelTexts = ['a\u0001', 'b\ud800', 'c￿', 'd&<"', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l'];
    const xml = numberingWithListDefinitions(LISTS_NUMBERING, [{
      key: '9:"<&>',
      kind: 'bullet',
      levelTexts,
      levelFormats: levelTexts.map(() => 'bullet"/>'),
      numId: 4,
      abstractNumId: 2,
    }]);
    const created = elements(xml, 'abstractNum').at(-1) ?? '';
    expect(attribute(created, ABSTRACT_ID)).toBe(2);
    expect(levels(created)).toEqual(LEVELS);
    expect(created).not.toMatch(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/u);
    expect(created).not.toMatch(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/);
    expect(created).toContain('d&amp;&lt;&quot;');
    expect(created).not.toContain('bullet"/>"');
    expect(xml.startsWith(LISTS_NUMBERING.slice(0, LISTS_NUMBERING.indexOf('<w:abstractNum ')))).toBe(true);
  });
});

const numberingParts = (numbering: string) => unzipContainer(makeDocx(DOCUMENT, numbering));
const uniqueIds = (xml: string, name: 'numId' | 'abstractNumId') => {
  const ids = [...xml.matchAll(new RegExp(`:${name}\\s*=\\s*["'](\\d+)["']`, 'g'))].map(match => match[1]);
  return new Set(ids).size === ids.length;
};
const BULLETS = {
  key: '1:0',
  kind: 'bullet' as const,
  levelTexts: LEVELS.map(level => `b${level}`),
  levelFormats: LEVELS.map(() => 'bullet'),
};

describe('the ids a save reads from a numbering part', () => {
  test('single quotes and spaces around = are read, so the created list goes above them', () => {
    const quoted = LISTS_NUMBERING.replace(/"/g, "'").replace(/='/g, " = '");
    const saved = savedListDefinitions({} as YrsSession, [BULLETS], numberingParts(quoted));
    expect(saved.map(({ numId, abstractNumId }) => ({ numId, abstractNumId }))).toEqual([{ numId: 4, abstractNumId: 2 }]);
    const written = numberingWithListDefinitions(quoted, saved);
    expect(uniqueIds(written, 'numId')).toBe(true);
    expect(uniqueIds(written, 'abstractNumId')).toBe(true);
  });

  test('an id the element scan misses still raises the ids handed out, as does the model of the source', () => {
    const foreign = LISTS_NUMBERING.replace(
      '</w:numbering>',
      `<ns0:abstractNum xmlns:ns0="${W}" ns0:abstractNumId="7"/><ns0:num xmlns:ns0="${W}" ns0:numId="12"><ns0:abstractNumId ns0:val="7"/></ns0:num></w:numbering>`
    );
    const [scanned] = savedListDefinitions({} as YrsSession, [BULLETS], numberingParts(foreign));
    expect(scanned?.numId).toBeGreaterThan(12);
    expect(scanned?.abstractNumId).toBeGreaterThan(7);
    const model = { abstractNums: [{ abstractNumId: 30 }], nums: [{ numId: 40, abstractNumId: 30 }] } as unknown as NumberingDefinitions;
    const [modelled] = savedListDefinitions({} as YrsSession, [BULLETS], numberingParts(LISTS_NUMBERING), model);
    expect(modelled?.numId).toBeGreaterThan(40);
    expect(modelled?.abstractNumId).toBeGreaterThan(30);
  });
});

describe('created lists carried by a saved file', () => {
  test('a key drawn again by a new session on the saved file names a different definition', async () => {
    const first = await session(77);
    first.seedFromDocx(makeDocx(DOCUMENT, LISTS_NUMBERING));
    createList(first, 'Alpha', 'bullet');
    const earlier = (await saveYrsDocx(first, { updateModifiedDate: false })).bytes;
    const reopened = await session(77);
    reopened.seedFromDocx(earlier.slice());
    createList(reopened, 'Beta', 'numbered');
    const keys = (target: YrsSession) => target.createdListDefinitions().map(definition => definition.key);
    expect(keys(reopened)).toEqual(keys(first));
    const bytes = (await saveYrsDocx(reopened, { updateModifiedDate: false })).bytes;
    expect(counts(bytes)).toEqual({ abstractNums: SOURCE.abstractNums + 2, nums: SOURCE.nums + 2 });
    const alpha = savedNumId(bytes, 'Alpha');
    const beta = savedNumId(bytes, 'Beta');
    expect(beta).not.toBe(alpha);
    expect(savedFormat(bytes, alpha)).toBe('bullet');
    expect(savedFormat(bytes, beta)).toBe('decimal');
  });

  test('a structural copy of the projection still saves the created ids', async () => {
    const created = await session();
    created.seedFromDocx(makeDocx(DOCUMENT, LISTS_NUMBERING));
    createList(created, 'Alpha', 'bullet');
    const projected = structuredClone(yrsToDocument(created, created.materializeDocx()!));
    const bytes = new Uint8Array(await saveEditorDocument(created, projected, [], { full: true }));
    const alpha = savedNumId(bytes, 'Alpha');
    expect(alpha).toBeGreaterThan(3);
    expect(savedFormat(bytes, alpha)).toBe('bullet');
  });
});

describe('a paragraph that joins a created list without its numPr changing', () => {
  const ITEM = `<w:document xmlns:w="${W}"><w:body>${[
    paragraph('Lead'), paragraph('Item', 1), paragraph('Tail'),
  ].join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  const saves: Record<string, (target: YrsSession) => Promise<Uint8Array>> = {
    saveYrsDocx: async target => (await saveYrsDocx(target, { updateModifiedDate: false })).bytes,
    editor: async target => {
      const projected = yrsToDocument(target, target.materializeDocx()!);
      return new Uint8Array(await saveEditorDocument(target, projected, projected.package.document.comments ?? []));
    },
  };
  for (const [name, save] of Object.entries(saves)) {
    test(`${name}: the saved paragraph points at the created definition`, async () => {
      const created = await session();
      created.seedFromDocx(makeDocx(ITEM, LISTS_NUMBERING));
      createList(created, 'Item', 'bullet');
      const bytes = await save(created);
      const item = savedNumId(bytes, 'Item');
      expect(item).toBeGreaterThan(3);
      expect(savedFormat(bytes, item)).toBe('bullet');
    });
  }
});
