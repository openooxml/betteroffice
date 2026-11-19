import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { createdListKey } from './listDefinitionSave';

const ROOT = resolve(import.meta.dir, '../../../..');
const FIXTURE = resolve(import.meta.dir, '__fixtures__/seed-parity/lists/word');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const STYLES = `${XML}<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="240" w:after="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style></w:styles>`;
const BULLET_NUMBERING = `${XML}<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const NUMBERED_HEADING_STYLES = STYLES.replace(
  '<w:pPr><w:outlineLvl w:val="0"/>',
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/>'
);
const HEADING_NUMBERING = `${XML}<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

/** The paragraph the list commands are applied to; the last paragraph added to each document. */
export const PLAIN = 'Plain paragraph QA';
/** A paragraph on no list, put directly above `PLAIN` when no list may be beside it. */
export const GAP = 'Gap paragraph QA';
/** A paragraph on no list that can follow `PLAIN`. */
export const FOLLOWER = 'Follower paragraph QA';
/** The item of the bullet list 1 that opens document B. */
export const EXISTING_BULLET = 'Existing bullet item QA';
/** An item of the decimal list 1 of document A, directly above an item of the bullet list 2. */
export const DECIMAL_ITEM = 'Second numbered item';
/** The level 2 item of the decimal list 1 of document A directly above `DECIMAL_ITEM`. */
const NESTED_DECIMAL = 'Nested level three';
/** The item of the decimal list 3 of document A directly below a level 1 item of the bullet list 2. */
const RESTARTED_ITEM = 'Restarted list with explicit indent override';
/** The two headings of document D, numbered by their style. */
const HEADING = 'Introduction heading QA';
const NEXT_HEADING = 'Second heading QA';
const SECOND_FOLLOWER = 'Second follower QA';
const THIRD_FOLLOWER = 'Third follower QA';
export const LIST_LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});

/**
 * A: the seed-parity lists document, whose list 1 is decimal, list 2 a bullet list and list 3
 * decimal again, its last item directly above `PLAIN`. B: `EXISTING_BULLET` of bullet list 1,
 * directly above `PLAIN`, and no `w:num` 2. C: no numbering part. D: `PLAIN` between two Heading 1
 * paragraphs numbered by their style's list 1.
 */
export type ListDocument = 'A' | 'B' | 'C' | 'D';
export type ListKind = 'bullet' | 'numbered';
export interface ListDocumentOptions {
  /** Puts `GAP` directly above `PLAIN`. */
  isolated?: boolean;
  /** Plain paragraphs after `PLAIN`. */
  after?: readonly string[];
}

/** The `w:numId`s of the numbering each document opens with. */
export const SOURCE_NUM_IDS: Readonly<Record<ListDocument, readonly string[]>> = { A: ['1', '2', '3'], B: ['1'], C: [], D: ['1'] };

export async function preloadListWasm(): Promise<void> {
  const generated = (path: string) => new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated', path)));
  await preloadEditWasm(generated('edit/docx_edit_bg.wasm'));
  await preloadOpcWasm(generated('opc/ooxml_opc_bg.wasm'));
  await preloadParseWasm(generated('parse/docx_parse_bg.wasm'));
}

export function listFont(): Uint8Array {
  return new Uint8Array(readFileSync(resolve(ROOT, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')));
}

const plain = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function pack(document: string, numbering: string | null, relationships?: Uint8Array, styles = STYLES): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/>${numbering ? `<Override PartName="/word/numbering.xml" ContentType="${OFFICE}.numbering+xml"/>` : ''}</Types>`));
  parts.set('_rels/.rels', toBytes(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`));
  parts.set('word/_rels/document.xml.rels', relationships ?? toBytes(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/>${numbering ? `<Relationship Id="rId2" Type="${R}/numbering" Target="numbering.xml"/>` : ''}</Relationships>`));
  parts.set('word/styles.xml', toBytes(styles));
  if (numbering) parts.set('word/numbering.xml', toBytes(numbering));
  parts.set('word/document.xml', toBytes(document));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

export function listDocument(doc: ListDocument, options: ListDocumentOptions = {}): Uint8Array {
  const added = [...(options.isolated ? [GAP] : []), PLAIN, ...(options.after ?? [])].map(plain).join('');
  if (doc === 'A') {
    const source = readFileSync(resolve(FIXTURE, 'document.xml'), 'utf8');
    const anchor = '    <w:sectPr>';
    expect(source.split(anchor)).toHaveLength(2);
    return pack(
      source.replace(anchor, `${added}\n${anchor}`),
      readFileSync(resolve(FIXTURE, 'numbering.xml'), 'utf8'),
      new Uint8Array(readFileSync(resolve(FIXTURE, '_rels/document.xml.rels')))
    );
  }
  if (doc === 'D') {
    const heading = (text: string) => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    return pack(
      `${XML}<w:document xmlns:w="${W}"><w:body>${heading(HEADING)}${added}${heading(NEXT_HEADING)}${SECT}</w:body></w:document>`,
      HEADING_NUMBERING, undefined, NUMBERED_HEADING_STYLES
    );
  }
  const lead = doc === 'B'
    ? `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${EXISTING_BULLET}</w:t></w:r></w:p>`
    : plain('Opening paragraph QA');
  return pack(`${XML}<w:document xmlns:w="${W}"><w:body>${lead}${added}${SECT}</w:body></w:document>`, doc === 'B' ? BULLET_NUMBERING : null);
}

function flatten(primitives: readonly unknown[], out: Record<string, unknown>[] = []): Record<string, unknown>[] {
  for (const primitive of primitives) {
    if (!primitive || typeof primitive !== 'object') continue;
    const record = primitive as Record<string, unknown>;
    out.push(record);
    if (Array.isArray(record.primitives)) flatten(record.primitives, out);
  }
  return out;
}

const blockKey = (primitive: Record<string, unknown>) =>
  primitive.blockKey !== undefined ? `k${String(primitive.blockKey)}`
    : primitive.blockId !== undefined ? `b${String(primitive.blockId)}`
      : primitive.paraId !== undefined ? `p${String(primitive.paraId)}` : undefined;

const unspaced = (value: string) => value.replace(/\s+/g, '');

/**
 * The list marker the display list `frame` shows before the one block reading `text`, without
 * white space; '' when it shows none, null when no single block reads `text`.
 */
export function shownMarker(frame: Uint8Array, text: string): string | null {
  const displayList = applyFrameDelta(null, decodeFrameDelta(frame)).displayList as unknown as {
    pages: ReadonlyArray<{ primitives?: readonly unknown[] }>;
  };
  for (const page of displayList.pages) {
    const blocks = new Map<string, { text: string; marker: string }>();
    for (const primitive of flatten(page.primitives ?? [])) {
      if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun') continue;
      const key = blockKey(primitive);
      if (key === undefined) continue;
      const block = blocks.get(key) ?? { text: '', marker: '' };
      blocks.set(key, block);
      if (primitive.listMarker === true) block.marker += String(primitive.text ?? '');
      else block.text += String(primitive.text ?? '');
    }
    const hits = [...blocks.values()].filter(block => unspaced(block.text) === unspaced(text));
    if (hits.length === 1) return unspaced(hits[0]!.marker);
    if (hits.length > 1) return null;
  }
  return null;
}

/** What a saved package says of the list of the paragraph reading `text`. */
export interface SavedList {
  /** The saved paragraphs reading `text`. */
  paragraphs: number;
  numId?: string;
  /** The `w:name` of the definition the paragraph's `w:num` points at. */
  name?: string;
  /** The `w:numFmt` of level 0 of that definition. */
  numFmt?: string;
  /** The `w:lvlText` of level 0 of that definition. */
  lvlText?: string;
}

export function savedList(bytes: Uint8Array, text: string): SavedList {
  const parts = unzipContainer(bytes);
  const decode = (path: string) => (parts[path] ? new TextDecoder().decode(parts[path]) : '');
  const textOf = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');
  const paragraphs = (decode('word/document.xml').match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []).filter(paragraph => textOf(paragraph) === text);
  const saved: SavedList = { paragraphs: paragraphs.length };
  const numPr = paragraphs.length === 1 ? paragraphs[0]!.match(/<w:numPr\b[^>]*>([\s\S]*?)<\/w:numPr>/)?.[1] : undefined;
  saved.numId = numPr?.match(/<w:numId\b[^>]*\bw:val="(\d+)"/)?.[1];
  if (saved.numId === undefined) return saved;
  const numbering = decode('word/numbering.xml');
  const num = numbering.match(new RegExp(`<w:num\\b[^>]*\\bw:numId="${saved.numId}"[^>]*>([\\s\\S]*?)</w:num>`))?.[1];
  const abstractNumId = num?.match(/<w:abstractNumId\b[^>]*\bw:val="(\d+)"/)?.[1];
  if (abstractNumId === undefined) return saved;
  const abstract = numbering.match(new RegExp(`<w:abstractNum\\b[^>]*\\bw:abstractNumId="${abstractNumId}"[^>]*>([\\s\\S]*?)</w:abstractNum>`))?.[1] ?? '';
  const level = abstract.match(/<w:lvl\b[^>]*\bw:ilvl="0"[^>]*>([\s\S]*?)<\/w:lvl>/)?.[1] ?? '';
  saved.name = abstract.match(/<w:name\b[^>]*\bw:val="([^"]*)"/)?.[1];
  saved.numFmt = level.match(/<w:numFmt\b[^>]*\bw:val="([^"]*)"/)?.[1];
  saved.lvlText = level.match(/<w:lvlText\b[^>]*\bw:val="([^"]*)"/)?.[1];
  return saved;
}

/** One way the list commands reach a document: the worker's command engine or the toolbar on a session. */
export interface ListCommandDriver {
  /** Bullets or Numbering over the paragraphs from `from` to `to`, a caret in `from` when `to` is left out. */
  toggle(kind: ListKind, from: string, to?: string): void;
  indent(direction: 'increase' | 'decrease', at: string): void;
  undo(): void;
  redo(): void;
  canUndo(): boolean;
  paragraphs(): ReadonlyArray<{ paraId: string; text: string; properties: Record<string, unknown> }>;
  definitions(): ReadonlyArray<{ key: string; kind: string }>;
  marker(text: string): string | null;
  save(): Promise<Uint8Array>;
}

export type OpenListCommandDriver = (doc: ListDocument, options?: ListDocumentOptions) => Promise<ListCommandDriver>;

function propertiesOf(driver: ListCommandDriver, text: string): Record<string, unknown> {
  const found = driver.paragraphs().filter(paragraph => paragraph.text === text);
  expect(found).toHaveLength(1);
  return found[0]!.properties;
}

const keyOf = (driver: ListCommandDriver, text: string) => createdListKey(propertiesOf(driver, text));

async function expectOnCreatedList(driver: ListCommandDriver, doc: ListDocument, text: string, kind: ListKind): Promise<string | undefined> {
  const key = keyOf(driver, text);
  expect(key).toBeDefined();
  expect(driver.definitions().find(definition => definition.key === key)?.kind).toBe(kind);
  expect(driver.marker(text)).toBe(kind === 'bullet' ? '•' : '1.');
  const saved = savedList(await driver.save(), text);
  expect(saved.numId ?? '').toMatch(/^\d+$/);
  expect(SOURCE_NUM_IDS[doc].includes(saved.numId ?? '')).toBe(false);
  expect(saved).toMatchObject({
    paragraphs: 1,
    name: `betteroffice-list:${key}`,
    numFmt: kind === 'bullet' ? 'bullet' : 'decimal',
    ...(kind === 'numbered' ? { lvlText: '%1.' } : {}),
  });
  return key;
}

function expectOffList(driver: ListCommandDriver, text: string): void {
  const properties = propertiesOf(driver, text);
  expect(properties.numPr ?? null).toBeNull();
  expect(properties.listDef ?? null).toBeNull();
}

const KINDS: readonly ListKind[] = ['bullet', 'numbered'];

/** The list command rows, run through `open`'s driver. */
export function listCommandScenarios(path: string, open: OpenListCommandDriver, timeout: number): void {
  for (const doc of ['A', 'B', 'C'] as const) {
    for (const kind of KINDS) {
      test(`${path}: ${kind} on a paragraph with no list beside it in document ${doc} creates a ${kind} list`, async () => {
        const driver = await open(doc, { isolated: true });
        driver.toggle(kind, PLAIN);
        await expectOnCreatedList(driver, doc, PLAIN, kind);
      }, timeout);
    }
  }

  test(`${path}: bullets directly below a source bullet item join its list`, async () => {
    const driver = await open('B');
    driver.toggle('bullet', PLAIN);
    const properties = propertiesOf(driver, PLAIN);
    expect(properties.numPr).toMatchObject({ numId: 1, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
    expect(driver.definitions()).toEqual([]);
    expect(driver.marker(PLAIN)).toBe('•');
    expect(savedList(await driver.save(), PLAIN)).toMatchObject({ paragraphs: 1, numId: '1', numFmt: 'bullet' });
  }, timeout);

  test(`${path}: numbering directly below a source numbered item joins its list`, async () => {
    const driver = await open('A');
    driver.toggle('numbered', PLAIN);
    const properties = propertiesOf(driver, PLAIN);
    expect(properties.numPr).toMatchObject({ numId: 3, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
    expect(driver.definitions()).toEqual([]);
    const above = driver.marker(RESTARTED_ITEM) ?? '';
    expect(above).toMatch(/^\d+\.$/);
    expect(driver.marker(PLAIN)).toBe(`${Number.parseInt(above, 10) + 1}.`);
    expect(savedList(await driver.save(), PLAIN)).toMatchObject({ paragraphs: 1, numId: '3', numFmt: 'decimal' });
  }, timeout);

  test(`${path}: numbering on the paragraph below a created numbered item continues that list`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('numbered', PLAIN);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    driver.toggle('numbered', FOLLOWER);
    expect(keyOf(driver, FOLLOWER)).toBe(key);
    expect(driver.definitions()).toHaveLength(1);
    expect(driver.marker(PLAIN)).toBe('1.');
    expect(driver.marker(FOLLOWER)).toBe('2.');
    const bytes = await driver.save();
    expect(savedList(bytes, FOLLOWER).numId).toBe(savedList(bytes, PLAIN).numId);
  }, timeout);

  for (const kind of KINDS) {
    test(`${path}: ${kind} on a ${kind} item takes the list off, and on again starts a list afresh`, async () => {
      const driver = await open('C');
      driver.toggle(kind, PLAIN);
      const first = keyOf(driver, PLAIN);
      expect(first).toBeDefined();
      driver.toggle(kind, PLAIN);
      expectOffList(driver, PLAIN);
      expect(driver.marker(PLAIN)).toBe('');
      driver.toggle(kind, PLAIN);
      const again = keyOf(driver, PLAIN);
      expect(again).toBeDefined();
      expect(again).not.toBe(first);
    }, timeout);
  }

  test(`${path}: numbering on a created bullet item puts it on a numbered list`, async () => {
    const driver = await open('C');
    driver.toggle('bullet', PLAIN);
    const bullet = keyOf(driver, PLAIN);
    expect(bullet).toBeDefined();
    driver.toggle('numbered', PLAIN);
    expect(await expectOnCreatedList(driver, 'C', PLAIN, 'numbered')).not.toBe(bullet);
  }, timeout);

  test(`${path}: numbering on a source bullet item with no numbered list beside it creates a numbered list`, async () => {
    const driver = await open('B', { isolated: true });
    driver.toggle('numbered', EXISTING_BULLET);
    await expectOnCreatedList(driver, 'B', EXISTING_BULLET, 'numbered');
  }, timeout);

  test(`${path}: bullets on an item of a decimal list 1 join the bullet list below instead of taking the list off`, async () => {
    const driver = await open('A');
    driver.toggle('bullet', DECIMAL_ITEM);
    const properties = propertiesOf(driver, DECIMAL_ITEM);
    expect(properties.numPr).toMatchObject({ numId: 2, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
    expect(driver.marker(DECIMAL_ITEM)).toBe('•');
    expect(savedList(await driver.save(), DECIMAL_ITEM)).toMatchObject({ paragraphs: 1, numId: '2', numFmt: 'bullet' });
  }, timeout);

  test(`${path}: indent and outdent on a created list change only its level, and outdent at the top takes the list off`, async () => {
    const driver = await open('C');
    driver.toggle('numbered', PLAIN);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    driver.indent('increase', PLAIN);
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ ilvl: 1 });
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(driver.marker(PLAIN)).toBe('a.');
    driver.indent('decrease', PLAIN);
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ ilvl: 0 });
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(driver.marker(PLAIN)).toBe('1.');
    driver.indent('decrease', PLAIN);
    expectOffList(driver, PLAIN);
  }, timeout);

  test(`${path}: one undo takes back the list and the definition bullets created`, async () => {
    const driver = await open('C');
    driver.toggle('bullet', PLAIN);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    expect(driver.definitions()).toHaveLength(1);
    driver.undo();
    expectOffList(driver, PLAIN);
    expect(driver.definitions()).toEqual([]);
    expect(driver.canUndo()).toBe(false);
    driver.redo();
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(driver.definitions()).toHaveLength(1);
  }, timeout);

  test(`${path}: bullets on two selected paragraphs put both on one list in one undo step`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('bullet', PLAIN, FOLLOWER);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    expect(keyOf(driver, FOLLOWER)).toBe(key);
    expect(driver.definitions()).toHaveLength(1);
    expect(driver.marker(PLAIN)).toBe('•');
    expect(driver.marker(FOLLOWER)).toBe('•');
    driver.undo();
    expectOffList(driver, PLAIN);
    expectOffList(driver, FOLLOWER);
    expect(driver.definitions()).toEqual([]);
    expect(driver.canUndo()).toBe(false);
  }, timeout);

  test(`${path}: between a created list above and a source list below of that kind a paragraph joins the one above`, async () => {
    const driver = await open('A');
    driver.toggle('bullet', NESTED_DECIMAL);
    const above = keyOf(driver, NESTED_DECIMAL);
    expect(above).toBeDefined();
    driver.toggle('bullet', DECIMAL_ITEM);
    expect(keyOf(driver, DECIMAL_ITEM)).toBe(above);
    expect(driver.definitions()).toHaveLength(1);
  }, timeout);

  test(`${path}: between a source list above and a created list below of that kind a paragraph joins the one above`, async () => {
    const driver = await open('B', { after: [FOLLOWER] });
    driver.toggle('bullet', FOLLOWER);
    expect(keyOf(driver, FOLLOWER)).toBeDefined();
    driver.toggle('bullet', PLAIN);
    const properties = propertiesOf(driver, PLAIN);
    expect(properties.numPr).toMatchObject({ numId: 1, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
    const bytes = await driver.save();
    expect(savedList(bytes, PLAIN)).toMatchObject({ paragraphs: 1, numId: '1', numFmt: 'bullet' });
    expect(savedList(bytes, FOLLOWER).numId).not.toBe('1');
  }, timeout);

  test(`${path}: a paragraph on no list joining a source list takes the level of the item above`, async () => {
    const driver = await open('A');
    driver.toggle('numbered', RESTARTED_ITEM);
    expectOffList(driver, RESTARTED_ITEM);
    driver.toggle('bullet', RESTARTED_ITEM);
    const properties = propertiesOf(driver, RESTARTED_ITEM);
    expect(properties.numPr).toMatchObject({ numId: 2, ilvl: 1 });
    expect(properties.listDef ?? null).toBeNull();
  }, timeout);

  test(`${path}: an item switched onto a source list of the other kind keeps its own level`, async () => {
    const driver = await open('A');
    driver.toggle('bullet', RESTARTED_ITEM);
    const properties = propertiesOf(driver, RESTARTED_ITEM);
    expect(properties.numPr).toMatchObject({ numId: 2, ilvl: 0 });
    expect(properties.listDef ?? null).toBeNull();
  }, timeout);

  test(`${path}: numbering over a bullet list keeps the level of every item`, async () => {
    const items = [PLAIN, FOLLOWER, SECOND_FOLLOWER, THIRD_FOLLOWER];
    const driver = await open('C', { after: items.slice(1) });
    driver.toggle('bullet', PLAIN, THIRD_FOLLOWER);
    const bullet = keyOf(driver, PLAIN);
    for (const text of [FOLLOWER, SECOND_FOLLOWER, THIRD_FOLLOWER, THIRD_FOLLOWER]) driver.indent('increase', text);
    driver.toggle('numbered', PLAIN, THIRD_FOLLOWER);
    expect(items.map(text => (propertiesOf(driver, text).numPr as { ilvl?: number } | undefined)?.ilvl)).toEqual([0, 1, 1, 2]);
    const key = keyOf(driver, PLAIN);
    expect(key).toBeDefined();
    expect(key).not.toBe(bullet);
    expect(items.map(text => keyOf(driver, text))).toEqual(items.map(() => key));
    expect(driver.definitions().find(definition => definition.key === key)?.kind).toBe('numbered');
    expect(items.map(text => driver.marker(text))).toEqual(['1.', 'a.', 'b.', 'i.']);
  }, timeout);

  test(`${path}: numbering on a level 2 bullet item keeps it at level 2`, async () => {
    const driver = await open('C');
    driver.toggle('bullet', PLAIN);
    driver.indent('increase', PLAIN);
    driver.indent('increase', PLAIN);
    driver.toggle('numbered', PLAIN);
    expect(propertiesOf(driver, PLAIN).numPr).toMatchObject({ ilvl: 2 });
    expect(keyOf(driver, PLAIN)).toBeDefined();
    expect(driver.marker(PLAIN)).toBe('i.');
  }, timeout);

  test(`${path}: numbering below a heading numbered by its style creates a list instead of joining the heading's`, async () => {
    const driver = await open('D');
    const heading = propertiesOf(driver, HEADING).numPr;
    expect(heading).toMatchObject({ numId: 1 });
    driver.toggle('numbered', PLAIN);
    await expectOnCreatedList(driver, 'D', PLAIN, 'numbered');
    expect(driver.definitions()).toHaveLength(1);
    expect(propertiesOf(driver, HEADING).numPr).toEqual(heading);
    expect(propertiesOf(driver, NEXT_HEADING).numPr).toEqual(heading);
  }, timeout);

  test(`${path}: bullets on a heading numbered by its style save a direct list of its own and leave the next heading on the style's list`, async () => {
    const driver = await open('D');
    driver.toggle('bullet', HEADING);
    const key = keyOf(driver, HEADING);
    expect(key).toBeDefined();
    expect(driver.marker(HEADING)).toBe('•');
    const bytes = await driver.save();
    const heading = savedList(bytes, HEADING);
    expect(heading.numId ?? '').toMatch(/^\d+$/);
    expect(SOURCE_NUM_IDS.D.includes(heading.numId ?? '')).toBe(false);
    expect(heading).toMatchObject({ paragraphs: 1, name: `betteroffice-list:${key}`, numFmt: 'bullet' });
    expect(savedList(bytes, NEXT_HEADING)).toEqual({ paragraphs: 1, numId: undefined });
  }, timeout);

  test(`${path}: a paragraph joining a created list takes the level of the item above`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('numbered', PLAIN);
    driver.indent('increase', PLAIN);
    driver.toggle('numbered', FOLLOWER);
    expect(propertiesOf(driver, FOLLOWER).numPr).toMatchObject({ ilvl: 1 });
    expect(keyOf(driver, FOLLOWER)).toBe(keyOf(driver, PLAIN));
    expect(driver.marker(PLAIN)).toBe('a.');
    expect(driver.marker(FOLLOWER)).toBe('b.');
  }, timeout);

  test(`${path}: a click on a selection whose first paragraph is on a list of that kind takes every paragraph off`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('bullet', PLAIN);
    expect(keyOf(driver, PLAIN)).toBeDefined();
    driver.toggle('bullet', PLAIN, FOLLOWER);
    expectOffList(driver, PLAIN);
    expectOffList(driver, FOLLOWER);
  }, timeout);

  test(`${path}: a click on a selection whose first paragraph is on no list puts every paragraph on one`, async () => {
    const driver = await open('C', { after: [FOLLOWER] });
    driver.toggle('bullet', FOLLOWER);
    const key = keyOf(driver, FOLLOWER);
    expect(key).toBeDefined();
    driver.toggle('bullet', PLAIN, FOLLOWER);
    expect(keyOf(driver, PLAIN)).toBe(key);
    expect(keyOf(driver, FOLLOWER)).toBe(key);
    expect(driver.definitions()).toHaveLength(1);
  }, timeout);
}
