import { headerFooterAliasOf, markHeaderFooterAlias } from '../docx/headerFooterAliasProjection';
import type { HeaderFooter, NumberingDefinitions, Paragraph } from '../types/document';
import type { YrsSession } from './index';

/** The `w:numId` and `w:abstractNumId` a save gives a list definition the editor created. */
export interface DocxListDefinitionIds {
  numId: number;
  abstractNumId: number;
}

/** A list definition the editor created, with the ids a save gives it. */
export type DocxSavedListDefinition = ReturnType<YrsSession['createdListDefinitions']>[number] & DocxListDefinitionIds;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NUMBERING_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const NUMBERING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
/** The `w:name` prefix of the `w:abstractNum` a save writes for a created list, before its key. */
const CARRIER = 'betteroffice-list:';
/** The levels OOXML defines. */
const LEVELS = 9;
/** The faces of the bullets of Word's default bullet list. */
const BULLET_FONTS: ReadonlyMap<string, string> = new Map([
  [String.fromCharCode(0xf0b7), 'Symbol'],
  ['o', 'Courier New'],
  [String.fromCharCode(0xf0a7), 'Wingdings'],
]);
/** The numbers Rust's `f64` parser reads from a string, as the engine reads attribute values. */
const ENGINE_NUMBER = /^[+-]?(?:inf|infinity|nan|(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)$/i;

/**
 * The property a projection adds to the `numPr` of a paragraph on a created list, holding the
 * list's key; `writeSessionSave` takes it off before the writer sees the document.
 */
const CREATED_LIST = 'createdList';

/** The ids the saves of each session gave the list definitions it created, so later saves give the same. */
const assigned = new WeakMap<YrsSession, ReadonlyMap<string, DocxListDefinitionIds>>();

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value);

const highest = (values: Iterable<number>, floor: number): number => {
  let result = floor;
  for (const value of values) if (value > result) result = value;
  return result;
};

/** `value` as the engine reads a number from an attribute value; `undefined` when it reads none. */
function engineNumber(value: unknown): number | undefined {
  if (typeof value === 'number') return value;
  if (typeof value === 'bigint') return Number(value);
  if (typeof value !== 'string' || !ENGINE_NUMBER.test(value)) return undefined;
  const unsigned = value.replace(/^[+-]/, '').toLowerCase();
  if (unsigned === 'nan') return Number.NaN;
  if (unsigned === 'inf' || unsigned === 'infinity') return value.startsWith('-') ? -Infinity : Infinity;
  return Number(value);
}

/**
 * The key of the list the editor created that a paragraph with the attributes `properties` is on,
 * by the engine's rule: its `listDef` names the key while `numPr.numId` equals the numId recorded
 * beside it; `undefined` otherwise.
 */
export function createdListKey(properties: Record<string, unknown>): string | undefined {
  const { numPr, listDef } = properties;
  if (!isRecord(numPr) || !isRecord(listDef)) return undefined;
  const numId = engineNumber(numPr.numId);
  if (numId === undefined || numId === 0) return undefined;
  const recorded = engineNumber(listDef.numId);
  return typeof listDef.key === 'string' && recorded === numId ? listDef.key : undefined;
}

/**
 * `paragraph`, projected from `properties`, with its `numPr` recognisable to a save as that of the
 * created list it is on; it keeps the numId earlier versions read until a save gives the list its
 * own.
 */
export function markCreatedListParagraph(paragraph: Paragraph, properties: Record<string, unknown>): Paragraph {
  const key = createdListKey(properties);
  const numPr = paragraph.formatting?.numPr;
  if (key === undefined || numPr === undefined) return paragraph;
  const marked = { ...numPr, [CREATED_LIST]: key } as typeof numPr;
  return { ...paragraph, formatting: { ...paragraph.formatting, numPr: marked } };
}

/**
 * `value`, a projected document or any part of it, with the created-list mark taken off every
 * `numPr` that carries one and, where `numIds` names its list, that list's numId in its place. The
 * objects on the way to a changed one are copied, header and footer aliases kept; every other
 * object is returned as it is.
 */
export function withCreatedListIds<T>(value: T, numIds: ReadonlyMap<string, number>): T {
  const done = new Map<object, unknown>();
  const visit = (node: unknown): unknown => {
    if (node === null || typeof node !== 'object' || ArrayBuffer.isView(node) || node instanceof ArrayBuffer) return node;
    if (done.has(node)) return done.get(node);
    done.set(node, node);
    let result: unknown = node;
    if (node instanceof Map) {
      const entries = [...node].map(([key, entry]) => [key, visit(entry)] as const);
      if (entries.some(([key, entry]) => entry !== node.get(key))) result = new Map(entries);
    } else if (Array.isArray(node)) {
      const items = node.map(visit);
      if (items.some((item, index) => item !== node[index])) result = items;
    } else if ([Object.prototype, null].includes(Object.getPrototypeOf(node))) {
      const record = node as Record<string, unknown>;
      const key = record[CREATED_LIST];
      if (typeof key === 'string') {
        const unmarked = Object.fromEntries(Object.entries(record).filter(([name]) => name !== CREATED_LIST));
        const numId = numIds.get(key);
        result = numId === undefined ? unmarked : { ...unmarked, numId };
      } else {
        const entries = Object.entries(record).map(([name, entry]) => [name, visit(entry)] as const);
        if (entries.some(([name, entry]) => entry !== record[name])) {
          const copy = Object.fromEntries(entries);
          const alias = headerFooterAliasOf(node as HeaderFooter);
          if (alias !== undefined) markHeaderFooterAlias(copy as HeaderFooter, alias);
          result = copy;
        }
      }
    }
    done.set(node, result);
    return result;
  };
  return visit(value) as T;
}

/**
 * The ids a save gives the list definitions the editor created, by key, whatever order the keys
 * come in: a definition an earlier save numbered, in `previous`, keeps its ids, and every other
 * one takes the next ids free above the source's and those already given.
 */
export function savedListDefinitionIds(
  source: { numIds: readonly number[]; abstractNumIds: readonly number[] },
  keys: readonly string[],
  previous: ReadonlyMap<string, DocxListDefinitionIds> = new Map()
): Map<string, DocxListDefinitionIds> {
  const given = [...previous.values()];
  let numId = highest([...source.numIds, ...given.map(ids => ids.numId)], 0) + 1;
  let abstractNumId = highest([...source.abstractNumIds, ...given.map(ids => ids.abstractNumId)], -1) + 1;
  const ids = new Map<string, DocxListDefinitionIds>();
  for (const key of [...new Set(keys)].sort((left, right) => (left < right ? -1 : left > right ? 1 : 0))) {
    const saved = previous.get(key);
    if (saved) {
      ids.set(key, saved);
      continue;
    }
    ids.set(key, { numId, abstractNumId });
    numId += 1;
    abstractNumId += 1;
  }
  return ids;
}

const decode = (bytes: Uint8Array | undefined) => (bytes === undefined ? undefined : new TextDecoder().decode(bytes));
const escapeXml = (value: string) =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unescapeXml = (value: string) =>
  value.replace(/&(lt|gt|quot|apos|amp|#\d+|#x[0-9a-f]+);/gi, (entity, name: string) => {
    const lower = name.toLowerCase();
    if (lower.startsWith('#')) {
      const code = lower.startsWith('#x') ? Number.parseInt(lower.slice(2), 16) : Number(lower.slice(1));
      return code <= 0x10ffff ? String.fromCodePoint(code) : entity;
    }
    return ({ lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' } as Record<string, string>)[lower] ?? entity;
  });

/** `value` without the characters XML 1.0 does not allow, unpaired surrogates included. */
function xmlText(value: string): string {
  let result = '';
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        result += value.slice(index, index + 2);
        index += 1;
      }
      continue;
    }
    if (code >= 0xdc00 && code <= 0xdfff) continue;
    if (code === 0x9 || code === 0xa || code === 0xd || (code >= 0x20 && code <= 0xfffd)) result += value[index];
  }
  return result;
}

const attributeValue = (value: string | number) => escapeXml(xmlText(String(value)));

/** `<{prefix}{name} ...>` elements of `masked`, with where they start and end. */
function elementsOf(masked: string, qualified: string): Array<{ index: number; start: string; body: string; end: number }> {
  const found: Array<{ index: number; start: string; body: string; end: number }> = [];
  for (const match of masked.matchAll(new RegExp(`<${qualified.replace(/\./g, '\\.')}(?=[\\s/>])[^>]*>`, 'g'))) {
    const index = match.index ?? 0;
    const after = index + match[0].length;
    if (match[0].endsWith('/>')) {
      found.push({ index, start: match[0], body: '', end: after });
      continue;
    }
    const close = masked.indexOf(`</${qualified}>`, after);
    const end = close < 0 ? masked.length : close + `</${qualified}>`.length;
    found.push({ index, start: match[0], body: masked.slice(after, close < 0 ? masked.length : close), end });
  }
  return found;
}

/** The value, in either quoting, of attribute `name` of the start tag `start`. */
const attributeOf = (start: string, name: string) => {
  const match = new RegExp(`\\s(?:[\\w.-]+:)?${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(start);
  const value = match?.[1] ?? match?.[2];
  return value === undefined ? undefined : unescapeXml(value);
};

const numberAttribute = (start: string, name: string) => {
  const value = attributeOf(start, name);
  return value !== undefined && /^\d+$/.test(value) ? Number(value) : undefined;
};

/** The `val` of the first `<{any prefix}{name}>` element in `xml`. */
const childValue = (xml: string, name: string) => {
  const start = new RegExp(`<(?:[\\w.-]+:)?${name}(?=[\\s/>])[^>]*>`).exec(xml)?.[0];
  return start === undefined ? undefined : attributeOf(start, 'val');
};

/** The level formats and texts of an `abstractNum` body, by level, as the writer controls them. */
function levelsOf(body: string, own: string): Array<{ format?: string; text?: string } | undefined> {
  const levels: Array<{ format?: string; text?: string } | undefined> = [];
  for (const level of elementsOf(body, `${own}lvl`)) {
    const index = numberAttribute(level.start, 'ilvl');
    if (index === undefined || index >= LEVELS) continue;
    levels[index] = { format: childValue(level.body, 'numFmt'), text: childValue(level.body, 'lvlText') };
  }
  return levels;
}

/** Every id written as attribute `name` in `xml`, in either quoting, wherever it stands. */
const looseIds = (xml: string, name: string) =>
  [...xml.matchAll(new RegExp(`(?:^|[\\s:])${name}\\s*=\\s*["'](\\d+)["']`, 'g'))].map(match => Number(match[1]));

/** What a numbering part holds: its ids, the created lists it carries and where its elements are. */
function scanNumbering(xml: string) {
  const masked = xml.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, match => ' '.repeat(match.length));
  const root = /<(?:([\w.-]+):)?numbering\b[^>]*>/.exec(masked) ?? undefined;
  const own = root?.[1] === undefined ? '' : `${root[1]}:`;
  const abstractNums = elementsOf(masked, `${own}abstractNum`);
  const nums = elementsOf(masked, `${own}num`);
  const carriers: Array<{ key: string; abstractNumId: number; levels: ReturnType<typeof levelsOf> }> = [];
  const abstractNumIds: number[] = looseIds(xml, 'abstractNumId');
  for (const element of abstractNums) {
    const id = numberAttribute(element.start, 'abstractNumId');
    if (id === undefined) continue;
    abstractNumIds.push(id);
    const value = childValue(element.body, 'name');
    if (value?.startsWith(CARRIER)) {
      carriers.push({ key: value.slice(CARRIER.length), abstractNumId: id, levels: levelsOf(element.body, own) });
    }
  }
  const pointers = new Map<number, number>();
  const numIds: number[] = looseIds(xml, 'numId');
  for (const element of nums) {
    const id = numberAttribute(element.start, 'numId');
    if (id === undefined) continue;
    numIds.push(id);
    const pointer = childValue(element.body, 'abstractNumId');
    if (pointer !== undefined && /^\d+$/.test(pointer) && !pointers.has(Number(pointer))) pointers.set(Number(pointer), id);
  }
  return { masked, root, own, abstractNums, nums, carriers, pointers, numIds, abstractNumIds };
}

/** The abstractNumId of the abstractNum `scan` carries for `definition` with what the writer writes for it. */
function carriedAbstractNum(
  scan: ReturnType<typeof scanNumbering>,
  definition: ReturnType<YrsSession['createdListDefinitions']>[number]
): number | undefined {
  const written = definition.levelTexts.slice(0, LEVELS).map((text, level) => ({
    format: xmlText(definition.levelFormats[level] ?? 'decimal'),
    text: xmlText(text),
  }));
  return scan.carriers.find(carrier =>
    carrier.key === definition.key &&
    carrier.levels.length === written.length &&
    written.every((level, index) => carrier.levels[index]?.format === level.format && carrier.levels[index]?.text === level.text)
  )?.abstractNumId;
}

/** Where the numbering part of the package `parts` is, or goes, and whether a relationship names it. */
function numberingPart(parts: Readonly<Record<string, Uint8Array>>) {
  const roots = decode(parts['_rels/.rels'])?.match(/<Relationship\b[^>]*>/g) ?? [];
  const main = (roots
    .find(entry => /\bType="[^"]*\/officeDocument"/.test(entry))
    ?.match(/\bTarget="([^"]*)"/)?.[1] ?? 'word/document.xml').replace(/^\//, '');
  const folder = main.slice(0, main.lastIndexOf('/') + 1);
  const relationshipsPath = `${folder}_rels/${main.slice(folder.length)}.rels`;
  const relationships = decode(parts[relationshipsPath]);
  const related = (relationships?.match(/<Relationship\b[^>]*>/g) ?? [])
    .find(entry => /\bType="[^"]*\/numbering"/.test(entry) && !/\bTargetMode="External"/.test(entry));
  const target = related?.match(/\bTarget="([^"]*)"/)?.[1];
  const path = target === undefined ? `${folder}numbering.xml` : target.startsWith('/') ? target.slice(1) : `${folder}${target}`;
  return { folder, relationshipsPath, relationships, related: related !== undefined, path };
}

/**
 * The list definitions the editor created, `definitions` as `session` lists them, with the ids a
 * save into the package `parts` gives them: a definition the package already carries, under its key
 * and with what the writer writes for it, keeps the ids it has there; one an earlier save of the
 * session numbered keeps those while they are above every id of the package and of `model`, the
 * numbering of the session's source; every other one takes the next ids above those.
 */
export function savedListDefinitions(
  session: YrsSession,
  definitions: ReturnType<YrsSession['createdListDefinitions']>,
  parts: Readonly<Record<string, Uint8Array>>,
  model?: NumberingDefinitions
): DocxSavedListDefinition[] {
  if (definitions.length === 0) return [];
  const scan = scanNumbering(decode(parts[numberingPart(parts).path]) ?? '');
  const numIds = [...scan.numIds, ...(model?.nums ?? []).map(num => num.numId)];
  const abstractNumIds = [...scan.abstractNumIds, ...(model?.abstractNums ?? []).map(abstract => abstract.abstractNumId)];
  const highestNumId = highest(numIds, 0);
  const highestAbstractNumId = highest(abstractNumIds, -1);
  const previous = assigned.get(session) ?? new Map<string, DocxListDefinitionIds>();
  const kept = new Map<string, DocxListDefinitionIds>();
  for (const [key, ids] of previous) {
    if (ids.numId > highestNumId && ids.abstractNumId > highestAbstractNumId) kept.set(key, ids);
  }
  const carried = new Map<string, number>();
  for (const definition of definitions) {
    const abstractNumId = carriedAbstractNum(scan, definition);
    if (abstractNumId === undefined) continue;
    carried.set(definition.key, abstractNumId);
    const numId = scan.pointers.get(abstractNumId);
    if (numId !== undefined) kept.set(definition.key, { numId, abstractNumId });
  }
  const ids = savedListDefinitionIds(
    { numIds, abstractNumIds },
    definitions.map(definition => definition.key),
    kept
  );
  assigned.set(session, new Map([...previous, ...ids]));
  return definitions.flatMap(definition => {
    const saved = ids.get(definition.key);
    if (!saved) return [];
    return [{ ...definition, ...saved, abstractNumId: carried.get(definition.key) ?? saved.abstractNumId }];
  });
}

/** `xml` with `child` as the last child of its element `name`. */
function appendChild(xml: string, name: string, child: string): string {
  const close = xml.lastIndexOf(`</${name}>`);
  if (close >= 0) return `${xml.slice(0, close)}${child}${xml.slice(close)}`;
  const empty = new RegExp(`<${name}\\b[^>]*?/>`).exec(xml);
  if (!empty) throw new Error(`The package part has no ${name} element`);
  const open = `${empty[0].slice(0, -2).trimEnd()}>`;
  return `${xml.slice(0, empty.index)}${open}${child}</${name}>${xml.slice(empty.index + empty[0].length)}`;
}

function abstractNumXml(w: string, declaration: string, definition: DocxSavedListDefinition): string {
  const levels = definition.levelTexts.slice(0, LEVELS).map((value, level) => {
    const format = definition.levelFormats[level] ?? 'decimal';
    const right = format === 'lowerRoman';
    const font = definition.kind === 'bullet' ? BULLET_FONTS.get(value) : undefined;
    const fonts = font === undefined ? '' : `<${w}rPr><${w}rFonts ${w}ascii="${font}" ${w}hAnsi="${font}" ${w}hint="default"/></${w}rPr>`;
    return `<${w}lvl ${w}ilvl="${level}"><${w}start ${w}val="1"/><${w}numFmt ${w}val="${attributeValue(format)}"/>` +
      `<${w}lvlText ${w}val="${attributeValue(value)}"/><${w}lvlJc ${w}val="${right ? 'right' : 'left'}"/>` +
      `<${w}pPr><${w}ind ${w}left="${720 * (level + 1)}" ${w}hanging="${right ? 180 : 360}"/></${w}pPr>${fonts}</${w}lvl>`;
  });
  return `<${w}abstractNum${declaration} ${w}abstractNumId="${attributeValue(definition.abstractNumId)}">` +
    `<${w}multiLevelType ${w}val="hybridMultilevel"/><${w}name ${w}val="${attributeValue(CARRIER + definition.key)}"/>` +
    `${levels.join('')}</${w}abstractNum>`;
}

const numXml = (w: string, declaration: string, definition: DocxSavedListDefinition) =>
  `<${w}num${declaration} ${w}numId="${attributeValue(definition.numId)}"><${w}abstractNumId ${w}val="${attributeValue(definition.abstractNumId)}"/></${w}num>`;

/**
 * Numbering part `xml`, or a new one when it is undefined, holding `definitions` after its own in
 * schema order: a `w:abstractNum` it does not yet carry after its last `w:abstractNum`, a `w:num`
 * it does not yet hold after its last `w:num`. At most nine levels are written, and only text XML
 * allows.
 */
export function numberingWithListDefinitions(xml: string | undefined, definitions: readonly DocxSavedListDefinition[]): string {
  const source = xml ?? `${XML_DECLARATION}<w:numbering xmlns:w="${W}"></w:numbering>`;
  const scan = scanNumbering(source);
  const root = scan.root;
  if (!root) throw new Error('The numbering part has no numbering element');
  const bound = [...root[0].matchAll(/\sxmlns:([\w.-]+)\s*=\s*(["'])(.*?)\2/g)].find(match => match[3] === W)?.[1];
  const w = bound === undefined ? 'w:' : `${bound}:`;
  const declaration = bound === undefined ? ` xmlns:w="${W}"` : '';
  const abstracts = definitions
    .filter(definition => carriedAbstractNum(scan, definition) !== definition.abstractNumId)
    .map(definition => abstractNumXml(w, declaration, definition))
    .join('');
  const nums = definitions
    .filter(definition => scan.pointers.get(definition.abstractNumId) !== definition.numId)
    .map(definition => numXml(w, declaration, definition))
    .join('');
  if (abstracts === '' && nums === '') return source;
  if (root[0].endsWith('/>')) {
    const open = `${source.slice(root.index, root.index + root[0].length - 2).trimEnd()}>`;
    return `${source.slice(0, root.index)}${open}${abstracts}${nums}</${scan.own}numbering>${source.slice(root.index + root[0].length)}`;
  }
  const close = scan.masked.lastIndexOf(`</${scan.own}numbering>`);
  if (close < 0) throw new Error('The numbering part has no numbering end');
  const cleanup = elementsOf(scan.masked, `${scan.own}numIdMacAtCleanup`)[0]?.index;
  const lastAbstractNum = scan.abstractNums[scan.abstractNums.length - 1];
  const lastNum = scan.nums[scan.nums.length - 1];
  const abstractAt = lastAbstractNum?.end ?? scan.nums[0]?.index ?? cleanup ?? close;
  const numAt = Math.max(abstractAt, lastNum?.end ?? cleanup ?? close);
  return `${source.slice(0, abstractAt)}${abstracts}${source.slice(abstractAt, numAt)}${nums}${source.slice(numAt)}`;
}

/**
 * Writes `definitions`, the list definitions the editor created with the ids this save gives them,
 * into the saved package `parts`: into the numbering part the document has, otherwise into a new
 * numbering part with its relationship and content type. Whether it changed a part.
 */
export function writeListDefinitionParts(parts: Record<string, Uint8Array>, definitions: readonly DocxSavedListDefinition[]): boolean {
  const located = numberingPart(parts);
  const source = decode(parts[located.path]);
  const written = numberingWithListDefinitions(source, definitions);
  if (written === source) return false;
  const encoder = new TextEncoder();
  parts[located.path] = encoder.encode(written);
  if (!located.related) {
    const relationships = located.relationships ?? `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS}"></Relationships>`;
    const taken = new Set([...relationships.matchAll(/\bId="([^"]*)"/g)].map(match => match[1]));
    let next = 1;
    while (taken.has(`rId${next}`)) next += 1;
    const entry = `<Relationship Id="rId${next}" Type="${NUMBERING_RELATIONSHIP}" Target="${escapeXml(located.path.slice(located.folder.length))}"/>`;
    parts[located.relationshipsPath] = encoder.encode(appendChild(relationships, 'Relationships', entry));
  }
  const types = decode(parts['[Content_Types].xml']);
  if (source === undefined && types !== undefined && !types.includes(`PartName="/${located.path}"`)) {
    const entry = `<Override PartName="/${escapeXml(located.path)}" ContentType="${NUMBERING_CONTENT_TYPE}"/>`;
    parts['[Content_Types].xml'] = encoder.encode(appendChild(types, 'Types', entry));
  }
  return true;
}
