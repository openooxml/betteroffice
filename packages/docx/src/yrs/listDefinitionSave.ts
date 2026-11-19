import type { YrsSession } from './index';
import { sessionMaterializedSource } from './sessionInternals';

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
/** The faces of the bullets of Word's default bullet list. */
const BULLET_FONTS: ReadonlyMap<string, string> = new Map([['', 'Symbol'], ['o', 'Courier New'], ['', 'Wingdings']]);

/** The ids the saves of each session gave the list definitions it created, so later saves give the same. */
const assigned = new WeakMap<YrsSession, ReadonlyMap<string, DocxListDefinitionIds>>();

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
  let numId = Math.max(0, ...source.numIds, ...given.map(ids => ids.numId)) + 1;
  let abstractNumId = Math.max(-1, ...source.abstractNumIds, ...given.map(ids => ids.abstractNumId)) + 1;
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

/**
 * The list definitions the editor created in `session`, with the ids its saves give them: the
 * same on every save of the session, above the ids of its source package.
 */
export function sessionListDefinitions(session: YrsSession): DocxSavedListDefinition[] {
  const definitions = session.createdListDefinitions?.() ?? [];
  if (definitions.length === 0) return [];
  const numbering = sessionMaterializedSource(session)?.package.numbering;
  const previous = assigned.get(session) ?? new Map<string, DocxListDefinitionIds>();
  const ids = savedListDefinitionIds(
    {
      numIds: numbering?.nums.map(num => num.numId) ?? [],
      abstractNumIds: numbering?.abstractNums.map(abstract => abstract.abstractNumId) ?? [],
    },
    definitions.map(definition => definition.key),
    previous
  );
  assigned.set(session, new Map([...previous, ...ids]));
  return definitions.flatMap(definition => {
    const saved = ids.get(definition.key);
    return saved ? [{ ...definition, ...saved }] : [];
  });
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value);

/**
 * The attributes of a paragraph as saved: on a list the editor created, its `numPr` carries the
 * `w:numId` saves give that list instead of the one it holds for earlier versions.
 */
export function savedListProperties(session: YrsSession, properties: Record<string, unknown>): Record<string, unknown> {
  const { listDef, numPr } = properties;
  if (!isRecord(listDef) || !isRecord(numPr) || typeof listDef.key !== 'string') return properties;
  if (typeof listDef.numId !== 'number' || listDef.numId !== Number(numPr.numId)) return properties;
  const key = listDef.key;
  const saved = sessionListDefinitions(session).find(definition => definition.key === key);
  return saved ? { ...properties, numPr: { ...numPr, numId: saved.numId } } : properties;
}

const decode = (bytes: Uint8Array | undefined) => (bytes === undefined ? undefined : new TextDecoder().decode(bytes));
const escapeXml = (value: string) =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

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
  const levels = definition.levelTexts.map((value, level) => {
    const format = definition.levelFormats[level] ?? 'decimal';
    const right = format === 'lowerRoman';
    const font = definition.kind === 'bullet' ? BULLET_FONTS.get(value) : undefined;
    const fonts = font === undefined ? '' : `<${w}rPr><${w}rFonts ${w}ascii="${font}" ${w}hAnsi="${font}" ${w}hint="default"/></${w}rPr>`;
    return `<${w}lvl ${w}ilvl="${level}"><${w}start ${w}val="1"/><${w}numFmt ${w}val="${format}"/>` +
      `<${w}lvlText ${w}val="${escapeXml(value)}"/><${w}lvlJc ${w}val="${right ? 'right' : 'left'}"/>` +
      `<${w}pPr><${w}ind ${w}left="${720 * (level + 1)}" ${w}hanging="${right ? 180 : 360}"/></${w}pPr>${fonts}</${w}lvl>`;
  });
  return `<${w}abstractNum${declaration} ${w}abstractNumId="${definition.abstractNumId}">` +
    `<${w}multiLevelType ${w}val="hybridMultilevel"/>${levels.join('')}</${w}abstractNum>`;
}

const numXml = (w: string, declaration: string, definition: DocxSavedListDefinition) =>
  `<${w}num${declaration} ${w}numId="${definition.numId}"><${w}abstractNumId ${w}val="${definition.abstractNumId}"/></${w}num>`;

/**
 * Numbering part `xml` with `definitions` after its own, in schema order: the `w:abstractNum`
 * elements after its last `w:abstractNum`, the `w:num` elements after its last `w:num`.
 */
function withListDefinitions(xml: string, definitions: readonly DocxSavedListDefinition[]): string {
  const root = /<(?:([\w.-]+):)?numbering\b[^>]*>/.exec(xml);
  if (!root) throw new Error('The numbering part has no numbering element');
  const own = root[1] === undefined ? '' : `${root[1]}:`;
  const bound = [...root[0].matchAll(/\sxmlns:([\w.-]+)\s*=\s*(["'])(.*?)\2/g)].find(match => match[3] === W)?.[1];
  const w = bound === undefined ? 'w:' : `${bound}:`;
  const declaration = bound === undefined ? ` xmlns:w="${W}"` : '';
  const held = new Set([...xml.matchAll(/:(numId|abstractNumId)="(\d+)"/g)].map(match => `${match[1]}=${match[2]}`));
  if (definitions.some(definition => held.has(`numId=${definition.numId}`) || held.has(`abstractNumId=${definition.abstractNumId}`))) {
    throw new Error('A list the editor created takes an id the numbering part already holds');
  }
  const abstracts = definitions.map(definition => abstractNumXml(w, declaration, definition)).join('');
  const nums = definitions.map(definition => numXml(w, declaration, definition)).join('');
  if (root[0].endsWith('/>')) {
    const open = `${root[0].slice(0, -2).trimEnd()}>`;
    return `${xml.slice(0, root.index)}${open}${abstracts}${nums}</${own}numbering>${xml.slice(root.index + root[0].length)}`;
  }
  const close = xml.lastIndexOf(`</${own}numbering>`);
  if (close < 0) throw new Error('The numbering part has no numbering end');
  const after = (name: string) => {
    const at = xml.lastIndexOf(`</${own}${name}>`);
    return at < 0 ? undefined : at + `</${own}${name}>`.length;
  };
  const before = (name: string) => {
    const pattern = new RegExp(`<${own.replace(/\./g, '\\.')}${name}[\\s/>]`, 'g');
    pattern.lastIndex = root.index + root[0].length;
    const found = pattern.exec(xml);
    return found && found.index < close ? found.index : undefined;
  };
  const abstractAt = after('abstractNum') ?? before('num') ?? before('numIdMacAtCleanup') ?? close;
  const numAt = Math.max(abstractAt, after('num') ?? before('numIdMacAtCleanup') ?? close);
  return `${xml.slice(0, abstractAt)}${abstracts}${xml.slice(abstractAt, numAt)}${nums}${xml.slice(numAt)}`;
}

/** The main document part of the package `parts`. */
function mainDocumentPart(parts: Readonly<Record<string, Uint8Array>>): string {
  const relationships = decode(parts['_rels/.rels'])?.match(/<Relationship\b[^>]*>/g) ?? [];
  const target = relationships
    .find(entry => /\bType="[^"]*\/officeDocument"/.test(entry))
    ?.match(/\bTarget="([^"]*)"/)?.[1];
  return target === undefined ? 'word/document.xml' : target.replace(/^\//, '');
}

/**
 * Writes `definitions`, the list definitions the editor created, into the saved package `parts`:
 * after the definitions of the numbering part the document has, otherwise into a new numbering
 * part with its relationship and content type.
 */
export function writeListDefinitionParts(parts: Record<string, Uint8Array>, definitions: readonly DocxSavedListDefinition[]): void {
  const main = mainDocumentPart(parts);
  const folder = main.slice(0, main.lastIndexOf('/') + 1);
  const relationshipsPath = `${folder}_rels/${main.slice(folder.length)}.rels`;
  const relationships = decode(parts[relationshipsPath]) ?? `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS}"></Relationships>`;
  const numbering = (relationships.match(/<Relationship\b[^>]*>/g) ?? [])
    .find(entry => /\bType="[^"]*\/numbering"/.test(entry) && !/\bTargetMode="External"/.test(entry));
  const target = numbering?.match(/\bTarget="([^"]*)"/)?.[1];
  const path = target === undefined ? `${folder}numbering.xml` : target.startsWith('/') ? target.slice(1) : `${folder}${target}`;
  const source = decode(parts[path]);
  const encoder = new TextEncoder();
  parts[path] = encoder.encode(withListDefinitions(source ?? `${XML_DECLARATION}<w:numbering xmlns:w="${W}"></w:numbering>`, definitions));
  if (numbering === undefined) {
    const taken = new Set([...relationships.matchAll(/\bId="([^"]*)"/g)].map(match => match[1]));
    let next = 1;
    while (taken.has(`rId${next}`)) next += 1;
    const entry = `<Relationship Id="rId${next}" Type="${NUMBERING_RELATIONSHIP}" Target="${path.slice(folder.length)}"/>`;
    parts[relationshipsPath] = encoder.encode(appendChild(relationships, 'Relationships', entry));
  }
  const types = decode(parts['[Content_Types].xml']);
  if (source === undefined && types !== undefined && !types.includes(`PartName="/${path}"`)) {
    const entry = `<Override PartName="/${path}" ContentType="${NUMBERING_CONTENT_TYPE}"/>`;
    parts['[Content_Types].xml'] = encoder.encode(appendChild(types, 'Types', entry));
  }
}
