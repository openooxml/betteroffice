import JSZip from 'jszip';
import type { TablePayload, YrsSession } from '@betteroffice/docx/yrs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const LIMIT = 50;
const STRICT: Record<string, string> = {
  'http://purl.oclc.org/ooxml/wordprocessingml/main': W,
  'http://purl.oclc.org/ooxml/officeDocument/relationships': R,
  'http://purl.oclc.org/ooxml/officeDocument/math': M,
};

/** Model counts of one package part, from the stories rooted in it. */
export interface DocxPartCounts {
  paragraphs: number;
  tables: number;
  rows: number;
  cells: number;
  drawings: number;
  notes: number;
  controls: number;
  fields: Record<string, number>;
}

/** Model counts the structural comparison uses to tell edits apart from save losses. */
export interface DocxModelCounts {
  comments: number;
  parts: Record<string, DocxPartCounts>;
}

export interface StructureLoss {
  part: string;
  kind: 'part' | 'element' | 'field' | 'relationship' | 'reference' | 'contentType';
  name: string;
  source: number;
  saved: number;
  explained?: number;
}

export interface StructureReport {
  ok: boolean;
  /** One line for the reader; names unattributed drops even when `ok` is true. */
  summary: string;
  lost: StructureLoss[];
  unattributed: StructureLoss[];
  counts: { part: string; element: string; source: number; saved: number }[];
  parts: { source: number; saved: number };
  fields: { source: number; saved: number };
  drawings: { source: number; saved: number };
  relationships: { source: number; saved: number };
  truncated?: true;
  ms: number;
}

const DRAWING_EMBEDS = new Set(['image', 'shape', 'chart', 'opaqueDrawing', 'horizontalRule', 'opaque']);
const add = (map: Record<string, number>, key: string, by = 1) => { map[key] = (map[key] ?? 0) + by; };

/** Collapses whitespace between field instruction tokens; quoted arguments keep theirs. */
export function normalizeInstruction(instruction: string) {
  let result = '';
  let quoted = false;
  let space = false;
  for (const char of instruction.trim()) {
    if (char === '"') quoted = !quoted;
    if (!quoted && /\s/.test(char)) { space = true; continue; }
    if (space && result) result += ' ';
    space = false;
    result += char;
  }
  return result;
}

const emptyPart = (): DocxPartCounts => ({ paragraphs: 0, tables: 0, rows: 0, cells: 0, drawings: 0, notes: 0, controls: 0, fields: {} });
const EMBEDS = new Set(['table', 'field', 'noteRef', 'sdt', 'blockSdt', ...DRAWING_EMBEDS]);
const MAIN_PART = 'word/document.xml';

function storyParts(session: YrsSession) {
  const parts = new Map<string, string>();
  let stories: ReturnType<YrsSession['listStories']>['stories'] = [];
  try { stories = session.listStories().stories; } catch { return parts; }
  const info = new Map(stories.map(story => [story.story, story]));
  for (const story of stories) {
    const root = info.get(story.root) ?? story;
    const part = root.kind === 'body' ? MAIN_PART : root.kind === 'footnote' ? 'word/footnotes.xml' : root.kind === 'endnote' ? 'word/endnotes.xml' : root.part?.replace(/^\//, '');
    if (part) parts.set(story.story, part);
  }
  return parts;
}

export function docxModelCounts(session: YrsSession): DocxModelCounts {
  const result: DocxModelCounts = { comments: session.listComments().length, parts: {} };
  const owners = storyParts(session);
  let counts = emptyPart();
  const embed = (kind: string, payload: Record<string, unknown>) => {
    if (kind === 'table') {
      counts.tables++;
      const rows = (payload as unknown as TablePayload).rows ?? [];
      counts.rows += rows.length;
      for (const row of rows) counts.cells += row.cells.length;
      return;
    }
    if (kind === 'noteRef') counts.notes++;
    else if (kind === 'sdt' || kind === 'blockSdt') counts.controls++;
    else if (DRAWING_EMBEDS.has(kind)) counts.drawings++;
    visit(payload);
  };
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) { value.forEach(visit); return; }
    if (!value || typeof value !== 'object') return;
    const object = value as Record<string, unknown>;
    const kind = typeof object.embedKind === 'string' ? object.embedKind : typeof object.kind === 'string' && EMBEDS.has(object.kind) ? object.kind : undefined;
    if (kind && object.payload && typeof object.payload === 'object') { embed(kind, object.payload as Record<string, unknown>); return; }
    if (object.modelKind === 'field' && typeof object.instruction === 'string') add(counts.fields, normalizeInstruction(object.instruction));
    for (const [key, child] of Object.entries(object)) if (key !== 'resultProjection' && key !== 'rows') visit(child);
  };
  for (const story of session.storyIds()) {
    const part = owners.get(story) ?? (story === 'body' ? MAIN_PART : '*');
    counts = result.parts[part] ??= emptyPart();
    for (const segment of session.storySegments(story)) {
      if (segment.kind === 'pilcrow') counts.paragraphs++;
      else if (segment.kind === 'embed') embed(segment.embedKind, segment.payload);
    }
  }
  return result;
}

interface Part {
  elements: Map<string, number>;
  modeled: Map<string, number>;
  fallback: Map<string, number>;
  fields: Record<string, number>;
  references: Set<string>;
  label: Map<string, string>;
}

const TAG = /<(\/?)([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
const ATTRIBUTE = /([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)\s*=\s*("([^"]*)"|'([^']*)')/g;
const decode = (value: string) => value.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (_, entity: string) =>
  entity === 'lt' ? '<' : entity === 'gt' ? '>' : entity === 'amp' ? '&' : entity === 'quot' ? '"' : entity === 'apos' ? "'"
    : String.fromCodePoint(entity[1] === 'x' ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10)));
const FALLBACK = `{${MC}}Fallback`;

function scan(xml: string): Part {
  const part: Part = { elements: new Map(), modeled: new Map(), fallback: new Map(), fields: {}, references: new Set(), label: new Map() };
  const scopes: Map<string, string>[] = [new Map([['xml', 'http://www.w3.org/XML/1998/namespace']])];
  const open: string[] = [];
  const resolve = (prefix: string) => {
    for (let index = scopes.length - 1; index >= 0; index--) { const uri = scopes[index].get(prefix); if (uri !== undefined) return STRICT[uri] ?? uri; }
    return `prefix:${prefix}`;
  };
  const complex: { code: string; recorded: boolean }[] = [];
  let fallback = 0;
  let instruction = false;
  let last = 0;
  for (const match of xml.matchAll(TAG)) {
    const [, closing, name, body, empty] = match;
    const index = match.index!;
    if (instruction && complex.length) complex[complex.length - 1].code += decode(xml.slice(last, index));
    last = index + match[0].length;
    if (closing) {
      const key = open.pop();
      if (key === `{${W}}instrText`) instruction = false;
      if (key === FALLBACK) fallback--;
      scopes.pop();
      continue;
    }
    const scope = new Map<string, string>();
    const attributes: [string, string][] = [];
    for (const attribute of body.matchAll(ATTRIBUTE)) {
      const value = decode(attribute[3] ?? attribute[4] ?? '');
      if (attribute[1] === 'xmlns') scope.set('', value);
      else if (attribute[1].startsWith('xmlns:')) scope.set(attribute[1].slice(6), value);
      else attributes.push([attribute[1], value]);
    }
    scopes.push(scope);
    const colon = name.indexOf(':');
    const key = `{${resolve(colon < 0 ? '' : name.slice(0, colon))}}${colon < 0 ? name : name.slice(colon + 1)}`;
    part.elements.set(key, (part.elements.get(key) ?? 0) + 1);
    const pool = fallback ? part.fallback : part.modeled;
    pool.set(key, (pool.get(key) ?? 0) + 1);
    if (!part.label.has(key)) part.label.set(key, name);
    for (const [attribute, value] of attributes) {
      const split = attribute.indexOf(':');
      if (split > 0 && resolve(attribute.slice(0, split)) === R && value) part.references.add(value);
    }
    if (key === `{${W}}fldSimple`) {
      const instr = attributes.find(([attribute]) => attribute.endsWith(':instr') || attribute === 'instr')?.[1];
      if (instr !== undefined && !fallback) add(part.fields, normalizeInstruction(instr));
    } else if (key === `{${W}}fldChar` && !fallback) {
      const type = attributes.find(([attribute]) => attribute.endsWith(':fldCharType'))?.[1];
      if (type === 'begin') complex.push({ code: '', recorded: false });
      else if (type === 'separate' || type === 'end') {
        const field = type === 'end' ? complex.pop() : complex[complex.length - 1];
        if (field && !field.recorded) { add(part.fields, normalizeInstruction(field.code)); field.recorded = true; }
      }
    } else if (key === `{${W}}instrText` && !fallback) instruction = !empty;
    if (key === FALLBACK && !empty) fallback++;
    if (empty) scopes.pop();
    else open.push(key);
  }
  return part;
}

interface Relationship { id: string; type: string; target: string; external: boolean }

function relationships(xml: string): Relationship[] {
  const result: Relationship[] = [];
  for (const match of xml.matchAll(/<(?:\w+:)?Relationship\b((?:[^>"']|"[^"]*"|'[^']*')*)\/?>/g)) {
    const values = new Map([...match[1].matchAll(ATTRIBUTE)].map(attribute => [attribute[1], decode(attribute[3] ?? attribute[4] ?? '')]));
    result.push({ id: values.get('Id') ?? '', type: values.get('Type') ?? '', target: values.get('Target') ?? '', external: values.get('TargetMode') === 'External' });
  }
  return result;
}

const ownerOf = (rels: string) => {
  const slash = rels.lastIndexOf('/_rels/');
  const base = slash < 0 ? '' : rels.slice(0, slash + 1);
  const file = rels.slice(slash < 0 ? rels.indexOf('_rels/') + 6 : slash + 7, -'.rels'.length);
  return base + file;
};

function resolveTarget(owner: string, target: string) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = owner.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop();
    else if (segment && segment !== '.') parts.push(segment);
  }
  return parts.join('/');
}

function contentTypes(xml: string | undefined) {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const match of (xml ?? '').matchAll(/<(?:\w+:)?(Default|Override)\b((?:[^>"']|"[^"]*"|'[^']*')*)\/?>/g)) {
    const values = new Map([...match[2].matchAll(ATTRIBUTE)].map(attribute => [attribute[1], decode(attribute[3] ?? attribute[4] ?? '')]));
    const type = values.get('ContentType')?.trim();
    if (!type) continue;
    if (match[1] === 'Default') defaults.set((values.get('Extension') ?? '').toLowerCase(), type);
    else overrides.set((values.get('PartName') ?? '').replace(/^\//, '').toLowerCase(), type);
  }
  return (name: string) => overrides.get(name.toLowerCase()) ?? defaults.get(name.slice(name.lastIndexOf('.') + 1).toLowerCase());
}

async function readPackage(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const files = Object.values(zip.files).filter(file => !file.dir);
  const xml = new Map<string, string>();
  for (const file of files) if (/\.(xml|rels|vml)$/i.test(file.name)) xml.set(file.name, await file.async('string'));
  return { names: new Set(files.map(file => file.name)), xml };
}

const COMMENT_PART = /^word\/comments[A-Za-z]*\.xml$/;
const count = (part: Part | undefined, keys: readonly string[], pool: 'modeled' | 'fallback' = 'modeled') => keys.reduce((sum, key) => sum + (part?.[pool].get(`{${W}}${key}`) ?? 0), 0);
const total = (parts: Map<string, Part>, keys: readonly string[], all = false) => {
  let sum = 0;
  for (const [name, part] of parts) if (all || !COMMENT_PART.test(name)) sum += count(part, keys);
  return sum;
};

type Pick = (counts: DocxPartCounts) => number;
const NET: [readonly string[], Pick][] = [
  [['p'], counts => counts.paragraphs],
  [['tbl'], counts => counts.tables],
  [['tr'], counts => counts.rows],
  [['tc'], counts => counts.cells],
  [['drawing', 'pict', 'object'], counts => counts.drawings],
  [['footnoteReference', 'endnoteReference'], counts => counts.notes],
  [['sdt'], counts => counts.controls],
];
const relsOf = (pkg: { xml: Map<string, string> }) => new Map([...pkg.xml].filter(([name]) => name.endsWith('.rels')).map(([name, xml]) => [ownerOf(name), relationships(xml)]));
const NOTE_PARTS: Record<string, string> = { footnotes: 'word/footnotes.xml', endnotes: 'word/endnotes.xml' };
/** Removing these can take content the model holds only as opaque data: text boxes, nested fields, merge continuations. */
const CONTAINERS: Pick[] = [counts => counts.tables, counts => counts.rows, counts => counts.cells];
const COMMENT_KEYS = ['commentRangeStart', 'commentRangeEnd', 'commentReference', 'comment'];
const PROTECTED = new Set([
  ...['bookmarkStart', 'bookmarkEnd', 'sectPr', 'hyperlink', 'pPrChange', 'rPrChange', 'sectPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange',
    'numberingChange', 'ins', 'del', 'moveFrom', 'moveTo', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
    'permStart', 'permEnd', 'smartTag', 'customXml', 'ruby', 'noProof', 'specVanish', 'footnote', 'endnote', 'txbxContent', 'altChunk',
    'headerReference', 'footerReference'].map(local => `{${W}}${local}`),
  `{${M}}oMath`, `{${M}}oMathPara`, `{${MC}}AlternateContent`,
]);

/**
 * Compares the saved package with its source, part by part. A loss counts only beyond what the
 * edits explain: the change the session model shows between open and save for the same kind.
 */
export async function compareDocxStructure(source: Uint8Array, saved: Uint8Array, before: DocxModelCounts, after: DocxModelCounts): Promise<StructureReport> {
  const started = performance.now();
  const [from, to] = await Promise.all([readPackage(source), readPackage(saved)]);
  const lost: StructureLoss[] = [];
  const unattributed: StructureLoss[] = [];
  const counts: StructureReport['counts'] = [];
  const scanned = (pkg: typeof from) => new Map([...pkg.xml].filter(([name]) => !name.endsWith('.rels') && name !== '[Content_Types].xml').map(([name, xml]) => [name, scan(xml)]));
  const sourceParts = scanned(from);
  const savedParts = scanned(to);
  const sourceRels = relsOf(from);
  const savedRels = relsOf(to);
  const aliases = new Map<string, string>();
  const kind = (relationship: Relationship) => relationship.type.slice(relationship.type.lastIndexOf('/') + 1);
  const main = sourceRels.get('')?.find(relationship => kind(relationship) === 'officeDocument' && !relationship.external);
  if (main) {
    const name = resolveTarget('', main.target);
    aliases.set(name, MAIN_PART);
    for (const relationship of sourceRels.get(name) ?? []) {
      const canonical = NOTE_PARTS[kind(relationship)];
      if (canonical && !relationship.external) aliases.set(resolveTarget(name, relationship.target), canonical);
    }
  }
  const partNames = new Set([...sourceParts.keys()].filter(name => !COMMENT_PART.test(name)));
  const model = (counts: DocxModelCounts, name: string) => counts.parts[aliases.get(name) ?? name] ?? emptyPart();
  const removedComments = Math.max(0, before.comments - after.comments);
  const structural = removedComments > 0 || [...new Set([...partNames, ...Object.keys(before.parts)])].some(name =>
    NET.some(([, pick]) => pick(model(before, name)) > pick(model(after, name))) ||
    Object.entries(model(before, name).fields).some(([key, value]) => value > (model(after, name).fields[key] ?? 0)));

  for (const name of from.names) {
    if (to.names.has(name)) continue;
    const comments = COMMENT_PART.test(name) && after.comments === 0;
    const rels = name.endsWith('.rels') && !to.names.has(ownerOf(name)) && from.names.has(ownerOf(name));
    if (!comments && !rels) lost.push({ part: name, kind: 'part', name, source: 1, saved: 0 });
  }

  for (const name of partNames) {
    const was = sourceParts.get(name)!;
    const is = savedParts.get(name);
    const report = CONTAINERS.some(pick => pick(model(before, name)) > pick(model(after, name))) ? unattributed : lost;
    for (const [keys, pick] of NET) {
      const explained = pick(model(before, name)) - pick(model(after, name));
      const source = count(was, keys);
      const saved = count(is, keys);
      if (source - saved > explained) report.push({ part: name, kind: 'element', name: keys.map(key => `w:${key}`).join('|'), source, saved, explained });
    }
    const drawingKeys = NET[4][0];
    const fallbackDrop = count(was, drawingKeys, 'fallback') - count(is, drawingKeys, 'fallback');
    const drawingsRemoved = Math.max(0, model(before, name).drawings - model(after, name).drawings);
    if (fallbackDrop > drawingsRemoved) report.push({ part: name, kind: 'element', name: 'mc:Fallback w:drawing|w:pict|w:object', source: count(was, drawingKeys, 'fallback'), saved: count(is, drawingKeys, 'fallback'), explained: drawingsRemoved });
    const beforeFields = model(before, name).fields;
    const afterFields = model(after, name).fields;
    for (const [instruction, number] of Object.entries(was.fields)) {
      const kept = is?.fields[instruction] ?? 0;
      const explained = Math.max(0, (beforeFields[instruction] ?? 0) - (afterFields[instruction] ?? 0));
      if (number - kept > explained) report.push({ part: name, kind: 'field', name: instruction, source: number, saved: kept, explained });
    }
  }
  for (const [name, part] of sourceParts) {
    if (!COMMENT_PART.test(name)) continue;
    const next = savedParts.get(name);
    for (const [instruction, number] of Object.entries(part.fields)) {
      const kept = next?.fields[instruction] ?? 0;
      if (number > kept) (removedComments ? unattributed : lost).push({ part: name, kind: 'field', name: instruction, source: number, saved: kept });
    }
  }
  for (const key of COMMENT_KEYS) {
    const source = total(sourceParts, [key], key === 'comment');
    const saved = total(savedParts, [key], key === 'comment');
    if (source - saved > removedComments) (structural ? unattributed : lost).push({ part: '*', kind: 'element', name: `w:${key}`, source, saved, explained: removedComments });
  }

  const fieldTotals = (parts: Map<string, Part>) => { const all: Record<string, number> = {}; for (const part of parts.values()) for (const [key, value] of Object.entries(part.fields)) add(all, key, value); return all; };
  const sourceFields = fieldTotals(sourceParts);
  const savedFields = fieldTotals(savedParts);

  for (const [name, part] of sourceParts) {
    const next = savedParts.get(name);
    if (!next) continue;
    for (const [key, count] of part.elements) {
      const kept = next.elements.get(key) ?? 0;
      if (kept >= count) continue;
      const label = part.label.get(key)!;
      if (PROTECTED.has(key)) (structural ? unattributed : lost).push({ part: name, kind: 'element', name: label, source: count, saved: kept });
      else counts.push({ part: name, element: label, source: count, saved: kept });
    }
  }

  const sourceTypes = contentTypes(from.xml.get('[Content_Types].xml'));
  const savedTypes = contentTypes(to.xml.get('[Content_Types].xml'));
  for (const name of to.names) {
    if (name === '[Content_Types].xml') continue;
    const type = savedTypes(name);
    const was = from.names.has(name) ? sourceTypes(name) : undefined;
    if (type === was || (type && !from.names.has(name))) continue;
    if (!type || was) lost.push({ part: name, kind: 'contentType', name: `${name} ${was ?? 'none'} -> ${type ?? 'none'}`, source: 1, saved: 0 });
  }

  let sourceRelationships = 0;
  let savedRelationships = 0;
  const byId = (list: readonly Relationship[]) => new Map(list.map(relationship => [relationship.id, relationship]));
  for (const [owner, list] of sourceRels) {
    sourceRelationships += list.length;
    if (!to.names.has(owner) && owner !== '') continue;
    const next = byId(savedRels.get(owner) ?? []);
    const references = savedParts.get(owner)?.references;
    for (const relationship of list) {
      const kept = next.get(relationship.id);
      if (kept && kept.type === relationship.type && kept.target === relationship.target && kept.external === relationship.external) continue;
      if (!kept && references && !references.has(relationship.id) && sourceParts.get(owner)?.references.has(relationship.id)) continue;
      lost.push({ part: owner || '/', kind: 'relationship', name: `${relationship.id} ${relationship.type.slice(relationship.type.lastIndexOf('/') + 1)} ${relationship.target}${kept ? ' changed' : ''}`, source: 1, saved: 0 });
    }
  }
  for (const [owner, list] of savedRels) {
    savedRelationships += list.length;
    const previous = byId(sourceRels.get(owner) ?? []);
    for (const relationship of list) {
      if (relationship.external || to.names.has(resolveTarget(owner, relationship.target))) continue;
      const was = previous.get(relationship.id);
      if (was && !was.external && !from.names.has(resolveTarget(owner, was.target))) continue;
      lost.push({ part: owner || '/', kind: 'relationship', name: `${relationship.id} target ${relationship.target} missing`, source: 1, saved: 0 });
    }
  }
  for (const [name, part] of savedParts) {
    const ids = new Set((savedRels.get(name) ?? []).map(relationship => relationship.id));
    const known = new Set((sourceRels.get(name) ?? []).map(relationship => relationship.id));
    const dangling = sourceParts.get(name)?.references;
    for (const reference of part.references) {
      if (ids.has(reference) || (dangling?.has(reference) && !known.has(reference))) continue;
      lost.push({ part: name, kind: 'reference', name: reference, source: 1, saved: 0 });
    }
  }

  const drawingKeys = ['drawing', 'pict', 'object'];
  const truncated = lost.length > LIMIT || unattributed.length > LIMIT || counts.length > LIMIT;
  counts.sort((left, right) => (right.source - right.saved) - (left.source - left.saved));
  const summary = [
    lost.length ? `${lost.length} structural ${lost.length === 1 ? 'loss' : 'losses'} the edits do not explain; see lost.` : '',
    unattributed.length ? `${unattributed.length} structural ${unattributed.length === 1 ? 'element' : 'elements'} dropped after edits that removed content, such as rows or tables; could not attribute them to those edits, review unattributed.` : '',
  ].filter(Boolean).join(' ') || 'Structure intact: every drop is explained by the edits.';
  return {
    ok: lost.length === 0,
    summary,
    lost: lost.slice(0, LIMIT),
    unattributed: unattributed.slice(0, LIMIT),
    counts: counts.slice(0, LIMIT),
    parts: { source: from.names.size, saved: to.names.size },
    fields: { source: Object.values(sourceFields).reduce((a, b) => a + b, 0), saved: Object.values(savedFields).reduce((a, b) => a + b, 0) },
    drawings: { source: total(sourceParts, drawingKeys), saved: total(savedParts, drawingKeys) },
    relationships: { source: sourceRelationships, saved: savedRelationships },
    ...(truncated && { truncated: true as const }),
    ms: Math.round(performance.now() - started),
  };
}
