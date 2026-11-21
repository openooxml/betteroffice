import JSZip from 'jszip';
import type { DocxCommentSaveIdentity } from '@betteroffice/docx/yrs';
import { DocumentToolError } from './types';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const COMMENTS = 'word/comments.xml';
const DOCUMENT_RELS = 'word/_rels/document.xml.rels';
const COMMENT_COMPANIONS = [
  ['word/commentsExtended.xml', 'http://schemas.microsoft.com/office/word/2012/wordml', 'commentsEx', 'commentEx', 'paraId'],
  ['word/commentsIds.xml', 'http://schemas.microsoft.com/office/word/2016/wordml/cid', 'commentsIds', 'commentId', 'paraId'],
  ['word/commentsExtensible.xml', 'http://schemas.microsoft.com/office/word/2018/wordml/cex', 'commentsExtensible', 'commentExtensible', 'durableId'],
] as const;
const COMMENT_TYPES = new Map([COMMENTS, ...COMMENT_COMPANIONS.map(([name]) => name)].map(name => {
  const local = name.slice(5, -4);
  return [name, [`application/vnd.openxmlformats-officedocument.wordprocessingml.${local}+xml`, `application/vnd.ms-word.${local}+xml`]];
}));

interface XmlElement {
  name: string;
  attributes: Map<string, string>;
  namespaces: Map<string, string>;
  declaredPrefixes: Set<string>;
  prefixes: Set<string>;
  children: XmlElement[];
  content: Array<XmlElement | string | { markup: string }>;
  start: number;
  end: number;
  contentStart: number;
  contentEnd: number;
}

interface XmlPart {
  xml: string;
  root: XmlElement;
  elements: XmlElement[];
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  mode: string | undefined;
}

export interface CommentSource {
  zip: JSZip;
  relationships: Relationship[];
  comments: Array<{ xml: string; element: XmlElement }>;
}

interface CommentOutput {
  id: string;
  parentId: string | null;
  done?: boolean;
}

function requireOutput(value: unknown, message: string): asserts value {
  if (!value) throw new DocumentToolError('UNSUPPORTED_TARGET', message);
}

function expanded(namespace: string, local: string): string {
  return `{${namespace}}${local}`;
}

function xmlTokens(value: string | undefined): string[] {
  return value?.match(/\S+/g) ?? [];
}

function compatibilityPrefixes(attributes: Map<string, string>, includeIgnorable = true): string[] {
  return [
    ...xmlTokens(attributes.get(expanded(MC, 'MustUnderstand'))),
    ...xmlTokens(attributes.get(expanded(MC, 'ProcessContent'))).map(token => token.includes(':') ? token.split(':')[0] : ''),
    ...(includeIgnorable ? xmlTokens(attributes.get(expanded(MC, 'Ignorable'))) : []),
  ];
}

function xmlValue(value: string): string {
  return value.replace(/&(?:#x([\da-fA-F]+)|#(\d+)|(amp|lt|gt|apos|quot));|(&)/g, (entity, hex, decimal, named, invalid) => {
    requireOutput(!invalid, 'The comment package contains invalid XML entities.');
    if (named) return ({ amp: '&', lt: '<', gt: '>', apos: "'", quot: '"' } as Record<string, string>)[named] ?? entity;
    const code = Number.parseInt(hex ?? decimal, hex ? 16 : 10);
    requireOutput(code === 9 || code === 10 || code === 13 || code >= 32 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) && code !== 0xfffe && code !== 0xffff,
      'The comment package contains invalid XML characters.');
    return String.fromCodePoint(code);
  });
}

function parseXml(xml: string): XmlPart {
  const elements: XmlElement[] = [];
  const stack: Array<{ element: XmlElement; qname: string }> = [];
  const tokens = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(?:[^"'<>]|"[^"<]*"|'[^'<]*')*>|[^<]+/gy;
  const qname = '[\\p{L}_][\\p{L}\\p{N}_.\\-\\p{M}]*(?::[\\p{L}_][\\p{L}\\p{N}_.\\-\\p{M}]*)?';
  const opening = new RegExp(`^<(${qname})`, 'u');
  const closing = new RegExp(`^</(${qname})\\s*>$`, 'u');
  const attribute = new RegExp(`\\s+(${qname})\\s*=\\s*(?:"([^"<]*)"|'([^'<]*)')`, 'uy');
  const appendText = (value: string) => {
    const content = stack.at(-1)?.element.content;
    if (!content || !value) return;
    const previous = content.at(-1);
    if (typeof previous === 'string') content[content.length - 1] = previous + value;
    else content.push(value);
  };
  let root: XmlElement | undefined;
  let offset = 0;
  while (offset < xml.length) {
    tokens.lastIndex = offset;
    const match = tokens.exec(xml);
    requireOutput(match, 'The comment package contains malformed XML.');
    const token = match[0];
    const start = offset;
    offset = tokens.lastIndex;
    if (token.startsWith('<!--')) {
      requireOutput(!token.slice(4, -3).includes('--') && !token.slice(4, -3).endsWith('-'), 'The comment package contains malformed XML comments.');
      stack.at(-1)?.element.content.push({ markup: token });
      continue;
    }
    if (token.startsWith('<?')) {
      requireOutput(/^<\?[\p{L}_][\p{L}\p{N}_.:\-\p{M}]*(?:\s[\s\S]*)?\?>$/u.test(token), 'The comment package contains an invalid processing instruction.');
      stack.at(-1)?.element.content.push({ markup: token });
      continue;
    }
    if (token.startsWith('<![CDATA[')) {
      requireOutput(stack.length, 'The comment package contains misplaced CDATA.');
      appendText(token.slice(9, -3));
      continue;
    }
    if (!token.startsWith('<')) {
      requireOutput(stack.length || !token.replace(/^\uFEFF/, '').trim(), 'The comment package contains text outside its root.');
      requireOutput(!token.includes(']]>'), 'The comment package contains malformed XML text.');
      appendText(xmlValue(token));
      continue;
    }
    const end = token.match(closing);
    if (end) {
      const entry = stack.pop();
      requireOutput(entry && entry.qname === end[1], 'The comment package contains mismatched XML tags.');
      entry.element.end = offset;
      entry.element.contentEnd = start;
      continue;
    }
    const begin = token.match(opening);
    requireOutput(begin, 'The comment package contains unsupported XML markup.');
    const raw = new Map<string, string>();
    let position = begin[0].length;
    while (position < token.length) {
      attribute.lastIndex = position;
      const attr = attribute.exec(token);
      if (!attr) break;
      requireOutput(!raw.has(attr[1]), 'The comment package contains duplicate XML attributes.');
      raw.set(attr[1], xmlValue((attr[2] ?? attr[3]).replace(/[\t\r\n]/g, ' ')));
      position = attribute.lastIndex;
    }
    requireOutput(/^\s*\/?>$/.test(token.slice(position)), 'The comment package contains malformed XML attributes.');
    const parent = stack.at(-1)?.element;
    const namespaces = new Map<string, string>(parent?.namespaces ?? [['xml', 'http://www.w3.org/XML/1998/namespace']]);
    const declaredPrefixes = new Set<string>();
    for (const [name, value] of raw) {
      const prefix = name === 'xmlns' ? '' : name.startsWith('xmlns:') ? name.slice(6) : undefined;
      if (prefix !== undefined) { namespaces.set(prefix, value); declaredPrefixes.add(prefix); }
    }
    const resolve = (name: string, isAttribute = false): string => {
      const [prefix, local] = name.split(':');
      if (local === undefined) return expanded(isAttribute ? '' : namespaces.get('') ?? '', prefix);
      const namespace = namespaces.get(prefix);
      requireOutput(namespace, 'The comment package contains an unbound XML prefix.');
      return expanded(namespace, local);
    };
    const attributes = new Map<string, string>();
    const prefixes = new Set<string>([begin[1].includes(':') ? begin[1].split(':')[0] : '']);
    for (const [name, value] of raw) {
      if (name === 'xmlns' || name.startsWith('xmlns:')) continue;
      if (name.includes(':')) prefixes.add(name.split(':')[0]);
      const key = resolve(name, true);
      requireOutput(!attributes.has(key), 'The comment package contains duplicate expanded attributes.');
      attributes.set(key, value);
    }
    const element: XmlElement = { name: resolve(begin[1]), attributes, namespaces, declaredPrefixes, prefixes, children: [], content: [], start, end: offset, contentStart: offset, contentEnd: offset };
    if (parent) { parent.children.push(element); parent.content.push(element); }
    else { requireOutput(!root, 'The comment package contains multiple XML roots.'); root = element; }
    elements.push(element);
    if (!token.endsWith('/>')) stack.push({ element, qname: begin[1] });
  }
  requireOutput(root && !stack.length, 'The comment package contains incomplete XML.');
  return { xml, root, elements };
}

async function readXml(zip: JSZip, name: string, namespace: string, local: string): Promise<XmlPart> {
  const xml = await zip.file(name)?.async('string');
  requireOutput(xml !== undefined, `The comment package is missing ${name}.`);
  const part = parseXml(xml);
  requireOutput(part.root.name === expanded(namespace, local), `The comment package has an invalid root in ${name}.`);
  return part;
}

async function relationships(zip: JSZip, name: string): Promise<Relationship[]> {
  const { root } = await readXml(zip, name, RELS, 'Relationships');
  return root.children.map(element => {
    requireOutput(element.name === expanded(RELS, 'Relationship') && !element.children.length, 'The comment package contains an invalid relationship element.');
    const id = element.attributes.get('{}Id');
    const type = element.attributes.get('{}Type');
    const target = element.attributes.get('{}Target');
    requireOutput(id && type && target, 'The comment package contains an incomplete relationship.');
    return { id, type, target, mode: element.attributes.get('{}TargetMode') };
  });
}

function partName(base: string, target: string): string {
  const parts: string[] = [];
  for (const segment of (target.startsWith('/') ? target.slice(1) : `${base}${target}`).split('/')) {
    if (segment === '..') parts.pop();
    else if (segment && segment !== '.') parts.push(segment);
  }
  return parts.join('/');
}

export async function readCommentPart(source: ArrayBuffer, allowLegacy = false): Promise<CommentSource> {
  const zip = await JSZip.loadAsync(source);
  const main = (await relationships(zip, '_rels/.rels')).filter(rel => rel.type.endsWith('/officeDocument') && rel.mode !== 'External');
  requireOutput(main.length === 1 && partName('', main[0].target) === 'word/document.xml',
    'The document keeps its main text in a non-standard part, so a new comment cannot be added safely.');
  const rels = await relationships(zip, DOCUMENT_RELS);
  const comments = rels.filter(rel => rel.type.endsWith('/comments') && rel.mode !== 'External');
  requireOutput(comments.every(rel => partName('word/', rel.target) === COMMENTS),
    'The document keeps its comments in a non-standard part, so a new comment cannot be added without rewriting existing comments.');
  const sourceComments: CommentSource['comments'] = [];
  if (zip.file(COMMENTS)) {
    const part = await readXml(zip, COMMENTS, W, 'comments');
    requireOutput(allowLegacy || part.root.namespaces.get('w14') === W14,
      'The existing comments part predates paragraph ids, so a new comment cannot be added without rewriting existing comments.');
    for (const element of part.root.children) {
      if (element.name === expanded(W, 'comment')) sourceComments.push({ xml: part.xml.slice(element.start, element.end), element });
    }
  }
  return { zip, relationships: rels, comments: sourceComments };
}

function withoutWhitespaceText(xml: string): string {
  return xml.replace(/<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|[^<]+/g,
    token => /^[\t\r\n ]+$/.test(token) ? '' : token);
}

function companionContent(part: XmlPart, entries: XmlElement[]): string[] {
  const known = new Set(entries);
  let offset = part.root.contentStart;
  let content = '';
  for (const child of part.root.children) {
    content += withoutWhitespaceText(part.xml.slice(offset, child.start));
    if (!known.has(child)) content += part.xml.slice(child.start, child.end);
    offset = child.end;
  }
  const prolog = part.xml.slice(0, part.root.start).replace(/^\uFEFF/, '').replace(/^<\?xml\s[\s\S]*?\?>/, '');
  return [
    withoutWhitespaceText(prolog),
    content + withoutWhitespaceText(part.xml.slice(offset, part.root.contentEnd)),
    withoutWhitespaceText(part.xml.slice(part.root.end)),
  ];
}

function companionPrefixes(part: XmlPart): { used: Set<string>; inherited: Set<string> } {
  const inherited = new Set(compatibilityPrefixes(part.root.attributes, false));
  const used = new Set([...part.root.prefixes, ...inherited]);
  const visit = (element: XmlElement, local: Set<string>) => {
    const scoped = new Set([...local, ...element.declaredPrefixes]);
    for (const prefix of [...element.prefixes, ...compatibilityPrefixes(element.attributes)]) {
      used.add(prefix);
      if (!scoped.has(prefix)) inherited.add(prefix);
    }
    for (const child of element.children) visit(child, scoped);
  };
  for (const child of part.root.children) visit(child, new Set());
  return { used, inherited };
}

function sameXmlElement(left: XmlElement, right: XmlElement): boolean {
  if (left.name !== right.name || left.attributes.size !== right.attributes.size
    || ![...left.attributes].every(([key, value]) => right.attributes.get(key) === value)) return false;
  const significant = (element: XmlElement) => element.content.filter(node => typeof node !== 'string' || !/^[\t\r\n ]*$/.test(node));
  const content = significant(left);
  const other = significant(right);
  return content.length === other.length && content.every((node, index) => {
    const kept = other[index];
    if (typeof node === 'string' || typeof kept === 'string') return node === kept;
    if ('markup' in node || 'markup' in kept) return 'markup' in node && 'markup' in kept && node.markup === kept.markup;
    return sameXmlElement(node, kept);
  });
}

function withoutElements(part: XmlPart, elements: readonly XmlElement[]): string {
  let xml = part.xml;
  for (const element of [...elements].sort((left, right) => right.start - left.start)) xml = xml.slice(0, element.start) + xml.slice(element.end);
  return xml;
}

function commentIds(seeded: ReadonlySet<string>, comments: readonly CommentOutput[], identities: ReadonlyMap<string, DocxCommentSaveIdentity> = new Map()): Map<string, string> {
  const ids = new Map([...seeded].map(id => [id, id]));
  const used = new Set([...seeded].map(Number).concat([...identities.values()].map(identity => identity.id)));
  let next = Math.max(0, ...used) + 1;
  for (const comment of comments) {
    if (ids.has(comment.id)) continue;
    const identity = identities.get(comment.id);
    if (identity) { ids.set(comment.id, String(identity.id)); continue; }
    let id = /^\d+$/u.test(comment.id) ? Number(comment.id) : Number.NaN;
    if (!Number.isSafeInteger(id) || used.has(id)) {
      while (used.has(next)) next++;
      id = next;
    }
    used.add(id);
    ids.set(comment.id, String(id));
  }
  return ids;
}

export async function readSavedCommentIdentities(source: CommentSource, bytes: Uint8Array, comments: readonly CommentOutput[], previous: ReadonlyMap<string, DocxCommentSaveIdentity>): Promise<Map<string, DocxCommentSaveIdentity>> {
  const identities = new Map(previous);
  if (!comments.length) return identities;
  const zip = await JSZip.loadAsync(bytes);
  const part = await readXml(zip, COMMENTS, W, 'comments');
  const seeded = new Set(source.comments.map(comment => comment.element.attributes.get(expanded(W, 'id'))!));
  const ids = commentIds(seeded, comments, previous);
  const durableIds = new Map<string, string>();
  if (zip.file('word/commentsIds.xml')) {
    const [, namespace, root, entry] = COMMENT_COMPANIONS[1];
    const metadata = await readXml(zip, 'word/commentsIds.xml', namespace, root);
    for (const element of metadata.root.children.filter(element => element.name === expanded(namespace, entry))) {
      const paraId = element.attributes.get(expanded(namespace, 'paraId'));
      const durableId = element.attributes.get(expanded(namespace, 'durableId'));
      if (paraId && durableId) durableIds.set(paraId.toUpperCase(), durableId);
    }
  }
  for (const comment of comments) {
    const id = ids.get(comment.id)!;
    const element = part.root.children.find(element => element.name === expanded(W, 'comment') && element.attributes.get(expanded(W, 'id')) === id);
    requireOutput(element, 'Saving comments would lose a comment identity.');
    const paraId = lastParagraph(element)?.attributes.get(expanded(W14, 'paraId'));
    const durableId = paraId ? durableIds.get(paraId.toUpperCase()) : undefined;
    identities.set(comment.id, { id: Number(id), ...(paraId ? { paraId } : {}), ...(durableId ? { durableId } : {}) });
  }
  return identities;
}

function storyProjection(part: XmlPart, deleted: ReadonlySet<string>): string {
  const parents = new Map<XmlElement, XmlElement>();
  for (const element of part.elements) for (const child of element.children) parents.set(child, element);
  const isMarker = (element: XmlElement) => ['commentRangeStart', 'commentRangeEnd', 'commentReference'].some(name => element.name === expanded(W, name));
  const idOf = (element: XmlElement) => element.attributes.get(expanded(W, 'id'));
  const drop = (element: XmlElement) => { const id = idOf(element); return id !== undefined && deleted.has(id); };
  const edits = new Map<XmlElement, string>();
  for (const element of part.elements.filter(isMarker)) {
    const id = idOf(element);
    requireOutput(id !== undefined, 'Saving comments would leave a marker without a comment id.');
    if (drop(element)) {
      let target = element;
      const parent = parents.get(element);
      if (element.name === expanded(W, 'commentReference') && parent?.name === expanded(W, 'r')
        && parent.content.every(node => typeof node === 'string' ? !node.trim()
          : !('markup' in node) && (node.name === expanded(W, 'rPr') || node.name === expanded(W, 'commentReference') && drop(node)))) target = parent;
      edits.set(target, '');
    }
  }
  const outer = [...edits].filter(([element]) => {
    let parent = parents.get(element);
    while (parent) { if (edits.has(parent)) return false; parent = parents.get(parent); }
    return true;
  });
  let xml = part.xml;
  for (const [element, replacement] of outer.sort(([left], [right]) => right.start - left.start)) xml = xml.slice(0, element.start) + replacement + xml.slice(element.end);
  return xml;
}

async function verifyCommentRemoval(source: JSZip, zip: JSZip, seeded: ReadonlySet<string>, previous: readonly CommentOutput[], deleted: ReadonlySet<string>, identities: ReadonlyMap<string, DocxCommentSaveIdentity>): Promise<void> {
  const beforeIds = commentIds(seeded, previous, identities);
  const removedIds = new Set([...deleted].map(id => beforeIds.get(id) ?? String(identities.get(id)?.id ?? id)));
  const survivingIds = new Set<string>();
  const removedParaIds = new Set<string>();
  const removedDurableIds = new Set<string>();
  const verifyPart = async (name: string, namespace: string, root: string, entry: string, drop: (element: XmlElement) => boolean) => {
    if (!source.file(name)) {
      requireOutput(!zip.file(name), `Saving comments would add an unexpected comment part ${name}.`);
      return;
    }
    const before = await readXml(source, name, namespace, root);
    const entries = before.root.children.filter(element => element.name === expanded(namespace, entry));
    const removed = entries.filter(drop);
    requireOutput(zip.file(name), `Saving comments would remove the comment part ${name}.`);
    const saved = await readXml(zip, name, namespace, root);
    requireOutput(withoutElements(before, removed) === saved.xml, `Saving comments would rewrite a surviving comment or companion entry in ${name}.`);
  };
  if (source.file(COMMENTS)) {
    const before = await readXml(source, COMMENTS, W, 'comments');
    for (const element of before.root.children.filter(element => element.name === expanded(W, 'comment'))) {
      const id = element.attributes.get(expanded(W, 'id'));
      requireOutput(id !== undefined, 'Saving comments would lose a comment id.');
      if (!removedIds.has(id)) { survivingIds.add(id); continue; }
      const paraId = lastParagraph(element)?.attributes.get(expanded(W14, 'paraId'));
      if (paraId) removedParaIds.add(paraId.toUpperCase());
      const durableId = element.attributes.get(expanded(COMMENT_COMPANIONS[1][1], 'durableId'));
      if (durableId) removedDurableIds.add(durableId.toUpperCase());
    }
  }
  if (source.file('word/commentsIds.xml')) {
    const [name, namespace, root, entry] = COMMENT_COMPANIONS[1];
    const before = await readXml(source, name, namespace, root);
    for (const element of before.root.children.filter(element => element.name === expanded(namespace, entry))) {
      const paraId = element.attributes.get(expanded(namespace, 'paraId'));
      const durableId = element.attributes.get(expanded(namespace, 'durableId'));
      if (paraId && durableId && removedParaIds.has(paraId.toUpperCase())) removedDurableIds.add(durableId.toUpperCase());
    }
  }
  await verifyPart(COMMENTS, W, 'comments', 'comment', element => {
    const id = element.attributes.get(expanded(W, 'id'));
    return id !== undefined && removedIds.has(id);
  });
  for (const [name, namespace, root, entry, key] of COMMENT_COMPANIONS) {
    await verifyPart(name, namespace, root, entry, element => {
      const id = element.attributes.get(expanded(namespace, key));
      return id !== undefined && (key === 'paraId' ? removedParaIds : removedDurableIds).has(id.toUpperCase());
    });
  }
  const storyRoots = new Set(['document', 'hdr', 'ftr', 'footnotes', 'endnotes'].map(name => expanded(W, name)));
  const mutable = new Set(COMMENT_TYPES.keys());
  for (const [name, file] of Object.entries(source.files)) {
    if (file.dir || mutable.has(name)) continue;
    const kept = zip.file(name);
    requireOutput(kept, `Saving comments would remove the unrelated package part ${name}.`);
    if (name.endsWith('.xml') || name === DOCUMENT_RELS) {
      const xml = await file.async('string');
      const savedXml = await kept.async('string');
      if (xml.includes(W)) {
        const before = parseXml(xml);
        if (storyRoots.has(before.root.name)) {
          const saved = parseXml(savedXml);
          requireOutput(saved.elements.every(element => {
            const id = element.attributes.get(expanded(W, 'id'));
            return !['commentRangeStart', 'commentRangeEnd', 'commentReference'].some(local => element.name === expanded(W, local))
              || id !== undefined && survivingIds.has(id);
          }), `Saving comments would leave an orphan comment marker in ${name}.`);
          requireOutput(storyProjection(before, removedIds) === saved.xml,
            `Saving comments would change unrelated story markup in ${name}.`);
          continue;
        }
      }
    }
    const bytes = await file.async('uint8array');
    const saved = await kept.async('uint8array');
    requireOutput(bytes.length === saved.length && bytes.every((byte, index) => byte === saved[index]),
      `Saving comments would change the unrelated package part ${name}.`);
  }
  for (const [name, file] of Object.entries(zip.files)) {
    requireOutput(file.dir || source.file(name), `Saving comments would add the unrelated package part ${name}.`);
  }
}

async function verifyCommentCompanions(source: JSZip, zip: JSZip, comments: XmlElement[], fresh: XmlElement[], resolved: ReadonlySet<string>, deleted: readonly XmlElement[] = []): Promise<void> {
  const paraId = (comment: XmlElement) => comment.children.filter(child => child.name === expanded(W, 'p')).at(-1)?.attributes.get(expanded(W14, 'paraId'));
  const ids = comments.map(paraId).filter((id): id is string => id !== undefined);
  const paraIds = new Set(ids.map(id => id.toUpperCase()));
  const removedParaIds = new Set(deleted.map(paraId).filter((id): id is string => id !== undefined).map(id => id.toUpperCase()));
  const removedDurableIds = new Set(deleted.map(comment => comment.attributes.get(expanded('http://schemas.microsoft.com/office/word/2016/wordml/cid', 'durableId'))).filter((id): id is string => id !== undefined).map(id => id.toUpperCase()));
  if (source.file('word/commentsIds.xml')) {
    const [, namespace, root, entry] = COMMENT_COMPANIONS[1];
    const original = await readXml(source, 'word/commentsIds.xml', namespace, root);
    for (const element of original.root.children.filter(element => element.name === expanded(namespace, entry))) {
      const paraId = element.attributes.get(expanded(namespace, 'paraId'))?.toUpperCase();
      const durableId = element.attributes.get(expanded(namespace, 'durableId'))?.toUpperCase();
      if (paraId && durableId && removedParaIds.has(paraId)) removedDurableIds.add(durableId);
    }
  }
  requireOutput(fresh.every(comment => paraId(comment)) && paraIds.size === ids.length, 'Saving comments would create invalid or duplicate comment paragraph ids.');
  const durableIds = new Set<string>();
  const hasCommentIds = !!zip.file('word/commentsIds.xml');
  for (const [name, namespace, root, entry, key] of COMMENT_COMPANIONS) {
    if (!source.file(name) && !zip.file(name)) continue;
    const saved = await readXml(zip, name, namespace, root);
    const elements = saved.root.children.filter(element => element.name === expanded(namespace, entry));
    const byId = new Map<string, XmlElement>();
    for (const element of elements) {
      const id = element.attributes.get(expanded(namespace, key))?.toUpperCase();
      requireOutput(id && !(key === 'paraId' ? removedParaIds : removedDurableIds).has(id), `Saving comments would leave an orphan entry in ${name}.`);
      requireOutput(id && !byId.has(id), `Saving comments would create invalid or duplicate entries in ${name}.`);
      byId.set(id, element);
      if (entry === 'commentId') {
        const durableId = element.attributes.get(expanded(namespace, 'durableId'))?.toUpperCase();
        requireOutput(durableId && !durableIds.has(durableId), `Saving comments would create invalid or duplicate durable ids in ${name}.`);
        durableIds.add(durableId);
      }
    }
    if (source.file(name)) {
      const original = await readXml(source, name, namespace, root);
      const ignorable = expanded(MC, 'Ignorable');
      const { used, inherited } = companionPrefixes(original);
      const requiredIgnorable = xmlTokens(original.root.attributes.get(ignorable)).filter(prefix => used.has(prefix));
      const savedIgnorable = new Set(xmlTokens(saved.root.attributes.get(ignorable)));
      for (const prefix of requiredIgnorable) inherited.add(prefix);
      requireOutput([...original.root.attributes].every(([key, value]) => key === ignorable || saved.root.attributes.get(key) === value)
        && [...inherited].every(prefix => {
          const uri = original.root.namespaces.get(prefix) ?? '';
          return (!prefix || !!uri) && uri === (saved.root.namespaces.get(prefix) ?? '');
        })
        && requiredIgnorable.every(prefix => savedIgnorable.has(prefix)),
        `Saving comments would change the root metadata in ${name}.`);
      const entries = original.root.children.filter(element => element.name === expanded(namespace, entry));
      const removedEntries: XmlElement[] = [];
      const changedEntries: XmlElement[] = [];
      const savedChanged: XmlElement[] = [];
      for (const element of entries) {
        const id = element.attributes.get(expanded(namespace, key))?.toUpperCase();
        if (id && (key === 'paraId' ? removedParaIds : removedDurableIds).has(id)) { removedEntries.push(element); continue; }
        let kept = id === undefined ? undefined : byId.get(id);
        if (id && kept && namespace === W15 && resolved.has(id.toUpperCase())) {
          const attributes = new Map(kept.attributes);
          const done = expanded(W15, 'done');
          const value = element.attributes.get(done);
          if (value === undefined) attributes.delete(done);
          else attributes.set(done, value);
          changedEntries.push(element);
          savedChanged.push(kept);
          kept = { ...kept, attributes };
        }
        requireOutput(id && kept && sameXmlElement(element, kept),
          `Saving comments would rewrite an existing entry in ${name}.`);
        byId.delete(id);
      }
      if (deleted.length) {
        const freshEntries = elements.filter(element => byId.has(element.attributes.get(expanded(namespace, key))?.toUpperCase() ?? ''));
        requireOutput(withoutElements(original, [...removedEntries, ...changedEntries]) === withoutElements(saved, [...freshEntries, ...savedChanged]),
          `Saving comments would rewrite unrelated markup in ${name}.`);
      }
      const content = companionContent(saved, elements);
      requireOutput(companionContent(original, entries).every((value, index) => value === content[index]),
        `Saving comments would change non-entry markup in ${name}.`);
    }
    for (const id of byId.keys()) {
      requireOutput(key === 'paraId' ? paraIds.has(id) : !hasCommentIds || durableIds.has(id),
        `Saving comments would add an entry for an unknown comment in ${name}.`);
    }
  }
}

function lastParagraph(comment: XmlElement): XmlElement | undefined {
  return comment.children.filter(element => element.name === expanded(W, 'p')).at(-1);
}

const XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space';

function sameExceptCommentText(left: XmlElement, right: XmlElement, paragraphIdTarget?: XmlElement): boolean {
  const attributes = new Map(left.attributes);
  if (left.name === expanded(W, 't') && right.name === left.name) {
    const sourceText = right.content.join('');
    if (attributes.get(XML_SPACE) === 'preserve' && sourceText === sourceText.trim()) {
      if (!right.attributes.has(XML_SPACE)) attributes.delete(XML_SPACE);
      else if (right.attributes.get(XML_SPACE) === 'default') attributes.set(XML_SPACE, 'default');
    }
    if (attributes.size !== right.attributes.size
      || ![...attributes].every(([key, value]) => right.attributes.get(key) === value)) return false;
    return ![...left.content, ...right.content].some(node => typeof node !== 'string');
  }
  if (left === paragraphIdTarget) attributes.delete(expanded(W14, 'paraId'));
  if (left.name !== right.name || attributes.size !== right.attributes.size
    || ![...attributes].every(([key, value]) => right.attributes.get(key) === value)) return false;
  const significant = (element: XmlElement) => element.content.filter(node => typeof node !== 'string' || !/^[\t\r\n ]*$/.test(node));
  const content = significant(left);
  const other = significant(right);
  return content.length === other.length && content.every((node, index) => {
    const kept = other[index];
    if (typeof node === 'string' || typeof kept === 'string') return node === kept;
    if ('markup' in node || 'markup' in kept) return 'markup' in node && 'markup' in kept && node.markup === kept.markup;
    return sameExceptCommentText(node, kept, paragraphIdTarget);
  });
}

function preservesComment(saved: XmlPart, element: XmlElement, original: CommentSource['comments'][number], allowParagraphId: boolean, allowText = false): boolean {
  const fragment = saved.xml.slice(element.start, element.end);
  if (fragment === original.xml) return true;
  if (allowText && sameExceptCommentText(element, original.element)) return true;
  if (!allowParagraphId) return false;
  const before = lastParagraph(original.element);
  const after = lastParagraph(element);
  const paraId = after?.attributes.get(expanded(W14, 'paraId'));
  if (!before || before.attributes.has(expanded(W14, 'paraId')) || !paraId) return false;
  if (allowText && sameExceptCommentText(element, original.element, after)) return true;
  const end = before.contentStart - original.element.start;
  const at = end - (original.xml[end - 2] === '/' ? 2 : 1);
  return fragment === `${original.xml.slice(0, at)} w14:paraId="${paraId}"${original.xml.slice(at)}`;
}

export async function verifyCommentOutput(source: CommentSource, bytes: Uint8Array, added: number, expected: readonly CommentOutput[] = [], textEdited: ReadonlySet<string> = new Set(), baseline?: Uint8Array, deleted: ReadonlySet<string> = new Set(), previous?: readonly CommentOutput[], identities: ReadonlyMap<string, DocxCommentSaveIdentity> = new Map()): Promise<void> {
  const zip = await JSZip.loadAsync(bytes);
  const rels = await relationships(zip, DOCUMENT_RELS);
  requireOutput(new Set(rels.map(rel => rel.id)).size === rels.length, 'Saving comments would create duplicate relationship ids.');
  const comments = rels.filter(rel => rel.type.endsWith('/comments'));
  if (!zip.file(COMMENTS)) {
    requireOutput(previous && !source.zip.file(COMMENTS) && added === 0 && expected.length === 0 && comments.length === 0
      && COMMENT_COMPANIONS.every(([name]) => !zip.file(name)), 'Saving comments would lose the comments part.');
    await verifyCommentRemoval(baseline ? await JSZip.loadAsync(baseline) : source.zip, zip, new Set(), previous, deleted, identities);
    return;
  }
  requireOutput(comments.length === 1 && comments[0].type === `${OFFICE_RELS}comments` && comments[0].mode !== 'External' && partName('word/', comments[0].target) === COMMENTS,
    'Saving comments would leave the comments part without a unique internal relationship.');
  for (const original of source.relationships) {
    requireOutput(rels.some(rel => rel.id === original.id && rel.type === original.type && rel.target === original.target && rel.mode === original.mode),
      'Saving comments would change an existing document relationship.');
  }
  const types = await readXml(zip, '[Content_Types].xml', TYPES, 'Types');
  for (const [name, contentType] of COMMENT_TYPES) {
    if (name !== COMMENTS && !zip.file(name)) continue;
    const overrides = types.root.children.filter(element => element.name === expanded(TYPES, 'Override') && element.attributes.get('{}PartName') === `/${name}`);
    requireOutput(overrides.length === 1 && contentType.includes(overrides[0].attributes.get('{}ContentType') ?? ''),
      `Saving comments would leave ${name} without its content type registration.`);
  }
  const saved = await readXml(zip, COMMENTS, W, 'comments');
  const elements = saved.root.children.filter(element => element.name === expanded(W, 'comment'));
  const seeded = new Set(source.comments.map(comment => comment.element.attributes.get(expanded(W, 'id'))!));
  const savedIds = commentIds(seeded, expected, identities);
  const removed = source.comments.filter(comment => deleted.has(comment.element.attributes.get(expanded(W, 'id'))!));
  requireOutput(elements.every(element => !deleted.has(element.attributes.get(expanded(W, 'id'))!)), 'Saving comments would retain a deleted comment.');
  const replies = expected.filter(comment => comment.parentId !== null && !seeded.has(comment.id));
  const resolved = expected.filter(comment => comment.done !== undefined);
  const paragraphIds = new Set([
    ...replies.map(comment => savedIds.get(comment.parentId!)),
    ...resolved.map(comment => savedIds.get(comment.id)),
  ]);
  const fresh = [...elements];
  const changedOriginal: XmlElement[] = [];
  const changedSaved: XmlElement[] = [];
  for (const original of source.comments) {
    const id = original.element.attributes.get(expanded(W, 'id'))!;
    if (deleted.has(id)) continue;
    const index = fresh.findIndex(element => preservesComment(saved, element, original, paragraphIds.has(id), textEdited.has(id)));
    requireOutput(index !== -1, 'Saving comments would rewrite an existing comment.');
    if (paragraphIds.has(id) || textEdited.has(id)) { changedOriginal.push(original.element); changedSaved.push(fresh[index]); }
    fresh.splice(index, 1);
  }
  if (removed.length) {
    const original = await readXml(source.zip, COMMENTS, W, 'comments');
    requireOutput(withoutElements(original, [...removed.map(comment => comment.element), ...changedOriginal]) === withoutElements(saved, [...fresh, ...changedSaved]),
      'Saving comments would rewrite unrelated comments markup.');
  }
  const ids = elements.map(element => element.attributes.get(expanded(W, 'id')));
  requireOutput(ids.every(id => id !== undefined) && new Set(ids).size === ids.length,
    'Saving comments would create invalid or duplicate comment ids.');
  const document = await readXml(zip, 'word/document.xml', W, 'document');
  const body = document.root.children.find(element => element.name === expanded(W, 'body'));
  const threadChanged = replies.length > 0 || resolved.length > 0;
  const references = new Set(document.elements.filter(element => element.name === expanded(W, 'commentReference') && (!threadChanged || body && element.start > body.start && element.end <= body.end)).map(element => element.attributes.get(expanded(W, 'id'))));
  requireOutput(fresh.length === added && fresh.every(element => references.has(element.attributes.get(expanded(W, 'id')))),
    'Saving comments would lose a new comment or its document reference.');
  const resolvedParagraphIds = new Set<string>();
  if (replies.length || resolved.length) {
    const extended = await readXml(zip, 'word/commentsExtended.xml', W15, 'commentsEx');
    const byId = new Map(elements.map(element => [element.attributes.get(expanded(W, 'id')), element]));
    const paragraphId = (id: string) => {
      const comment = byId.get(savedIds.get(id));
      return comment && lastParagraph(comment)?.attributes.get(expanded(W14, 'paraId'))?.toUpperCase();
    };
    const metadata = (id: string) => {
      const paraId = paragraphId(id);
      const rows = extended.root.children.filter(element => element.name === expanded(W15, 'commentEx') && element.attributes.get(expanded(W15, 'paraId'))?.toUpperCase() === paraId);
      requireOutput(paraId && rows.length === 1, 'Saving comments would lose comment thread metadata.');
      return rows[0];
    };
    for (const reply of replies) {
      requireOutput(references.has(savedIds.get(reply.id)), 'Saving comments would lose a new comment or its document reference.');
      const parentId = paragraphId(reply.parentId!);
      requireOutput(parentId && metadata(reply.id).attributes.get(expanded(W15, 'paraIdParent'))?.toUpperCase() === parentId,
        'Saving comments would lose a reply or its parent thread.');
    }
    for (const comment of resolved) {
      requireOutput(metadata(comment.id).attributes.get(expanded(W15, 'done')) === (comment.done ? '1' : '0'),
        'Saving comments would lose a comment resolution state.');
      if (comment.parentId === null) resolvedParagraphIds.add(paragraphId(comment.id)!);
    }
  }
  await verifyCommentCompanions(source.zip, zip, elements, fresh, resolvedParagraphIds, removed.map(comment => comment.element));
  if (previous) {
    await verifyCommentRemoval(baseline ? await JSZip.loadAsync(baseline) : source.zip, zip, seeded, previous, deleted, identities);
    return;
  }
  const mutable = new Set(['word/document.xml', '[Content_Types].xml', DOCUMENT_RELS, 'docProps/core.xml', ...COMMENT_TYPES.keys()]);
  const before = baseline ? await JSZip.loadAsync(baseline) : source.zip;
  for (const [name, file] of Object.entries(before.files)) {
    if (file.dir || mutable.has(name)) continue;
    const original = await file.async('uint8array');
    const saved = await zip.file(name)?.async('uint8array');
    requireOutput(saved && original.length === saved.length && original.every((byte, index) => byte === saved[index]),
      `Saving comments would change the unrelated package part ${name}.`);
  }
}
