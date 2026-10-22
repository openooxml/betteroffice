import JSZip from 'jszip';
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

export async function readCommentPart(source: ArrayBuffer): Promise<CommentSource> {
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
    requireOutput(part.root.namespaces.get('w14') === W14,
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

async function verifyCommentCompanions(source: JSZip, zip: JSZip, comments: XmlElement[], fresh: XmlElement[], resolved: ReadonlySet<string>): Promise<void> {
  const paraId = (comment: XmlElement) => comment.children.filter(child => child.name === expanded(W, 'p')).at(-1)?.attributes.get(expanded(W14, 'paraId'));
  const ids = comments.map(paraId).filter((id): id is string => id !== undefined);
  const paraIds = new Set(ids);
  requireOutput(fresh.every(comment => paraId(comment)) && paraIds.size === ids.length, 'Saving comments would create invalid or duplicate comment paragraph ids.');
  const durableIds = new Set<string>();
  const hasCommentIds = !!zip.file('word/commentsIds.xml');
  for (const [name, namespace, root, entry, key] of COMMENT_COMPANIONS) {
    if (!source.file(name) && !zip.file(name)) continue;
    const saved = await readXml(zip, name, namespace, root);
    const elements = saved.root.children.filter(element => element.name === expanded(namespace, entry));
    const byId = new Map<string, XmlElement>();
    for (const element of elements) {
      const id = element.attributes.get(expanded(namespace, key));
      requireOutput(id && !byId.has(id), `Saving comments would create invalid or duplicate entries in ${name}.`);
      byId.set(id, element);
      if (entry === 'commentId') {
        const durableId = element.attributes.get(expanded(namespace, 'durableId'));
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
      for (const element of entries) {
        const id = element.attributes.get(expanded(namespace, key));
        let kept = id === undefined ? undefined : byId.get(id);
        if (id && kept && namespace === W15 && resolved.has(id.toUpperCase())) {
          const attributes = new Map(kept.attributes);
          const done = expanded(W15, 'done');
          const value = element.attributes.get(done);
          if (value === undefined) attributes.delete(done);
          else attributes.set(done, value);
          kept = { ...kept, attributes };
        }
        requireOutput(id && kept && sameXmlElement(element, kept),
          `Saving comments would rewrite an existing entry in ${name}.`);
        byId.delete(id);
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

function preservesComment(saved: XmlPart, element: XmlElement, original: CommentSource['comments'][number], allowParagraphId: boolean): boolean {
  const fragment = saved.xml.slice(element.start, element.end);
  if (fragment === original.xml) return true;
  if (!allowParagraphId) return false;
  const before = lastParagraph(original.element);
  const after = lastParagraph(element);
  const paraId = after?.attributes.get(expanded(W14, 'paraId'));
  if (!before || before.attributes.has(expanded(W14, 'paraId')) || !paraId) return false;
  const end = before.contentStart - original.element.start;
  const at = end - (original.xml[end - 2] === '/' ? 2 : 1);
  return fragment === `${original.xml.slice(0, at)} w14:paraId="${paraId}"${original.xml.slice(at)}`;
}

export async function verifyCommentOutput(source: CommentSource, bytes: Uint8Array, added: number, expected: readonly CommentOutput[] = []): Promise<void> {
  const zip = await JSZip.loadAsync(bytes);
  const rels = await relationships(zip, DOCUMENT_RELS);
  requireOutput(new Set(rels.map(rel => rel.id)).size === rels.length, 'Saving comments would create duplicate relationship ids.');
  const comments = rels.filter(rel => rel.type.endsWith('/comments'));
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
  const savedIds = new Map([...seeded].map(id => [id, id]));
  const used = new Set([...seeded].map(Number));
  let next = Math.max(0, ...used) + 1;
  for (const comment of expected) {
    if (savedIds.has(comment.id)) continue;
    let id = /^\d+$/u.test(comment.id) ? Number(comment.id) : Number.NaN;
    if (!Number.isSafeInteger(id) || used.has(id)) {
      while (used.has(next)) next++;
      id = next;
    }
    used.add(id);
    savedIds.set(comment.id, String(id));
  }
  const replies = expected.filter(comment => comment.parentId !== null && !seeded.has(comment.id));
  const resolved = expected.filter(comment => comment.done !== undefined);
  const paragraphIds = new Set([
    ...replies.map(comment => savedIds.get(comment.parentId!)),
    ...resolved.map(comment => savedIds.get(comment.id)),
  ]);
  const fresh = [...elements];
  for (const original of source.comments) {
    const index = fresh.findIndex(element => preservesComment(saved, element, original, paragraphIds.has(original.element.attributes.get(expanded(W, 'id')))));
    requireOutput(index !== -1, 'Saving comments would rewrite an existing comment.');
    fresh.splice(index, 1);
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
  await verifyCommentCompanions(source.zip, zip, elements, fresh, resolvedParagraphIds);
  const mutable = new Set(['word/document.xml', '[Content_Types].xml', DOCUMENT_RELS, 'docProps/core.xml', ...COMMENT_TYPES.keys()]);
  for (const [name, file] of Object.entries(source.zip.files)) {
    if (file.dir || mutable.has(name)) continue;
    const original = await file.async('uint8array');
    const saved = await zip.file(name)?.async('uint8array');
    requireOutput(saved && original.length === saved.length && original.every((byte, index) => byte === saved[index]),
      `Saving comments would change the unrelated package part ${name}.`);
  }
}
