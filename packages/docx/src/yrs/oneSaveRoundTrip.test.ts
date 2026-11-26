import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createDocx, parseDocx, repackDocx, repackDocxWithWarnings } from '../docx';
import { repackDocxFromRaw } from '../docx/rezip';
import type { RawDocxContent } from '../docx/unzip';
import { rezipContainer, unzipContainer } from '../docx/wasm';
import type { BlockContent } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import {
  captureSessionSave,
  createYrsSession,
  saveYrsDocx,
  writeSessionSave,
  yrsToDocument,
  type YrsSession,
} from './index';
import {
  DIAGRAM_RELS,
  EDIT_PARAGRAPH_ID,
  EXTENDED_COMMENT_PARTS,
  ONE_SAVE_NS as NS,
  ONE_SAVE_PROBES,
  oneSavePackage,
} from './__fixtures__/oneSavePackage';

const BODY = 'word/document.xml';
const CORE = 'docProps/core.xml';
const TYPES = '[Content_Types].xml';
const EDIT = ' (edited)';
const MOVE_STARTS = { moveFromRangeStart: '10', moveToRangeStart: '20' };
const WORD_ID_REQUIRED = new Set([
  ...Object.keys(MOVE_STARTS), 'moveFromRangeEnd', 'moveToRangeEnd', 'moveFrom', 'moveTo',
  'comment', 'commentRangeStart', 'commentRangeEnd', 'commentReference', 'bookmarkStart', 'bookmarkEnd',
]);
const DOCUMENT_REFERENCES = {
  chart: ['rIdChart', 'charts/chart1.xml'],
  headerReference: ['rIdHeader', 'header1.xml'],
  footerReference: ['rIdFooter', 'footer1.xml'],
} as const;

type Parts = ReturnType<typeof unzipContainer>;
type Saved = { bytes: Uint8Array; warnings: string[] | 'n/a' };
type Save = (bytes: Uint8Array) => Promise<Saved>;

function partChanges(before: Parts, after: Parts) {
  const source = Object.keys(before).sort();
  const byteIdentical = source.filter((name) => Object.hasOwn(after, name) &&
    Buffer.from(before[name]).equals(after[name]));
  const identical = new Set(byteIdentical);
  return {
    byteIdentical,
    changed: source.filter((name) => Object.hasOwn(after, name) && !identical.has(name)),
    added: Object.keys(after).filter((name) => !Object.hasOwn(before, name)).sort(),
    removed: source.filter((name) => !Object.hasOwn(after, name)),
  };
}

interface Element {
  namespace: string;
  local: string;
  attributes: Map<string, string>;
  parent?: Element;
  start: number;
  end: number;
  text: string;
}

const qualified = (namespace: string, local: string) => namespace ? `{${namespace}}${local}` : local;
const attr = (node: Element, local: string, namespace = '') =>
  node.attributes.get(qualified(namespace, local));
const isElement = (node: Element | undefined, namespace: string, local: string) =>
  node?.namespace === namespace && node.local === local;

function xmlValue(value: string): string {
  if (/&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/.test(value)) {
    throw new Error('invalid entity reference');
  }
  return value.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (_, entity: string) => {
    const named: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
    if (Object.hasOwn(named, entity)) return named[entity];
    return String.fromCodePoint(entity.startsWith('#x')
      ? Number.parseInt(entity.slice(2), 16) : Number(entity.slice(1)));
  });
}

function readElements(xml: string, elements: Element[]): void {
  const token = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<\/([A-Za-z_][\w.:-]*)\s*>|<([A-Za-z_][\w.:-]*)((?:\s+[A-Za-z_][\w.:-]*\s*=\s*(?:"[^"<]*"|'[^'<]*'))*)\s*(\/?)>|([^<]+)/gy;
  const attribute = /([A-Za-z_][\w.:-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  const stack: Array<{ name: string; node: Element; scope: Map<string, string> }> = [];
  const xmlScope = new Map([['xml', 'http://www.w3.org/XML/1998/namespace']]);
  let roots = 0;
  let end = 0;
  for (const match of xml.matchAll(token)) {
    const [raw, cdata, close, open, attributes, empty, text] = match;
    const start = match.index;
    end = start + raw.length;
    const parent = stack[stack.length - 1]?.node;
    if (close) {
      const frame = stack.pop();
      if (frame?.name !== close) throw new Error(`unexpected closing tag </${close}>`);
      frame.node.end = end;
    } else if (open) {
      if (!parent && ++roots > 1) throw new Error('multiple document elements');
      const scope = new Map(stack[stack.length - 1]?.scope ?? xmlScope);
      const names = new Set<string>();
      const values = [...attributes.matchAll(attribute)].map(([, name, double, single]) => {
        if (names.has(name)) throw new Error(`duplicate attribute ${name}`);
        names.add(name);
        const value = xmlValue(double ?? single);
        if (name === 'xmlns') scope.set('', value);
        else if (name.startsWith('xmlns:')) scope.set(name.slice(6), value);
        return [name, value] as const;
      });
      const expand = (name: string, element = false) => {
        const [first, local, extra] = name.split(':');
        if (local === undefined) return [element ? scope.get('') ?? '' : '', first] as const;
        const namespace = scope.get(first);
        if (!namespace || !local || extra !== undefined) throw new Error(`invalid XML name ${name}`);
        return [namespace, local] as const;
      };
      const [namespace, local] = expand(open, true);
      const node: Element = { namespace, local, attributes: new Map(), parent, start, end, text: '' };
      for (const [name, value] of values) {
        if (name === 'xmlns' || name.startsWith('xmlns:')) continue;
        const key = qualified(...expand(name));
        if (node.attributes.has(key)) throw new Error(`duplicate attribute ${name}`);
        node.attributes.set(key, value);
      }
      elements.push(node);
      if (!empty) stack.push({ name: open, node, scope });
    } else if (text !== undefined || cdata !== undefined) {
      if (parent) parent.text += cdata ?? xmlValue(text);
      else if (cdata !== undefined || text.trim()) throw new Error('text outside document element');
    }
  }
  if (end !== xml.length) throw new Error(`invalid XML at offset ${end}`);
  if (stack.length || roots !== 1) throw new Error('missing or unclosed document element');
}

type Relationship = { target: string; type: string; external: boolean };

function targetPart(owner: string, target: string): string | undefined {
  try {
    const base = new URL(`https://package.invalid/${owner}`);
    const url = new URL(target, base);
    if (url.origin === base.origin) return decodeURIComponent(url.pathname.slice(1));
  } catch {
    return undefined;
  }
}

function inspectStructure(parts: Parts) {
  const structure = {
    malformedXml: [] as string[],
    unresolvedReferences: [] as string[],
    missingTargets: [] as string[],
    missingContentTypes: [] as string[],
    missingRequiredAttributes: [] as string[],
  };
  const decoder = new TextDecoder('utf-8', { fatal: true });
  const xml = new Map<string, string>();
  const elements = new Map<string, Element[]>();
  const relationships = new Map<string, Map<string, Relationship>>();
  for (const name of Object.keys(parts).sort()) {
    if (!/\.(?:xml|rels)$/i.test(name)) continue;
    const nodes: Element[] = [];
    elements.set(name, nodes);
    try {
      const text = decoder.decode(parts[name]);
      xml.set(name, text);
      readElements(text, nodes);
    } catch (error) {
      structure.malformedXml.push(`${name}: ${String(error)}`);
    }
  }

  const contentTypes = elements.get(TYPES) ?? [];
  const overrides = new Set<string>();
  const defaults = new Set<string>();
  for (const node of contentTypes) {
    if (!isElement(node.parent, NS.types, 'Types') || !attr(node, 'ContentType')) continue;
    const partName = attr(node, 'PartName');
    const extension = attr(node, 'Extension');
    if (isElement(node, NS.types, 'Override') && partName) overrides.add(partName);
    if (isElement(node, NS.types, 'Default') && extension) defaults.add(extension.toLowerCase());
  }
  for (const name of Object.keys(parts).sort()) {
    if (name === TYPES) continue;
    const extension = name.split('.').pop()!.toLowerCase();
    if (!overrides.has(`/${name}`) && !defaults.has(extension)) structure.missingContentTypes.push(name);
  }

  for (const [name, nodes] of elements) {
    const path = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(name);
    const owner = name === '_rels/.rels' ? '' : path && `${path[1] ?? ''}${path[2]}`;
    if (owner === null) continue;
    const rels = new Map<string, Relationship>();
    relationships.set(owner, rels);
    for (const node of nodes) {
      if (!isElement(node, NS.rels, 'Relationship')) continue;
      const id = attr(node, 'Id');
      const type = attr(node, 'Type') ?? '';
      const target = attr(node, 'Target') ?? '';
      const external = attr(node, 'TargetMode') === 'External';
      if (id) rels.set(id, { target, type, external });
      const resolved = target && targetPart(owner, target);
      if (!external && (!resolved || !Object.hasOwn(parts, resolved))) {
        structure.missingTargets.push(`${name}: ${id ?? '(no Id)'} → ${target || '(no Target)'}`);
      }
    }
  }

  const relationshipNamespace = `{${NS.r}}`;
  for (const [name, nodes] of elements) {
    for (const node of nodes) {
      for (const [key, id] of node.attributes) {
        const attribute = key.startsWith(relationshipNamespace)
          ? `r:${key.slice(relationshipNamespace.length)}`
          : isElement(node, NS.dsp, 'dataModelExt') && key === 'relId' ? key : null;
        if (attribute && !relationships.get(name)?.has(id)) {
          structure.unresolvedReferences.push(`${name}: ${node.local} @${attribute}="${id}"`);
        }
      }
      if (node.namespace !== NS.w) continue;
      const suffix = `${name}: w:${node.local}`;
      if (WORD_ID_REQUIRED.has(node.local) && attr(node, 'id', NS.w) === undefined) {
        structure.missingRequiredAttributes.push(`${suffix} @w:id`);
      }
      if (Object.hasOwn(MOVE_STARTS, node.local) && attr(node, 'author', NS.w) === undefined) {
        structure.missingRequiredAttributes.push(`${suffix} @w:author`);
      }
    }
  }
  return { xml, elements, relationships, structure };
}

function inspectSave(path: string, source: Parts, saved: Saved, markers: readonly string[]) {
  const parts = unzipContainer(saved.bytes);
  const changes = partChanges(source, parts);
  const { xml, elements, relationships, structure } = inspectStructure(parts);
  const documentXml = xml.get(BODY) ?? '';
  const body = elements.get(BODY) ?? [];
  const rels = relationships.get(BODY);
  const edited = body.find((node) =>
    isElement(node, NS.w, 'p') && attr(node, 'paraId', NS.w14) === EDIT_PARAGRAPH_ID);
  const outsideEdit = edited
    ? documentXml.slice(0, edited.start) + documentXml.slice(edited.end)
    : documentXml;
  const probes = Object.fromEntries(markers.map((marker) =>
    [marker, outsideEdit.includes(`"${marker}"`) || outsideEdit.includes(`'${marker}'`)]));
  const relIds = body.find((node) => isElement(node, NS.dgm, 'relIds') &&
    node.parent && isElement(node.parent, NS.a, 'graphicData') && attr(node.parent, 'uri') === NS.dgm);
  const resolves = (id: string | null, target: string, type: string) => {
    const rel = id && rels?.get(id);
    return !!rel && !rel.external && rel.type === `${NS.r}/${type}` &&
      targetPart(BODY, rel.target) === `word/${target}` && Object.hasOwn(parts, `word/${target}`);
  };
  const smartArt = {
    relIdsPresent: relIds !== undefined,
    relationships: Object.fromEntries(Object.entries(DIAGRAM_RELS).map(([attribute, expected]) => {
      const id = relIds ? attr(relIds, attribute, NS.r) ?? null : null;
      return [`r:${attribute}`, { id, resolves: resolves(id, expected.target, expected.type) }];
    })),
  };
  const moveRangeAuthors = Object.fromEntries(Object.entries(MOVE_STARTS).map(([name, id]) => {
    const start = body.find((node) => isElement(node, NS.w, name) && attr(node, 'id', NS.w) === id);
    const author = start ? attr(start, 'author', NS.w) : undefined;
    return [name, { present: !!start, hasAuthor: author !== undefined, author: author ?? null }];
  }));
  const extendedComments = Object.fromEntries(EXTENDED_COMMENT_PARTS.map((part) =>
    [part, changes.byteIdentical.includes(part)]));
  const documentReferences = Object.fromEntries(Object.entries(DOCUMENT_REFERENCES).map(
    ([name, [id, target]]) => [
      name,
      body.some((node) => isElement(node, name === 'chart' ? NS.c : NS.w, name) &&
        attr(node, 'id', NS.r) === id) &&
        resolves(id, target, name === 'chart' ? name : name.replace('Reference', '')),
    ]
  ));
  const editedText = edited
    ? body.filter((node) => isElement(node, NS.w, 't') && node.start > edited.start && node.end < edited.end)
      .map((node) => node.text).join('')
    : null;
  return {
    report: {
      path, ...changes, warnings: saved.warnings, probes, smartArt, moveRangeAuthors,
      extendedComments, documentReferences, structure,
    },
    parts,
    editedText,
  };
}

let fixture: ReturnType<typeof oneSavePackage>;
let sourceParts: Parts;

beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  );
  fixture = oneSavePackage();
  sourceParts = unzipContainer(fixture.bytes);
}, 30_000);

describe('one-save source package', () => {
  it('starts with valid structure and every probe present', () => {
    const { report, editedText } = inspectSave('source', sourceParts, {
      bytes: fixture.bytes, warnings: 'n/a',
    }, fixture.probes);
    expect(Object.values(report.structure).flat()).toEqual([]);
    expect(Object.entries(report.probes).filter(([, present]) => !present)).toEqual([]);
    expect(report.smartArt.relIdsPresent).toBe(true);
    expect(Object.values(report.smartArt.relationships).every((rel) => rel.resolves)).toBe(true);
    expect(Object.values(report.moveRangeAuthors).every((move) => move.hasAuthor)).toBe(true);
    expect(Object.values(report.extendedComments).every(Boolean)).toBe(true);
    expect(Object.values(report.documentReferences).every(Boolean)).toBe(true);
    expect(editedText).toBe('EDIT ME');
  });
});

function describeSave(
  path: string,
  save: Save,
  options: { secondPass?: boolean; returnsWarnings?: boolean } = {}
) {
  describe(path, () => {
    let inspected: ReturnType<typeof inspectSave>;
    let secondPass: ReturnType<typeof partChanges> | undefined;

    beforeAll(async () => {
      const saved = await save(fixture.bytes);
      inspected = inspectSave(path, sourceParts, saved, fixture.probes);
      if (options.secondPass) {
        const next = await save(saved.bytes);
        secondPass = partChanges(inspected.parts, unzipContainer(next.bytes));
      }
    }, 30_000);

    it('keeps unchanged parts byte-identical', () => {
      expect(inspected.report.changed.filter((part) => part !== CORE)).toEqual([]);
    });

    it('adds no parts', () => {
      expect(inspected.report.added).toEqual([]);
    });

    it('removes no parts', () => {
      expect(inspected.report.removed).toEqual([]);
    });

    for (const marker of ONE_SAVE_PROBES) {
      it(`preserves probe ${marker}`, () => {
        expect(inspected.report.probes[marker]).toBe(true);
      });
    }

    it('keeps the SmartArt relationship element in the body', () => {
      expect(inspected.report.smartArt.relIdsPresent).toBe(true);
    });

    for (const [attribute, { id }] of Object.entries(DIAGRAM_RELS)) {
      it(`keeps SmartArt r:${attribute} and resolves it to its diagram part`, () => {
        expect(inspected.report.smartArt.relationships[`r:${attribute}`])
          .toEqual({ id, resolves: true });
      });
    }

    for (const name of Object.keys(MOVE_STARTS)) {
      it(`keeps the author on w:${name}`, () => {
        expect(inspected.report.moveRangeAuthors[name])
          .toEqual({ present: true, hasAuthor: true, author: 'Author' });
      });
    }

    for (const part of EXTENDED_COMMENT_PARTS) {
      it(`keeps ${part} byte-identical`, () => {
        expect(inspected.report.extendedComments[part]).toBe(true);
      });
    }

    for (const name of Object.keys(DOCUMENT_REFERENCES)) {
      it(`keeps the body’s ${name} reference and target`, () => {
        expect(inspected.report.documentReferences[name]).toBe(true);
      });
    }

    it('writes well-formed XML in every XML part', () => {
      expect(inspected.report.structure.malformedXml).toEqual([]);
    });

    it('resolves relationship attributes against the owning part’s relationships', () => {
      expect(inspected.report.structure.unresolvedReferences).toEqual([]);
    });

    it('keeps every internal relationship target in the package', () => {
      expect(inspected.report.structure.missingTargets).toEqual([]);
    });

    it('assigns every part a content type', () => {
      expect(inspected.report.structure.missingContentTypes).toEqual([]);
    });

    it('keeps required WordprocessingML attributes', () => {
      expect(inspected.report.structure.missingRequiredAttributes).toEqual([]);
    });

    if (options.returnsWarnings) {
      it('reports detected loss in its warnings', () => {
        const report = inspected.report;
        const loss = [
          ...report.removed,
          ...Object.keys(report.probes).filter((marker) => !report.probes[marker]),
          ...Object.keys(report.smartArt.relationships).filter((key) =>
            !report.smartArt.relationships[key].resolves),
          ...Object.keys(report.moveRangeAuthors).filter((key) =>
            !report.moveRangeAuthors[key].hasAuthor),
          ...Object.keys(report.documentReferences).filter((key) => !report.documentReferences[key]),
          ...report.structure.missingRequiredAttributes,
        ];
        expect(Array.isArray(report.warnings)).toBe(true);
        expect(report.warnings === 'n/a' || report.warnings.length === 0 ? loss : []).toEqual([]);
      });
    }

    if (options.secondPass) {
      it('keeps package parts stable on a second pass apart from core properties', () => {
        expect(secondPass && {
          changed: secondPass.changed.filter((part) => part !== CORE),
          added: secondPass.added,
          removed: secondPass.removed,
        }).toEqual({ changed: [], added: [], removed: [] });
      });
    }
  });
}

async function withSession<T>(bytes: Uint8Array, save: (session: YrsSession) => Promise<T>) {
  const session = await createYrsSession();
  try {
    session.seedFromDocx(bytes);
    return await save(session);
  } finally {
    session.destroy();
  }
}

describeSave('parseDocx -> repackDocx', async (bytes) => {
  const document = await parseDocx(bytes, { preloadFonts: false });
  return { bytes: new Uint8Array(await repackDocx(document)), warnings: 'n/a' };
}, { secondPass: true });

describeSave('parseDocx -> repackDocxWithWarnings', async (bytes) => {
  const document = await parseDocx(bytes, { preloadFonts: false });
  const { buffer, warnings } = await repackDocxWithWarnings(document);
  return { bytes: new Uint8Array(buffer), warnings };
}, { secondPass: true, returnsWarnings: true });

describeSave('seedFromDocx -> saveYrsDocx (unedited)', (bytes) =>
  withSession<Saved>(bytes, async (session) => ({
    bytes: (await saveYrsDocx(session)).bytes, warnings: 'n/a',
  }))
);

describeSave('seedFromDocx -> yrsToDocument -> repackDocx', (bytes) =>
  withSession<Saved>(bytes, async (session) => {
    const base = session.materializeDocx();
    if (!base) throw new Error('expected a document from the DOCX session');
    const document = yrsToDocument(session, base);
    return { bytes: new Uint8Array(await repackDocx(document)), warnings: 'n/a' };
  })
);

describeSave('parseDocx -> createDocx with originalBuffer', async (bytes) => {
  const document = await parseDocx(bytes, { preloadFonts: false });
  return { bytes: new Uint8Array(await createDocx(document)), warnings: 'n/a' };
});

function firstText(blocks: BlockContent[], paraId?: string) {
  for (const block of blocks) {
    if (block.type !== 'paragraph' || (paraId !== undefined && block.paraId !== paraId)) continue;
    for (const inline of block.content) {
      if (inline.type !== 'run') continue;
      const text = inline.content.find((node) => node.type === 'text');
      if (text) return text;
    }
  }
  throw new Error('Expected a text run in the synthetic package');
}

it('writes caller edits in the body, header and comment while preserving the other parts', async () => {
  const document = await parseDocx(fixture.bytes, { preloadFonts: false });
  const pkg = document.package;
  firstText(pkg.document.content, EDIT_PARAGRAPH_ID).text = 'Body by model';
  firstText(pkg.headers!.get('rIdHeader')!.content).text = 'Header by model';
  firstText(pkg.document.comments![0]!.content).text = 'Comment by model';

  const buffer = await repackDocx(document);
  const changes = partChanges(sourceParts, unzipContainer(new Uint8Array(buffer)));
  expect(changes.changed.filter((part) => part !== CORE)).toEqual([
    'word/comments.xml', BODY, 'word/header1.xml',
  ]);
  expect(changes.added).toEqual([]);
  expect(changes.removed).toEqual([]);

  const { package: reopened } = await parseDocx(buffer, { preloadFonts: false });
  expect(firstText(reopened.document.content, EDIT_PARAGRAPH_ID).text).toBe('Body by model');
  expect(firstText(reopened.headers!.get('rIdHeader')!.content).text).toBe('Header by model');
  expect(firstText(reopened.document.comments![0]!.content).text).toBe('Comment by model');
});

it('names the rewritten body and preserves every other package part', async () => {
  const document = await parseDocx(fixture.bytes, { preloadFonts: false });
  firstText(document.package.document.content, EDIT_PARAGRAPH_ID).text = 'Body edit';
  const { buffer, warnings } = await repackDocxWithWarnings(document, { updateModifiedDate: false });
  expect(warnings.filter(warning => warning.startsWith('DOCX_PART_REWRITTEN '))).toEqual([
    expect.stringMatching(/^DOCX_PART_REWRITTEN word\/document\.xml .*saveYrsDocx/),
  ]);
  expect(partChanges(sourceParts, unzipContainer(new Uint8Array(buffer)))).toEqual({
    byteIdentical: Object.keys(sourceParts).filter(path => path !== BODY).sort(),
    changed: [BODY], added: [], removed: [],
  });
});

const DOCUMENT_RELS = 'word/_rels/document.xml.rels';
const HEADER = 'word/header1.xml';
const FOOTNOTES = 'word/footnotes.xml';
const STYLES = 'word/styles.xml';
const decodePart = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
const encodePart = (xml: string) => new TextEncoder().encode(xml);
const noStamp = { updateModifiedDate: false };

function changedPart(parts: Parts, path: string, change: (xml: string) => string) {
  parts[path] = encodePart(change(decodePart(parts[path])));
}

it('preserves single-quoted relationship IDs and content type registrations', async () => {
  const parts = { ...sourceParts };
  changedPart(parts, DOCUMENT_RELS, xml => xml.replace('Id="rIdHeader"', "Id='rIdHeader'"));
  changedPart(parts, TYPES, xml => xml.replace('PartName="/word/header1.xml"', "PartName='/word/header1.xml'"));
  const source = rezipContainer(parts);
  const document = await parseDocx(source, { preloadFonts: false });
  const saved = await repackDocx(document, noStamp);
  expect(unzipContainer(new Uint8Array(saved))).toEqual(parts);
  firstText(document.package.document.content, EDIT_PARAGRAPH_ID).text = 'Edit beside quoted parts';
  const edited = unzipContainer(new Uint8Array(await repackDocx(document, noStamp)));
  expect(edited[DOCUMENT_RELS]).toEqual(parts[DOCUMENT_RELS]);
  expect(edited[TYPES]).toEqual(parts[TYPES]);
});

it('rejects conflicting edits through header aliases and accepts equivalent edits', async () => {
  const parts = { ...sourceParts };
  changedPart(parts, DOCUMENT_RELS, xml => xml.replace('</Relationships>',
    `<Relationship Id="rIdAlias" Type="${NS.r}/header" Target="HEADER1.XML"/></Relationships>`));
  parts[HEADER] = encodePart(`<w:hdr xmlns:w="${NS.w}" xmlns:w14="${NS.w14}"><w:p w14:paraId="10000001" w14:textId="invalid"><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`);
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  const first = firstText(document.package.headers!.get('rIdHeader')!.content);
  const alias = firstText(document.package.headers!.get('rIdAlias')!.content);
  first.text = 'First edit';
  alias.text = 'Conflicting edit';
  await expect(repackDocx(document, noStamp)).rejects.toThrow('conflicting header/footer edits to word/header1.xml');
  alias.text = first.text;
  const saved = await repackDocx(document, noStamp);
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(firstText(reopened.package.headers!.get('rIdAlias')!.content).text).toBe('First edit');
  expect(partChanges(parts, unzipContainer(new Uint8Array(saved))).changed).toEqual([HEADER]);
});

it('keeps an edited header when a footer relationship also targets its part', async () => {
  const parts = { ...sourceParts };
  changedPart(parts, DOCUMENT_RELS, xml => xml.replace('</Relationships>',
    `<Relationship Id="rIdFooterAlias" Type="${NS.r}/footer" Target="header1.xml"/></Relationships>`));
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  expect(document.package.footers?.get('rIdFooterAlias')).toBeDefined();
  firstText(document.package.headers!.get('rIdHeader')!.content).text = 'Header owns the edit';
  const saved = await repackDocx(document, noStamp);
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(firstText(reopened.package.headers!.get('rIdHeader')!.content).text).toBe('Header owns the edit');
  expect(partChanges(parts, unzipContainer(new Uint8Array(saved))).changed).toEqual([HEADER]);
});

it('writes an edited header and note using their original archive spelling', async () => {
  const parts = { ...sourceParts };
  for (const [from, to] of [[BODY, 'word/Document.xml'], [HEADER, 'word/Header1.xml'], [FOOTNOTES, 'word/Footnotes.xml']]) {
    parts[to] = parts[from];
    delete parts[from];
  }
  const source = rezipContainer(parts);
  const document = await parseDocx(source, { preloadFonts: false });
  expect(unzipContainer(new Uint8Array(await repackDocx(document, noStamp)))).toEqual(parts);
  firstText(document.package.headers!.get('rIdHeader')!.content).text = 'Header edit';
  const note = document.package.footnotes![0]!;
  firstText(note.content).text = 'Note edit';
  const saved = await repackDocx(document, noStamp);
  const savedParts = unzipContainer(new Uint8Array(saved));
  expect(partChanges(parts, savedParts)).toMatchObject({
    changed: ['word/Footnotes.xml', 'word/Header1.xml'], added: [], removed: [],
  });
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(firstText(reopened.package.headers!.get('rIdHeader')!.content).text).toBe('Header edit');
  expect(firstText(reopened.package.footnotes![0]!.content).text).toBe('Note edit');
});

it('writes changes to unused custom document namespace bindings', async () => {
  const parts = { ...sourceParts };
  changedPart(parts, BODY, xml => xml.replace(/<w:document(?=\s)/, '<w:document xmlns:unused="urn:before"'));
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  document.package.document.customRootBindings!.find(binding => binding.name === 'xmlns:unused')!.value = 'urn:after';
  const saved = await repackDocx(document, noStamp);
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(reopened.package.document.customRootBindings).toContainEqual({ name: 'xmlns:unused', value: 'urn:after' });
});

function bookmarkedNotePackage() {
  return rezipContainer({
    ...sourceParts,
    [FOOTNOTES]: encodePart(`<w:footnotes xmlns:w="${NS.w}"><w:footnote w:id="73"><w:bookmarkStart w:id="7" w:name="note"/><w:p><w:r><w:t>Original note</w:t></w:r></w:p><w:bookmarkEnd w:id="7"/></w:footnote></w:footnotes>`),
  });
}

it('uses the partial parse coverage when comparing notes', async () => {
  const parts = unzipContainer(bookmarkedNotePackage());
  const id = '10000001';
  changedPart(parts, HEADER, xml => xml.replace(/<w:p[ >]/, match =>
    match.replace('<w:p', `<w:p xmlns:w14="${NS.w14}" w14:textId="invalid"`)));
  changedPart(parts, FOOTNOTES, xml => xml.replace('<w:p>',
    `<w:p xmlns:w14="${NS.w14}" w14:paraId="${id}" w14:textId="invalid">`));
  const document = await parseDocx(rezipContainer(parts), {
    preloadFonts: false, parseHeadersFooters: false,
  });
  firstText(document.package.document.content, EDIT_PARAGRAPH_ID).text = 'Partial parse edit';
  const saved = unzipContainer(new Uint8Array(await repackDocx(document, noStamp)));
  expect(saved[HEADER]).toEqual(parts[HEADER]);
  expect(saved[FOOTNOTES]).toEqual(parts[FOOTNOTES]);
});

it('writes note model edits when the source note moves to the separator collection', async () => {
  const document = await parseDocx(bookmarkedNotePackage(), { preloadFonts: false });
  const note = document.package.footnotes!.pop()!;
  expect(note.verbatimXml).toContain('Original note');
  note.noteType = 'separator';
  firstText(note.content).text = 'Edited separator';
  document.package.footnoteSeparators!.push(note);
  const saved = await repackDocx(document, noStamp);
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(firstText(reopened.package.footnoteSeparators![0]!.content).text).toBe('Edited separator');
  expect(reopened.package.footnotes).toEqual([]);
});

it('writes a copied note under its model ID and text when its verbatim XML is foreign', async () => {
  const original = await parseDocx(bookmarkedNotePackage(), { preloadFonts: false });
  const foreign = original.package.footnotes![0]!;
  expect(foreign.verbatimXml).toContain('w:id="73"');
  const document = await parseDocx(fixture.bytes, { preloadFonts: false });
  foreign.id = 74;
  firstText(foreign.content).text = 'Copied note';
  document.package.footnotes!.push(foreign);
  const saved = await repackDocx(document, noStamp);
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(firstText(reopened.package.footnotes!.find(note => note.id === 74)!.content).text).toBe('Copied note');
  expect(reopened.package.footnotes!.some(note => note.id === 73)).toBe(false);
});

it('preserves header, footer and note parts omitted by a partial parse', async () => {
  const document = await parseDocx(fixture.bytes, {
    preloadFonts: false,
    parseHeadersFooters: false,
    parseNotes: false,
  });
  const bytes = new Uint8Array(await repackDocx(document, { updateModifiedDate: false }));
  expect(unzipContainer(bytes)).toEqual(sourceParts);
});

function editBody(session: YrsSession, text: string): void {
  const result = session.applyEdits({
    expectVersion: session.version(),
    steps: [{
      op: 'insertText',
      target: { kind: 'paragraph', story: 'body', paraId: EDIT_PARAGRAPH_ID },
      at: 'end',
      text,
    }],
  });
  expect(result).toMatchObject({ ok: true, applied: true });
}

it('saves an edited projection based on parseDocx with the session plan', () =>
  withSession(fixture.bytes, async (session) => {
    editBody(session, EDIT);
    const base = await parseDocx(fixture.bytes, { preloadFonts: false });
    const projected = yrsToDocument(session, base);
    const options = { updateModifiedDate: false };
    const { buffer, warnings } = await repackDocxWithWarnings(projected, options);
    expect(warnings.some(warning => warning.startsWith("DOCX_PART_REWRITTEN "))).toBe(false);
    const bytes = new Uint8Array(buffer);
    expect(unzipContainer(bytes)).toEqual(unzipContainer((await saveYrsDocx(session, options)).bytes));
    expect(inspectSave('edited projection', sourceParts, { bytes, warnings: 'n/a' }, fixture.probes).editedText)
      .toBe(`EDIT ME${EDIT}`);
  })
);

it('saves an old projection without writing a later session edit', () =>
  withSession(fixture.bytes, async (session) => {
    editBody(session, EDIT);
    const projected = yrsToDocument(session, session.materializeDocx()!);
    editBody(session, ' (later)');
    const bytes = new Uint8Array(await repackDocx(projected));
    expect(inspectSave('stale projection', sourceParts, { bytes, warnings: 'n/a' }, fixture.probes).editedText)
      .toBe(`EDIT ME${EDIT}`);
  })
);

it('writes an in-place edit to a current projection', () =>
  withSession(fixture.bytes, async (session) => {
    const projected = yrsToDocument(session, session.materializeDocx()!);
    firstText(projected.package.document.content, EDIT_PARAGRAPH_ID).text = 'Edited projection';
    const bytes = new Uint8Array(await repackDocx(projected));
    expect(inspectSave('mutated projection', sourceParts, { bytes, warnings: 'n/a' }, fixture.probes).editedText)
      .toBe('Edited projection');
  })
);


it('keeps the same projection current after publishing paragraph IDs', () => {
  const parts = { ...sourceParts };
  parts[BODY] = encodePart(`<w:document xmlns:w="${NS.w}" xmlns:w14="${NS.w14}"><w:body><w:p w14:paraId="${EDIT_PARAGRAPH_ID}"><w:r><w:t>Start</w:t></w:r></w:p></w:body></w:document>`);
  parts[STYLES] = encodePart(`<w:styles xmlns:w="${NS.w}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
  return withSession(rezipContainer(parts), async session => {
    const result = session.applyEdits({
      expectVersion: session.version(),
      steps: [{ op: 'insertParagraphs', target: { story: 'body', paraId: EDIT_PARAGRAPH_ID },
        at: 'end', paragraphs: [{ text: 'New paragraph' }] }],
    });
    expect(result).toMatchObject({ ok: true, applied: true });
    expect(session.setStyleLanguage('Normal', { lang: 'fr-FR' })).toBe(true);
    const range = { story: 'body', start: { paraId: EDIT_PARAGRAPH_ID, offset: 0 }, end: { paraId: EDIT_PARAGRAPH_ID, offset: 0 } };
    const list = session.ensureListDefinition('body', EDIT_PARAGRAPH_ID, 'bullet');
    session.setParagraphAttrs(range, list);
    const projected = yrsToDocument(session, session.materializeDocx()!);
    const version = session.version();
    const first = await repackDocxWithWarnings(projected, noStamp);
    expect(session.version()).not.toBe(version);
    const second = await repackDocxWithWarnings(projected, noStamp);
    const saved = unzipContainer(new Uint8Array(first.buffer));
    expect(decodePart(saved[BODY])).toContain('New paragraph');
    expect(decodePart(saved[STYLES])).toContain('fr-FR');
    expect(decodePart(saved['word/numbering.xml'])).toContain('betteroffice-list:');
    expect(unzipContainer(new Uint8Array(second.buffer))).toEqual(saved);
    expect(second.warnings.some(warning => warning.startsWith('DOCX_PART_REWRITTEN '))).toBe(false);
  });
});

it('uses captured style language and lists when the session changes during the write', () => {
  const parts = { ...sourceParts, [STYLES]: encodePart(
    `<w:styles xmlns:w="${NS.w}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`
  ) };
  return withSession(rezipContainer(parts), async session => {
    expect(session.setStyleLanguage('Normal', { lang: 'fr-FR' })).toBe(true);
    const range = { story: 'body', start: { paraId: EDIT_PARAGRAPH_ID, offset: 0 }, end: { paraId: EDIT_PARAGRAPH_ID, offset: 0 } };
    const before = session.ensureListDefinition('body', EDIT_PARAGRAPH_ID, 'bullet');
    session.setParagraphAttrs(range, before);
    const [originalList] = session.createdListDefinitions();
    expect(originalList).toBeDefined();
    const document = yrsToDocument(session, session.materializeDocx()!);
    const capture = captureSessionSave(session);
    let newListKey: string | undefined;
    const saved = await writeSessionSave(
      session, document, capture, document.originalBuffer!, noStamp,
      () => true, false, () => true, async () => {
        if (newListKey !== undefined) return;
        expect(session.setStyleLanguage('Normal', { lang: 'de-DE' })).toBe(true);
        const later = session.ensureListDefinition('body', EDIT_PARAGRAPH_ID, 'numbered');
        session.setParagraphAttrs(range, later);
        newListKey = session.createdListDefinitions().find(definition => definition.key !== originalList!.key)?.key;
        expect(newListKey).toBeDefined();
      }
    );
    const written = unzipContainer(saved.bytes);
    const styles = decodePart(written[STYLES]);
    expect(styles).toContain('fr-FR');
    expect(styles).not.toContain('de-DE');
    const numbering = decodePart(written['word/numbering.xml']);
    expect(numbering).toContain(`betteroffice-list:${originalList!.key}`);
    expect(numbering).not.toContain(`betteroffice-list:${newListKey}`);
  });
});

it('keeps bookmarked notes and valid revision IDs when saving a cloned projection from raw', () => {
  const source = bookmarkedNotePackage();
  return withSession(source, async session => {
    const result = session.applyEdits({
      expectVersion: session.version(),
      steps: [{ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: EDIT_PARAGRAPH_ID },
        at: 'end', text: ' new revision', suggest: { author: 'Editor', date: '2026-01-02T03:04:05Z' } }],
    });
    expect(result).toMatchObject({ ok: true, applied: true });
    const projected = yrsToDocument(session, await parseDocx(source, { preloadFonts: false }));
    const clone = structuredClone(projected);
    const raw = { originalBuffer: Uint8Array.from(source).buffer } as RawDocxContent;
    const saved = await repackDocxFromRaw(clone, raw, noStamp);
    const savedParts = unzipContainer(new Uint8Array(saved));
    expect(savedParts[FOOTNOTES]).toEqual(unzipContainer(source)[FOOTNOTES]);
    const revision = decodePart(savedParts[BODY]).match(/<w:ins\b[^>]*w:id="(\d+)"[^>]*w:author="Editor"/);
    expect(revision).not.toBeNull();
    expect(Number(revision![1])).toBeLessThan(2 ** 31);
    expect(decodePart(savedParts[BODY])).toContain('new revision');
  });
});

it('keeps the fresh create path when the model has unrelated relationship IDs', async () => {
  const document = await parseDocx(fixture.bytes, { preloadFonts: false });
  document.originalBuffer = undefined;
  document.package.relationships = new Map([
    ['rId1', { id: 'rId1', type: `${NS.r}/numbering`, target: 'numbering.xml' }],
  ]);
  document.package.document = { content: [{
    type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Fresh text' }] }],
  }] };
  const saved = unzipContainer(new Uint8Array(await createDocx(document)));
  expect(decodePart(saved[BODY])).toContain('Fresh text');
});

it('keeps the original Windows-1252 body bytes when only a header changes', async () => {
  const body = `<?xml version="1.0" encoding="windows-1252"?><w:document xmlns:w="${NS.w}"><w:body><w:p><w:r><w:t>Café</w:t></w:r></w:p></w:body></w:document>`;
  const parts = { ...sourceParts, [BODY]: Uint8Array.from(body, char => char.charCodeAt(0)) };
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  firstText(document.package.headers!.get('rIdHeader')!.content).text = 'Header edit';
  const saved = unzipContainer(new Uint8Array(await repackDocx(document, noStamp)));
  expect(saved[BODY]).toEqual(parts[BODY]);
});

it('keeps BOM-prefixed relationships and content types after a body edit', async () => {
  const parts = { ...sourceParts };
  for (const part of [TYPES, DOCUMENT_RELS]) changedPart(parts, part, xml => `\uFEFF${xml}`);
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  firstText(document.package.document.content, EDIT_PARAGRAPH_ID).text = 'Body edit';
  const saved = unzipContainer(new Uint8Array(await repackDocx(document, noStamp)));
  for (const part of [TYPES, DOCUMENT_RELS]) expect(saved[part]).toEqual(parts[part]);
});

it('empties a note part when the model explicitly removes its last note', async () => {
  const source = bookmarkedNotePackage();
  const document = await parseDocx(source, { preloadFonts: false });
  document.package.footnotes = [];
  const { buffer, warnings } = await repackDocxWithWarnings(document, noStamp);
  const saved = unzipContainer(new Uint8Array(buffer));
  expect(saved[BODY]).toEqual(unzipContainer(source)[BODY]);
  expect(decodePart(saved[FOOTNOTES])).toContain('</w:footnotes>');
  expect(decodePart(saved[FOOTNOTES])).not.toContain('<w:footnote ');
  expect(warnings.some(warning => warning.startsWith(`DOCX_PART_REWRITTEN ${FOOTNOTES} `))).toBe(true);
  expect(warnings.some(warning => warning.startsWith('DOCX_WHOLE_MODEL_WRITE '))).toBe(false);
});

it('keeps an explicit raw bookmark attribute on a note with unchanged text', async () => {
  const document = await parseDocx(bookmarkedNotePackage(), { preloadFonts: false });
  const note = document.package.footnotes![0]!;
  note.verbatimXml = note.verbatimXml!.replace('w:name="note"', 'w:name="note" w:displacedByCustomXml="next"');
  const saved = unzipContainer(new Uint8Array(await repackDocx(document, noStamp)));
  expect(decodePart(saved[FOOTNOTES])).toContain('w:displacedByCustomXml="next"');
});

it('cleans deleted comment markers from headers with mixed-case archive names', async () => {
  const parts = { ...sourceParts };
  const original = await parseDocx(fixture.bytes, { preloadFonts: false });
  const id = original.package.document.comments![0]!.id;
  changedPart(parts, HEADER, xml => xml.replace(/<w:p(?:\s[^>]*)?>/, tag =>
    `${tag}<w:commentRangeStart w:id="${id}"/>`).replace('</w:p>',
    `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p>`));
  parts['word/Header1.xml'] = parts[HEADER];
  delete parts[HEADER];
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  document.package.document.comments = [];
  const saved = unzipContainer(new Uint8Array(await repackDocx(document, noStamp)));
  expect(decodePart(saved['word/Header1.xml'])).not.toMatch(/<w:comment(?:RangeStart|RangeEnd|Reference)\b/);
});

it('saves a current projection of a thin host envelope with source numbering', async () => {
  const source = rezipContainer({
    ...sourceParts,
    'word/numbering.xml': encodePart(`<w:numbering xmlns:w="${NS.w}"><w:abstractNum w:abstractNumId="2"/><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>`),
  });
  const session = await createYrsSession();
  try {
    const host = session.seedFromDocx(source);
    editBody(session, EDIT);
    const projection = yrsToDocument(session, host.document);
    const { buffer, warnings } = await repackDocxWithWarnings(projection, noStamp);
    const saved = unzipContainer(new Uint8Array(buffer));
    expect(decodePart(saved[BODY])).toContain(EDIT);
    expect(saved['word/numbering.xml']).toEqual(unzipContainer(source)['word/numbering.xml']);
    expect(warnings.some(w => /^DOCX_(?:KEPT_|STRUCTURE_DROPPED|MINIMAL_PATCH replace counts)/.test(w))).toBe(false);
  } finally {
    session.destroy();
  }
});

it('keeps an explicitly changed revision ID on a public session projection', () =>
  withSession(fixture.bytes, async session => {
    const result = session.applyEdits({
      expectVersion: session.version(),
      steps: [{ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: EDIT_PARAGRAPH_ID },
        at: 'end', text: ' New revision', suggest: { author: 'Editor', date: '2026-01-02T03:04:05Z' } }],
    });
    expect(result.ok).toBe(true);
    const projection = yrsToDocument(session, session.materializeDocx()!);
    for (const block of projection.package.document.content) {
      if (block.type !== 'paragraph') continue;
      for (const inline of block.content) {
        if (inline.type === 'insertion' && inline.info.author === 'Editor') inline.info.id = 12345;
      }
    }
    const saved = unzipContainer(new Uint8Array(await repackDocx(projection, noStamp)));
    expect(decodePart(saved[BODY])).toMatch(/<w:ins\b[^>]*w:id="12345"[^>]*w:author="Editor"/);
  })
);

it('does not create a relationship part for a preserved header using the main relationships', async () => {
  const parts = { ...sourceParts };
  delete parts['word/_rels/header1.xml.rels'];
  changedPart(parts, DOCUMENT_RELS, xml => xml.replace('</Relationships>',
    `<Relationship Id="rIdSharedLink" Type="${NS.r}/hyperlink" Target="https://example.test" TargetMode="External"/></Relationships>`));
  changedPart(parts, HEADER, xml => xml.replace('</w:p>',
    '<w:hyperlink r:id="rIdSharedLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>'));
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  firstText(document.package.document.content, EDIT_PARAGRAPH_ID).text = 'Body edit';
  const saved = unzipContainer(new Uint8Array(await repackDocx(document, noStamp)));
  expect(saved[HEADER]).toEqual(parts[HEADER]);
  expect(saved['word/_rels/header1.xml.rels']).toBeUndefined();
});

it('does not report a rewrite when setting empty comments on a package without comments', async () => {
  const parts = { ...sourceParts };
  for (const path of ['word/comments.xml', ...EXTENDED_COMMENT_PARTS]) delete parts[path];
  const source = rezipContainer(parts);
  const document = await parseDocx(source, { preloadFonts: false });
  document.package.document.comments = [];
  const { buffer, warnings } = await repackDocxWithWarnings(document, noStamp);
  expect(new Uint8Array(buffer)).toEqual(source);
  expect(warnings.some(warning => warning.startsWith('DOCX_PART_REWRITTEN '))).toBe(false);
});

it('uses the session save plan with a separate byte-identical raw buffer', () =>
  withSession(fixture.bytes, async session => {
    editBody(session, EDIT);
    const projection = yrsToDocument(session, session.materializeDocx()!);
    const raw = { originalBuffer: Uint8Array.from(fixture.bytes).buffer } as RawDocxContent;
    const saved = await repackDocxFromRaw(projection, raw, noStamp);
    expect(unzipContainer(new Uint8Array(saved))).toEqual(unzipContainer((await saveYrsDocx(session, noStamp)).bytes));
  })
);

it('removes comment markers from a Windows-1252 main part without changing its text', async () => {
  const original = await parseDocx(fixture.bytes, { preloadFonts: false });
  const id = original.package.document.comments![0]!.id;
  const body = `<?xml version="1.0" encoding="windows-1252"?><w:document xmlns:w="${NS.w}"><w:body><w:p><w:commentRangeStart w:id="${id}"/><w:r><w:t>Café</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p></w:body></w:document>`;
  const parts = { ...sourceParts, [BODY]: Uint8Array.from(body, char => char.charCodeAt(0)) };
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  document.package.document.comments = [];
  const saved = unzipContainer(new Uint8Array(await repackDocx(document, noStamp)));
  const xml = new TextDecoder('windows-1252').decode(saved[BODY]);
  expect(xml).toContain('encoding="windows-1252"');
  expect(xml).toContain('<w:t>Café</w:t>');
  expect(xml).not.toMatch(/<w:comment(?:RangeStart|RangeEnd|Reference)\b/);
});

it('reports a conflicting verbatim note edit with its note ID', async () => {
  const document = await parseDocx(bookmarkedNotePackage(), { preloadFonts: false });
  const note = document.package.footnotes![0]!;
  note.verbatimXml = note.verbatimXml!.replace('Original note', 'Conflicting note');
  await expect(repackDocx(document, noStamp))
    .rejects.toThrow('footnote 73 in word/footnotes.xml has verbatimXml that conflicts with its content');
});

it('keeps a cleared verbatim note cleared when another note changes', async () => {
  const parts = unzipContainer(bookmarkedNotePackage());
  changedPart(parts, FOOTNOTES, xml => xml
    .replace('<w:p>', '<w:customXml w:element="note" w:uri="urn:note"><w:p>')
    .replace('</w:p>', '</w:p></w:customXml>')
    .replace('</w:footnotes>', '<w:footnote w:id="74"><w:p><w:r><w:t>Other note</w:t></w:r></w:p></w:footnote></w:footnotes>'));
  const doc = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  const [first, second] = doc.package.footnotes!;
  expect(first!.verbatimXml).toContain('<w:customXml');
  delete first!.verbatimXml;
  firstText(second!.content).text = 'Changed note';
  const saved = unzipContainer(new Uint8Array(await repackDocx(doc, noStamp)));
  expect(decodePart(saved[FOOTNOTES])).not.toContain('<w:customXml');
  expect(decodePart(saved[FOOTNOTES])).toContain('Changed note');
});

it('keeps document bindings from a foreign thin envelope in a raw repack', async () => {
  const parts = { ...sourceParts };
  changedPart(parts, BODY, xml => xml.replace(/<w:document(?=\s)/, '<w:document xmlns:foreign="urn:foreign"'));
  await withSession(fixture.bytes, async session => {
    const foreign = await createYrsSession();
    try {
      const base = foreign.seedFromDocx(rezipContainer(parts)).document;
      const projected = yrsToDocument(session, base);
      const raw = { originalBuffer: Uint8Array.from(fixture.bytes).buffer } as RawDocxContent;
      const saved = unzipContainer(new Uint8Array(await repackDocxFromRaw(projected, raw, noStamp)));
      expect(decodePart(saved[BODY])).toContain('xmlns:foreign="urn:foreign"');
    } finally {
      foreign.destroy();
    }
  });
});

it('returns dangling drawing warnings on an unchanged save', async () => {
  const parts = { ...sourceParts };
  changedPart(parts, BODY, xml => xml.replace('</w:p>',
    '<w:r><w:object xmlns:o="urn:schemas-microsoft-com:office:office"><o:OLEObject r:id="rIdMissing"/></w:object></w:r></w:p>'));
  const doc = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  const unchanged = await repackDocxWithWarnings(doc, noStamp);
  const compared = await repackDocxWithWarnings(structuredClone(doc), noStamp);
  expect(unchanged.warnings).toEqual(compared.warnings);
  expect(unchanged.warnings).toContain('opaque drawing references relationship "rIdMissing" absent from word/_rels/document.xml.rels; keeping the markup verbatim');
});

it('rejects a replacement numbering definition before rewriting the body', async () => {
  const parts = {
    ...sourceParts,
    'word/numbering.xml': encodePart(`<w:numbering xmlns:w="${NS.w}"><w:abstractNum w:abstractNumId="2"/><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>`),
  };
  const document = await parseDocx(rezipContainer(parts), { preloadFonts: false });
  document.package.numbering!.nums[0]!.abstractNumId = 3;
  await expect(repackDocx(document, noStamp)).rejects.toThrow('cannot replace existing numbering definitions');
});

it('rejects cycles before comparing a parsed model with its save fingerprint', async () => {
  const document = await parseDocx(fixture.bytes, { preloadFonts: false });
  const body = document.package.document;
  Object.assign(body, { cycle: body });
  await expect(repackDocx(document, noStamp)).rejects.toThrow('save.document.cycle contains a cycle');
});
