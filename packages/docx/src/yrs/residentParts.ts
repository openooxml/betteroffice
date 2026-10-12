import type { Document, SectionProperties, HeaderFooterType, HeaderFooter, Footnote, Endnote, Relationship } from '../types/document';
import type { ResidentEngineSession } from './residentEngineSession';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { resolvedFinalSectionProperties, updateFinalSectionProperties } from '../editor/finalSection';

export interface DocxPartTarget {
  kind: 'header' | 'footer' | 'footnote' | 'endnote';
  section: string;
  page: number;
  variant?: HeaderFooterType;
  noteId?: number;
  expectVersion: string;
}

export interface ResidentPartTarget extends Omit<DocxPartTarget, 'expectVersion'> { story: string }
export interface ResidentPartGeometry {
  kind: DocxPartTarget['kind'];
  story: string;
  pageIndex: number;
  pageId: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ResidentPartMetadata {
  body: Pick<Document['package']['document'], 'sections' | 'finalSectionProperties'>;
  headers?: [string, HeaderFooter][];
  footers?: [string, HeaderFooter][];
  footnotes?: Footnote[];
  endnotes?: Endnote[];
  relationships?: [string, Relationship][];
}

export function partMetadata(document: Document): ResidentPartMetadata {
  const pkg = document.package;
  return { body: { sections: pkg.document.sections?.map(section => ({ ...section, content: [] })), finalSectionProperties: pkg.document.finalSectionProperties },
    headers: pkg.headers ? [...pkg.headers].map(([id, part]) => [id, { ...part, content: [] }]) : undefined,
    footers: pkg.footers ? [...pkg.footers].map(([id, part]) => [id, { ...part, content: [] }]) : undefined,
    footnotes: pkg.footnotes?.map(note => ({ ...note, content: [] })),
    endnotes: pkg.endnotes?.map(note => ({ ...note, content: [] })),
    relationships: pkg.relationships ? [...pkg.relationships] : undefined };
}

export function withPartMetadata(document: Document, metadata: ReturnType<typeof partMetadata>): Document {
  const value = metadata;
  return { ...document, package: { ...document.package, document: { ...document.package.document, ...value.body },
    headers: value.headers ? new Map(value.headers) : undefined, footers: value.footers ? new Map(value.footers) : undefined,
    footnotes: value.footnotes, endnotes: value.endnotes, relationships: value.relationships ? new Map(value.relationships) : undefined } };
}

export function validPartTarget(value: unknown): value is DocxPartTarget {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const p = value as Record<string, unknown>;
  return Object.keys(p).every(key => ['kind', 'section', 'page', 'variant', 'noteId', 'expectVersion'].includes(key)) &&
    ['header', 'footer', 'footnote', 'endnote'].includes(String(p.kind)) && typeof p.section === 'string' &&
    Number.isSafeInteger(p.page) && Number(p.page) >= 0 && typeof p.expectVersion === 'string' &&
    (p.variant === undefined || ['default', 'first', 'even'].includes(String(p.variant))) &&
    (p.noteId === undefined || Number.isSafeInteger(p.noteId) && Number(p.noteId) > 0);
}

function properties(document: Document, section: string): SectionProperties | null {
  return section === 'final' ? resolvedFinalSectionProperties(document.package.document) ?? {}
    : document.package.document.sections?.find(entry => entry.id === section || entry.properties.sectionId === section)?.properties ?? null;
}

function partVariant(engine: ResidentEngineSession, target: DocxPartTarget, properties: SectionProperties): HeaderFooterType | null {
  const pages = (JSON.parse(engine.retainedLayoutJson()) as { layout: { pages: { sectionIndex?: number; sectionPageIndex?: number; sectionPageNumber?: number; number: number }[] } }).layout.pages;
  const page = pages[target.page];
  if (!page) return null;
  const localPage = page.sectionPageIndex ?? pages.slice(0, target.page).filter(other => other.sectionIndex === page.sectionIndex).length;
  const number = page.sectionPageNumber ?? page.number;
  return target.variant ?? (properties.titlePg && localPage === 0 ? 'first'
    : (properties.evenAndOddHeaders || engine.partDocument().package.settings?.evenAndOddHeaders) && number % 2 === 0 ? 'even' : 'default');
}

export function resolvePart(engine: ResidentEngineSession, target: DocxPartTarget): ResidentPartTarget | null {
  const document = engine.partDocument();
  const props = properties(document, target.section);
  if (!props) return null;
  if (target.kind === 'footnote' || target.kind === 'endnote') {
    const story = `${target.kind === 'footnote' ? 'fn' : 'en'}:${target.noteId}`;
    return target.noteId !== undefined && engine.geometryReader.hasStory(story) ? { ...target, story } : null;
  }
  const variant = partVariant(engine, target, props);
  if (!variant) return null;
  const refs = target.kind === 'header' ? props.headerReferences : props.footerReferences;
  const ref = refs?.find(ref => ref.type === variant);
  if (!ref) return null;
  const story = engine.partStory(ref.rId);
  return engine.geometryReader.hasStory(story) ? { ...target, variant, story } : null;
}

export function ensurePartDocument(engine: ResidentEngineSession, target: DocxPartTarget): Document | null {
  const document = engine.partDocument();
  const props = properties(document, target.section);
  if (!props) return null;
  if (target.kind === 'footnote' || target.kind === 'endnote') {
    if (target.noteId === undefined) return null;
    const key = target.kind === 'footnote' ? 'footnotes' : 'endnotes';
    const notes = document.package[key] ?? [];
    const relationships = new Map(document.package.relationships);
    const type = `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${key}`;
    if (![...relationships.values()].some(relationship => relationship.type === type)) {
      const id = `rIdPart${crypto.randomUUID().replaceAll('-', '')}`;
      relationships.set(id, { id, type, target: `${key}.xml` });
    }
    return { ...document, package: { ...document.package, relationships,
      [key]: [...notes.filter(note => note.id !== target.noteId), { type: target.kind, id: target.noteId, noteType: 'normal', content: [] }] } };
  }
  const variant = partVariant(engine, target, props);
  if (!variant) return null;
  const id = `rIdPart${crypto.randomUUID().replaceAll('-', '')}`;
  const name = target.kind === 'header' ? 'headers' : 'footers';
  const refKey = target.kind === 'header' ? 'headerReferences' : 'footerReferences';
  const parts = new Map(document.package[name]);
  parts.set(id, { type: target.kind, hdrFtrType: variant, content: [] });
  const refs = [...props[refKey] ?? [], { type: variant, rId: id }];
  const body = target.section === 'final' ? updateFinalSectionProperties(document.package.document, properties => ({ ...properties, [refKey]: refs }))
    : { ...document.package.document, sections: document.package.document.sections?.map(section => section.id === target.section || section.properties.sectionId === target.section
      ? { ...section, properties: { ...section.properties, [refKey]: refs } } : section) };
  const relationships = new Map(document.package.relationships);
  relationships.set(id, { id, type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${target.kind}`, target: `${target.kind}-${id}.xml` });
  return { ...document, package: { ...document.package, [name]: parts, relationships, document: body } };
}

export function removePartDocument(engine: ResidentEngineSession, target: DocxPartTarget): Document | null {
  const part = resolvePart(engine, target);
  if (!part) return null;
  const document = engine.partDocument();
  if (target.kind === 'footnote' || target.kind === 'endnote') {
    const key = target.kind === 'footnote' ? 'footnotes' : 'endnotes';
    return { ...document, package: { ...document.package, [key]: document.package[key]?.filter(note => note.id !== target.noteId) } };
  }
  const refKey = target.kind === 'header' ? 'headerReferences' : 'footerReferences';
  const strip = (props: SectionProperties) => ({ ...props, [refKey]: props[refKey]?.filter(ref => engine.partStory(ref.rId) !== part.story) });
  const body = { ...document.package.document, sections: document.package.document.sections?.map(section => ({ ...section, properties: strip(section.properties) })),
    finalSectionProperties: strip(document.package.document.finalSectionProperties ?? {}) };
  const name = target.kind === 'header' ? 'headers' : 'footers';
  const parts = new Map(document.package[name]);
  const relationships = new Map(document.package.relationships);
  for (const id of parts.keys()) if (engine.partStory(id) === part.story) { parts.delete(id); relationships.delete(id); }
  return { ...document, package: { ...document.package, [name]: parts, relationships, document: body } };
}

export function readPartGeometry(engine: ResidentEngineSession, frame: RetainedFrame): ResidentPartGeometry[] {
  const result: ResidentPartGeometry[] = [];
  for (const page of frame.pages) {
    if (page.page.unbuilt) continue;
    for (const part of [page.page.header, page.page.footer]) if (part) result.push({ kind: part.kind, story: engine.partStory(part.rId),
      pageIndex: page.pageIndex, pageId: page.pageId.toString(), x: 0, y: part.y, width: page.page.width, height: part.height });
    for (const area of page.page.noteAreas ?? []) for (const id of area.noteIds ?? area.notes?.flatMap(note => note.id === undefined ? [] : [note.id]) ?? []) result.push({ kind: area.kind ?? 'footnote',
      story: `${area.kind === 'endnote' ? 'en' : 'fn'}:${id}`, pageIndex: page.pageIndex, pageId: page.pageId.toString(),
      x: 0, y: area.y ?? 0, width: page.page.width, height: area.height ?? 0 });
  }
  return result;
}
