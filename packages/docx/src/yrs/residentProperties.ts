import type { Document, SectionProperties, Watermark, HeaderFooter, Relationship } from '../types/document';
import type { YrsHyperlinkAttrs, YrsLoc, YrsStickyPosition, YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandSelection } from './residentCommandProtocol';
import { commandSelectionRange } from './residentCommandAnchors';
import { getDocumentWatermark, setDocumentWatermark } from '../docx/watermarkApi';
import { updateFinalSectionProperties, resolvedFinalSectionProperties } from '../editor/finalSection';

export interface ResidentPackageMetadata {
  body: Pick<Document['package']['document'], 'sections' | 'finalSectionProperties'>;
  headers?: [string, HeaderFooter][];
  footers?: [string, HeaderFooter][];
  relationships?: [string, Relationship][];
}

export type ResidentPropertyTarget =
  | { kind: 'section'; section: string }
  | { kind: 'watermark' }
  | { kind: 'hyperlink'; selection: DocxCommandSelection }
  | { kind: 'objectSelection'; selection: DocxCommandSelection }
  | { kind: 'object'; at: YrsLoc };

export type ResidentPropertySnapshot =
  | { kind: 'section'; section: string; properties: SectionProperties }
  | { kind: 'watermark'; watermark: Watermark | null }
  | { kind: 'hyperlink'; range: YrsStoryRange; hyperlink: YrsHyperlinkAttrs | null }
  | { kind: 'object'; handle: string; properties: Record<string, unknown> };

export function packageMetadata(document: Document): ResidentPackageMetadata {
  const pkg = document.package;
  return {
    body: {
      sections: pkg.document.sections?.map(section => ({ ...section, content: [] })),
      finalSectionProperties: pkg.document.finalSectionProperties,
    },
    headers: pkg.headers ? [...pkg.headers].map(([id, value]) => [id, { ...value, content: [] }]) : undefined,
    footers: pkg.footers ? [...pkg.footers].map(([id, value]) => [id, { ...value, content: [] }]) : undefined,
    relationships: pkg.relationships ? [...pkg.relationships] : undefined,
  };
}

export function withPackageMetadata(document: Document, metadata: ResidentPackageMetadata): Document {
  const headers = metadata.headers?.map(([id, value]): [string, HeaderFooter] => {
    const watermark = value.watermark;
    if (watermark?.kind !== 'picture' || !watermark.data || watermark.data instanceof Uint8Array) return [id, value];
    return [id, { ...value, watermark: { ...watermark, data: Uint8Array.from(Object.values(watermark.data)) } }];
  });
  return { ...document, package: { ...document.package,
    document: { ...document.package.document, ...metadata.body },
    headers: headers ? new Map(headers) : undefined,
    footers: metadata.footers ? new Map(metadata.footers) : undefined,
    relationships: metadata.relationships ? new Map(metadata.relationships) : undefined,
  } };
}

export function sectionProperties(document: Document, section: string): SectionProperties | null {
  if (section === 'final') return resolvedFinalSectionProperties(document.package.document) ?? {};
  const found = document.package.document.sections?.[sectionIndex(document, section)];
  return found?.properties ?? null;
}

function sectionIndex(document: Document, section: string): number {
  const sections = document.package.document.sections ?? [];
  const found = sections.findIndex(entry => entry.id === section || entry.properties.sectionId === section);
  if (found >= 0) return found;
  const index = /^(0|[1-9]\d*)$/.test(section) ? Number(section) : -1;
  return index >= 0 && index < sections.length - 1 ? index : -1;
}

export function patchSection(document: Document, section: string, patch: Partial<SectionProperties>): Document | null {
  if (!sectionProperties(document, section)) return null;
  const body = document.package.document;
  const targetIndex = sectionIndex(document, section);
  const next = section === 'final' ? updateFinalSectionProperties(body, properties => ({ ...properties, ...patch }))
    : { ...body, sections: body.sections?.map((entry, index) => index === targetIndex
      ? { ...entry, properties: { ...entry.properties, ...patch } } : entry) };
  return { ...document, package: { ...document.package, document: next } };
}

interface ImageHandle { position: YrsStickyPosition }
const images = new WeakMap<ResidentEngineSession, Map<string, ImageHandle>>();

function imageAt(engine: ResidentEngineSession, at: YrsLoc): Record<string, unknown> | null {
  if (!engine.geometryReader.hasStory(at.story)) return null;
  const p = engine.geometryReader.paragraphSpans(at.story).findIndex(entry => entry.paraId === at.paraId);
  if (p < 0) return null;
  let offset = 0;
  const segment = engine.paragraphSegments(at.story, [p])[0]?.find(entry => {
    const start = offset;
    offset += entry.kind === 'text' ? entry.text.length : 1;
    return start === at.offset;
  });
  return segment?.kind === 'embed' && segment.embedKind === 'image' ? segment.payload : null;
}

export function resolvePropertyImage(engine: ResidentEngineSession, handle: string): YrsLoc | null {
  const entry = images.get(engine)?.get(handle);
  if (!entry) return null;
  const at = engine.resolveStickyPosition(entry.position);
  if (!at || !imageAt(engine, at)) return null;
  const current = engine.encodeStickyPosition(at);
  return String(current.encoded) === String(entry.position.encoded) ? at : null;
}

export function readResidentProperties(engine: ResidentEngineSession, target: ResidentPropertyTarget): ResidentPropertySnapshot | null {
  if (target.kind === 'section') {
    const properties = sectionProperties(engine.propertyDocument(), target.section);
    return properties ? { kind: 'section', section: target.section, properties } : null;
  }
  if (target.kind === 'watermark') return { kind: 'watermark', watermark: getDocumentWatermark(engine.propertyDocument()) ?? null };
  if (target.kind === 'hyperlink') {
    const range = commandSelectionRange(engine, target.selection);
    if (!range) return null;
    const focus = target.selection.focus;
    const span = engine.geometryReader.locateParagraph(range.story, focus.paraId);
    const position = span.start + focus.offset;
    let offset = 0;
    const segments = engine.storySegments(range.story).map(entry => {
      const start = offset;
      offset += entry.kind === 'text' ? entry.text.length : 1;
      const hyperlink = entry.attributes.hyperlink as YrsHyperlinkAttrs | undefined;
      return { start, end: offset, hyperlink: hyperlink?.href ? hyperlink : null };
    });
    let index = segments.findIndex(segment => segment.hyperlink && segment.start < position && position < segment.end);
    if (index < 0) index = segments.findIndex(segment => segment.hyperlink && (position === segment.start || position === segment.end));
    const hyperlink = segments[index]?.hyperlink;
    if (!hyperlink) return { kind: 'hyperlink', range, hyperlink: null };
    let first = index;
    let last = index;
    while (first > 0 && segments[first - 1]!.end === segments[first]!.start && segments[first - 1]!.hyperlink?.href === hyperlink.href) first -= 1;
    while (last + 1 < segments.length && segments[last]!.end === segments[last + 1]!.start && segments[last + 1]!.hyperlink?.href === hyperlink.href) last += 1;
    const locate = (position: number) => {
      for (const paragraph of engine.geometryReader.paragraphSpans(range.story)) {
        const span = engine.geometryReader.locateParagraph(range.story, paragraph.paraId);
        if (span.start <= position && position <= span.end) return { paraId: paragraph.paraId, offset: position - span.start };
      }
      return null;
    };
    const start = locate(segments[first]!.start);
    const end = locate(segments[last]!.end);
    return start && end ? { kind: 'hyperlink', range: { story: range.story, start, end }, hyperlink } : null;
  }
  const range = target.kind === 'objectSelection' ? commandSelectionRange(engine, target.selection) : null;
  if (target.kind === 'objectSelection' && (!range || range.start.paraId !== range.end.paraId || range.end.offset - range.start.offset !== 1)) return null;
  const at = target.kind === 'objectSelection' ? { story: range!.story, ...range!.start } : target.at;
  const properties = imageAt(engine, at);
  if (!properties) return null;
  let handles = images.get(engine);
  if (!handles) { handles = new Map(); images.set(engine, handles); }
  const position = engine.encodeStickyPosition(at);
  const existing = [...handles].find(([, entry]) => entry.position.story === position.story && String(entry.position.encoded) === String(position.encoded));
  const handle = existing?.[0] ?? crypto.randomUUID();
  if (!existing) handles.set(handle, { position });
  return { kind: 'object', handle, properties };
}

export function watermarkDocument(document: Document, watermark: Watermark | null): Document {
  return setDocumentWatermark(document, watermark);
}
