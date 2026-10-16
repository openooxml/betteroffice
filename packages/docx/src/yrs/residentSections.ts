import type { Document, SectionProperties } from '../types/document';
import type { ResidentEngineSession } from './residentEngineSession';
import type { YrsLoc, YrsParagraphAttrs, YrsParagraphTabStop } from './index';
import { resolvedFinalSectionProperties } from '../editor/finalSection';

export type DocxSectionMargins = Partial<Pick<SectionProperties,
  'marginLeft' | 'marginRight' | 'marginTop' | 'marginBottom' | 'headerDistance' | 'footerDistance' | 'gutter'>>;

export interface ResidentSectionSnapshot {
  version: string;
  sections: { section: string; properties: SectionProperties }[];
  paragraphs: { story: string; paraId: string; tabs: NonNullable<YrsParagraphAttrs['tabs']> }[];
}

const tabAlignments: readonly YrsParagraphTabStop['alignment'][] = ['left', 'center', 'right', 'decimal', 'bar', 'clear', 'num'];
const tabLeaders: readonly NonNullable<YrsParagraphTabStop['leader']>[] = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];

export function nativeTabStops(value: unknown): NonNullable<YrsParagraphAttrs['tabs']> {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) throw new Error('Invalid native tab stops');
  return value.map(tab => {
    if (!tab || typeof tab !== 'object' || Array.isArray(tab)) throw new Error('Invalid native tab stop');
    const position = tab.position ?? tab.pos;
    const rawAlignment = tab.alignment ?? tab.val;
    const alignment = tabAlignments.find(value => value === (rawAlignment === 'start' ? 'left' : rawAlignment === 'end' ? 'right' : rawAlignment));
    const leader = tabLeaders.find(value => value === tab.leader);
    if (typeof position !== 'number' || !Number.isFinite(position) || alignment === undefined ||
      (tab.leader != null && leader === undefined)) throw new Error('Invalid native tab stop');
    return { position, alignment, ...(leader === undefined ? {} : { leader }) };
  });
}

export function validSectionMargins(value: unknown): value is DocxSectionMargins {
  return !!value && typeof value === 'object' && !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null) &&
    Object.entries(value).every(([key, entry]) => ['marginLeft', 'marginRight', 'marginTop', 'marginBottom', 'headerDistance', 'footerDistance', 'gutter'].includes(key) &&
      typeof entry === 'number' && Number.isFinite(entry) && entry >= 0);
}

export function residentSectionPropertiesAt(engine: ResidentEngineSession, at: YrsLoc): SectionProperties | null {
  if (at.story !== 'body') return null;
  const snapshot = engine.sectionProperties();
  const paragraphs = snapshot.paragraphs.filter(paragraph => paragraph.story === 'body');
  const ordinal = paragraphs.findIndex(paragraph => paragraph.paraId === at.paraId);
  if (ordinal < 0) return null;
  return snapshot.sections.find(section => section.section === 'final' ||
    paragraphs.findIndex(paragraph => paragraph.paraId === section.section) >= ordinal)?.properties ?? null;
}

export function sectionMarkKey(mark: unknown): string {
  const sorted = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(sorted);
    if (value !== null && typeof value === 'object') {
      return Object.fromEntries(Object.keys(value).sort().map(key =>
        [key, sorted((value as Record<string, unknown>)[key])]));
    }
    return value;
  };
  return JSON.stringify(sorted(mark));
}

export function readSectionSnapshot(engine: ResidentEngineSession, base: Document, margins: Record<string, DocxSectionMargins>, seeded: ReadonlyMap<string, string>, nativeSections = false): ResidentSectionSnapshot {
  const sections: ResidentSectionSnapshot['sections'] = [];
  const paragraphs: ResidentSectionSnapshot['paragraphs'] = [];
  for (const story of engine.geometryReader.storyIds()) {
    for (const paragraph of engine.paragraphs(story)) {
      paragraphs.push({ story, paraId: paragraph.paraId, tabs: nativeTabStops(paragraph.properties.tabs) });
      const properties = (paragraph.properties.sectPr ?? paragraph.properties._sectionProperties) as SectionProperties | undefined;
      if (story === 'body' && properties) {
        const source = !nativeSections && seeded.get(paragraph.paraId) === sectionMarkKey(properties)
          ? base.package.document.sections?.find(section => section.id === paragraph.paraId)?.properties
          : undefined;
        sections.push({ section: paragraph.paraId,
          properties: { ...properties, ...source, ...margins[paragraph.paraId] } });
      }
    }
  }
  sections.push({ section: 'final', properties: { ...resolvedFinalSectionProperties(base.package.document), ...margins.final } });
  return { version: engine.proposalEngine.version(), sections, paragraphs };
}

/** `document` with each section keyed by the body paragraph that closes it in the freshly seeded session. */
export function withSectionIds(document: Document, closing: readonly string[]): Document {
  const body = document.package.document;
  if (!body.sections || closing.length === 0) return document;
  return { ...document, package: { ...document.package, document: { ...body,
    sections: body.sections.map((section, index) => index < closing.length ? { ...section, id: closing[index] } : section),
  } } };
}

export function withSectionNames(document: Document, names: ReadonlyMap<string, string>): Document {
  const body = document.package.document;
  if (!body.sections || names.size === 0) return document;
  return { ...document, package: { ...document.package, document: { ...body,
    sections: body.sections.map(section => section.id !== undefined && names.has(section.id) ? { ...section, id: names.get(section.id) } : section),
  } } };
}

export function withSectionSnapshot(document: Document, snapshot: ResidentSectionSnapshot): Document {
  const body = document.package.document;
  const finalSectionProperties = snapshot.sections.at(-1)!.properties;
  const live = new Map(snapshot.sections.map(({ section, properties }) => [section, properties]));
  const sections = body.sections?.map((section, index, all) => ({ ...section,
    properties: (section.id === undefined ? undefined : live.get(section.id))
      ?? (index === all.length - 1 ? finalSectionProperties : section.properties),
  }));
  return { ...document, package: { ...document.package, document: { ...body,
    ...(sections ? { sections } : {}), finalSectionProperties,
  } } };
}
