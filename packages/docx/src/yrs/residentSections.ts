import type { Document, SectionProperties } from '../types/document';
import type { ResidentEngineSession } from './residentEngineSession';
import type { YrsParagraphAttrs, YrsParagraphTabStop } from './index';
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

export function readSectionSnapshot(engine: ResidentEngineSession, base: Document, margins: Record<string, DocxSectionMargins>): ResidentSectionSnapshot {
  const sections: ResidentSectionSnapshot['sections'] = [];
  const paragraphs: ResidentSectionSnapshot['paragraphs'] = [];
  for (const story of engine.geometryReader.storyIds()) {
    for (const paragraph of engine.paragraphs(story)) {
      paragraphs.push({ story, paraId: paragraph.paraId, tabs: nativeTabStops(paragraph.properties.tabs) });
      const properties = (paragraph.properties.sectPr ?? paragraph.properties._sectionProperties) as SectionProperties | undefined;
      if (story === 'body' && properties) sections.push({ section: paragraph.paraId,
        properties: { ...properties, ...base.package.document.sections?.[sections.length]?.properties, ...margins[paragraph.paraId] } });
    }
  }
  sections.push({ section: 'final', properties: { ...resolvedFinalSectionProperties(base.package.document), ...margins.final } });
  return { version: engine.proposalEngine.version(), sections, paragraphs };
}

export function withSectionSnapshot(document: Document, snapshot: ResidentSectionSnapshot): Document {
  const body = document.package.document;
  const finalSectionProperties = snapshot.sections.at(-1)!.properties;
  const sections = body.sections?.map((section, index) => ({ ...section,
    properties: snapshot.sections[index]?.properties ?? finalSectionProperties,
  }));
  return { ...document, package: { ...document.package, document: { ...body,
    ...(sections ? { sections } : {}), finalSectionProperties,
  } } };
}
