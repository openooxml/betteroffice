/** Public note compatibility helpers backed by the Rust story parser. */

import type {
  Endnote,
  EndnotePosition,
  EndnoteProperties,
  Footnote,
  FootnotePosition,
  FootnoteProperties,
  MediaFile,
  NoteNumberRestart,
  NumberFormat,
  RelationshipMap,
  Theme,
} from '../types/document';
import type { ChartPartsMap } from './chartParser';
import type { NumberingMap } from './numberingParser';
import { parseDocumentWithRust } from './rustParseFacade';
import type { SmartArtContext } from './smartArtParser';
import type { StyleMap } from './styleParser';
import { rezipContainer } from './wasm';
import {
  findChild,
  findChildren,
  getAttribute,
  parseNumericAttribute,
  type XmlElement,
} from './xmlParser';

export interface FootnoteMap {
  byId: Map<number, Footnote>;
  footnotes: Footnote[];
  getFootnote(id: number): Footnote | undefined;
  hasFootnote(id: number): boolean;
  getNormalFootnotes(): Footnote[];
  getSeparator(): Footnote | undefined;
  getContinuationSeparator(): Footnote | undefined;
}

export interface EndnoteMap {
  byId: Map<number, Endnote>;
  endnotes: Endnote[];
  getEndnote(id: number): Endnote | undefined;
  hasEndnote(id: number): boolean;
  getNormalEndnotes(): Endnote[];
  getSeparator(): Endnote | undefined;
  getContinuationSeparator(): Endnote | undefined;
}

export function parseFootnotes(
  footnotesXml: string | null,
  _styles: StyleMap | null = null,
  _theme: Theme | null = null,
  _numbering: NumberingMap | null = null,
  _rels: RelationshipMap | null = null,
  _media: Map<string, MediaFile> | null = null,
  _charts?: ChartPartsMap | null,
  _smartArt: SmartArtContext | null = null
): FootnoteMap {
  if (!footnotesXml) return createFootnoteMap([], new Map());
  const pkg = parseNotePart('footnote', footnotesXml);
  const footnotes = [...(pkg.footnoteSeparators ?? []), ...(pkg.footnotes ?? [])];
  return createFootnoteMap(footnotes, new Map(footnotes.map((note) => [note.id, note])));
}

export function parseEndnotes(
  endnotesXml: string | null,
  _styles: StyleMap | null = null,
  _theme: Theme | null = null,
  _numbering: NumberingMap | null = null,
  _rels: RelationshipMap | null = null,
  _media: Map<string, MediaFile> | null = null,
  _charts?: ChartPartsMap | null,
  _smartArt: SmartArtContext | null = null
): EndnoteMap {
  if (!endnotesXml) return createEndnoteMap([], new Map());
  const pkg = parseNotePart('endnote', endnotesXml);
  const endnotes = [...(pkg.endnoteSeparators ?? []), ...(pkg.endnotes ?? [])];
  return createEndnoteMap(endnotes, new Map(endnotes.map((note) => [note.id, note])));
}

function parseNotePart(kind: 'footnote' | 'endnote', xml: string) {
  const plural = `${kind}s`;
  const encoder = new TextEncoder();
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        `<Override PartName="/word/${plural}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${plural}+xml"/>` +
        '</Types>'
    ),
    '_rels/.rels': encoder.encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>'
    ),
    'word/document.xml': encoder.encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'
    ),
    'word/_rels/document.xml.rels': encoder.encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${plural}" Target="${plural}.xml"/>` +
        '</Relationships>'
    ),
    [`word/${plural}.xml`]: encoder.encode(xml),
  };
  const zipped = rezipContainer(entries);
  const buffer = new ArrayBuffer(zipped.byteLength);
  new Uint8Array(buffer).set(zipped);
  return parseDocumentWithRust(buffer, {
    parseHeadersFooters: false,
    parseNotes: true,
    detectVariables: false,
  }).document.package;
}

function parseNumberFormat(value: string | null): NumberFormat | undefined {
  const formats: Record<string, NumberFormat> = {
    decimal: 'decimal',
    upperRoman: 'upperRoman',
    lowerRoman: 'lowerRoman',
    upperLetter: 'upperLetter',
    lowerLetter: 'lowerLetter',
    ordinal: 'ordinal',
    cardinalText: 'cardinalText',
    ordinalText: 'ordinalText',
    bullet: 'bullet',
    chicago: 'chicago',
    none: 'none',
  };
  return value ? formats[value] : undefined;
}

function parseFootnotePosition(value: string | null): FootnotePosition | undefined {
  return value === 'pageBottom' || value === 'beneathText' || value === 'sectEnd' || value === 'docEnd'
    ? value
    : undefined;
}

function parseEndnotePosition(value: string | null): EndnotePosition | undefined {
  return value === 'sectEnd' || value === 'docEnd' ? value : undefined;
}

function parseNumberRestart(value: string | null): NoteNumberRestart | undefined {
  return value === 'continuous' || value === 'eachSect' || value === 'eachPage'
    ? value
    : undefined;
}

export function parseFootnoteProperties(element: XmlElement | null): FootnoteProperties {
  const props: FootnoteProperties = {};
  if (!element) return props;
  props.position = parseFootnotePosition(getAttribute(findChild(element, 'w', 'pos'), 'w', 'val'));
  const format = findChild(element, 'w', 'numFmt');
  props.numFmt = parseNumberFormat(getAttribute(format, 'w', 'val'));
  const custom = getAttribute(format, 'w', 'format');
  if (custom) props.customNumberFormat = custom.slice(0, 255);
  props.numStart = parseNumericAttribute(findChild(element, 'w', 'numStart'), 'w', 'val');
  props.numRestart = parseNumberRestart(
    getAttribute(findChild(element, 'w', 'numRestart'), 'w', 'val')
  );
  const separators = noteSeparators(element, 'footnote');
  if (separators.length) props.separators = separators;
  return removeUndefined(props);
}

export function parseEndnoteProperties(element: XmlElement | null): EndnoteProperties {
  const props: EndnoteProperties = {};
  if (!element) return props;
  props.position = parseEndnotePosition(getAttribute(findChild(element, 'w', 'pos'), 'w', 'val'));
  const format = findChild(element, 'w', 'numFmt');
  props.numFmt = parseNumberFormat(getAttribute(format, 'w', 'val'));
  const custom = getAttribute(format, 'w', 'format');
  if (custom) props.customNumberFormat = custom.slice(0, 255);
  props.numStart = parseNumericAttribute(findChild(element, 'w', 'numStart'), 'w', 'val');
  props.numRestart = parseNumberRestart(
    getAttribute(findChild(element, 'w', 'numRestart'), 'w', 'val')
  );
  const separators = noteSeparators(element, 'endnote');
  if (separators.length) props.separators = separators;
  return removeUndefined(props);
}

function noteSeparators(element: XmlElement, name: 'footnote' | 'endnote') {
  return findChildren(element, 'w', name)
    .slice(0, 32)
    .map((reference) => parseNumericAttribute(reference, 'w', 'id'))
    .filter((id): id is number => id !== undefined)
    .map((noteId) => ({
      noteId,
      kind:
        noteId === -1
          ? ('separator' as const)
          : noteId === 0
            ? ('continuationSeparator' as const)
            : ('continuationNotice' as const),
    }));
}

function removeUndefined<T extends object>(value: T): T {
  for (const key of Object.keys(value) as Array<keyof T>) {
    if (value[key] === undefined) delete value[key];
  }
  return value;
}

export function getFootnoteText(footnote: Footnote): string {
  return noteText(footnote);
}

export function getEndnoteText(endnote: Endnote): string {
  return noteText(endnote);
}

function noteText(note: Footnote | Endnote): string {
  return note.content
    .filter((block) => block.type === 'paragraph')
    .map((paragraph) =>
      paragraph.content
        .filter((item) => item.type === 'run')
        .flatMap((run) => run.content)
        .filter((content) => content.type === 'text')
        .map((content) => content.text)
        .join('')
    )
    .join('\n');
}

export function isSeparatorFootnote(note: Footnote): boolean {
  return note.noteType !== 'normal';
}

export function isSeparatorEndnote(note: Endnote): boolean {
  return note.noteType !== 'normal';
}

function createFootnoteMap(footnotes: Footnote[], byId: Map<number, Footnote>): FootnoteMap {
  return {
    byId,
    footnotes,
    getFootnote: (id) => byId.get(id),
    hasFootnote: (id) => byId.has(id),
    getNormalFootnotes: () => footnotes.filter((note) => !isSeparatorFootnote(note)),
    getSeparator: () => footnotes.find((note) => note.noteType === 'separator'),
    getContinuationSeparator: () =>
      footnotes.find((note) => note.noteType === 'continuationSeparator'),
  };
}

function createEndnoteMap(endnotes: Endnote[], byId: Map<number, Endnote>): EndnoteMap {
  return {
    byId,
    endnotes,
    getEndnote: (id) => byId.get(id),
    hasEndnote: (id) => byId.has(id),
    getNormalEndnotes: () => endnotes.filter((note) => !isSeparatorEndnote(note)),
    getSeparator: () => endnotes.find((note) => note.noteType === 'separator'),
    getContinuationSeparator: () =>
      endnotes.find((note) => note.noteType === 'continuationSeparator'),
  };
}

export function createEmptyFootnoteMap(): FootnoteMap {
  return createFootnoteMap([], new Map());
}

export function createEmptyEndnoteMap(): EndnoteMap {
  return createEndnoteMap([], new Map());
}

export function mergeFootnoteMaps(...maps: FootnoteMap[]): FootnoteMap {
  const byId = new Map<number, Footnote>();
  for (const map of maps) for (const note of map.footnotes) if (!byId.has(note.id)) byId.set(note.id, note);
  return createFootnoteMap([...byId.values()], byId);
}

export function mergeEndnoteMaps(...maps: EndnoteMap[]): EndnoteMap {
  const byId = new Map<number, Endnote>();
  for (const map of maps) for (const note of map.endnotes) if (!byId.has(note.id)) byId.set(note.id, note);
  return createEndnoteMap([...byId.values()], byId);
}
