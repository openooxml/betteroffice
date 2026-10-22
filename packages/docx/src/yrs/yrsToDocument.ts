/** yrs-to-OOXML save projection. */

/* eslint-disable max-lines -- the inverse mapping stays co-located with its save orchestrator */

import { engineProgressTick } from '../utils/engineProgress';
import { headerFooterAliasOf, markHeaderFooterAlias } from '../docx/headerFooterAliasProjection';
import { markRevisionInfo } from '../docx/sessionRevisionIds';
import { headerFooterStory } from './sessionInternals';
import { createStyleResolver, type StyleResolver } from '../styles';
import { hasTrackedControlContent } from '../utils/trackedControlContent';
import { isRawXml } from '../types/content/rawXml';
import { pixelsToEmu } from '../utils/units';
import {
  applyContentControlValue,
  type ContentControlValue,
} from './contentControlValues';
import { sdtAttrsToProps } from '../types/sdtAttributes';
import {
  paragraphAttrsToFormatting,
  tableAttrsToFormatting,
  tableCellAttrsToFormatting,
  tableRowAttrsToFormatting,
  type ParagraphSaveAttrs,
  type TableCellSaveAttrs,
  type TableRowSaveAttrs,
  type TableSaveAttrs,
} from './saveFormatting';
import type {
  Document,
  BlockContent,
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  Run,
  RunContent,
  HorizontalRuleContent,
  TextFormatting,
  Hyperlink,
  HyperlinkContent,
  TrackedChangeInfo,
  Table,
  TableRow,
  TableCell,
  TableBorders,
  SimpleField,
  ComplexField,
  FieldType,
  FieldInlineContent,
  MathEquation,
  Image,
  ColorValue,
  Shape,
  Chart,
  InlineSdt,
  SdtProperties,
  Comment,
  Footnote,
  Endnote,
  HeaderFooter,
} from '../types/document';
import type { YrsRunInheritance, YrsSession, YrsTableInheritance } from './index';

type Attrs = Record<string, unknown>;

interface YrsImageAttrs {
  src?: string;
  alt?: string;
  title?: string;
  width?: number;
  height?: number;
  rId?: string;
  wrapType?: Image['wrap']['type'];
  transform?: string;
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
  position?: {
    relativeHeight?: number;
    horizontal?: { relativeTo?: string; posOffset?: number; align?: string };
    vertical?: { relativeTo?: string; posOffset?: number; align?: string };
  };
  borderWidth?: number;
  borderColor?: string;
  borderColorValue?: ColorValue;
  borderStyle?: string;
  wrapText?: string;
  hlinkHref?: string;
  cropTop?: number;
  cropRight?: number;
  cropBottom?: number;
  cropLeft?: number;
  shapeType?: string;
  opacity?: number;
  layoutInCell?: boolean;
  allowOverlap?: boolean;
  effectExtentTop?: number;
  effectExtentRight?: number;
  effectExtentBottom?: number;
  effectExtentLeft?: number;
}

interface TextItem {
  kind: 'text';
  text: string;
  attributes: Attrs;
}

interface EmbedItem {
  kind: 'embed';
  embedKind: string;
  payload: Attrs;
  attributes: Attrs;
}

type InlineItem = TextItem | EmbedItem;

const PARAGRAPH_ATTR_DEFAULTS: Attrs = {
  paraId: null,
  textId: null,
  alignment: null,
  spaceBefore: null,
  spaceAfter: null,
  lineSpacing: null,
  lineSpacingRule: null,
  spacingExplicit: null,
  indentLeft: null,
  indentRight: null,
  indentFirstLine: null,
  hangingIndent: false,
  numPr: null,
  numPrFromStyle: null,
  listNumFmt: null,
  listIsBullet: null,
  listMarker: null,
  listMarkerHidden: null,
  listMarkerFontFamily: null,
  listMarkerFontSize: null,
  listMarkerSuffix: null,
  listLevelNumFmts: null,
  listAbstractNumId: null,
  listStartOverride: null,
  styleId: null,
  borders: null,
  shading: null,
  tabs: null,
  pageBreakBefore: null,
  renderedPageBreakBefore: null,
  keepNext: null,
  keepLines: null,
  widowControl: null,
  contextualSpacing: null,
  snapToGrid: null,
  autoSpaceDE: null,
  autoSpaceDN: null,
  adjustRightInd: null,
  cnfStyle: null,
  defaultTextFormatting: null,
  sectionBreakType: null,
  bidi: null,
  outlineLevel: null,
  bookmarks: null,
  _originalFormatting: null,
  _originalRunBoundaries: null,
  _sectionProperties: null,
  pPrIns: null,
  pPrDel: null,
  pPrChange: null,
};

const TABLE_ATTR_DEFAULTS: Attrs = {
  styleId: null,
  width: null,
  widthType: null,
  justification: null,
  columnWidths: null,
  tableLayout: null,
  floating: null,
  cellMargins: null,
  look: null,
  bidi: null,
  _originalFormatting: null,
  tblPrChange: null,
};

const TABLE_ROW_ATTR_DEFAULTS: Attrs = {
  height: null,
  heightRule: null,
  isHeader: false,
  _originalFormatting: null,
  trIns: null,
  trDel: null,
  trPrChange: null,
};

const TABLE_CELL_ATTR_DEFAULTS: Attrs = {
  colspan: 1,
  rowspan: 1,
  colwidth: null,
  width: null,
  widthType: null,
  verticalAlign: null,
  backgroundColor: null,
  borders: null,
  margins: null,
  textDirection: null,
  noWrap: false,
  _originalFormatting: null,
  _originalResolvedFill: null,
  cellMarker: null,
  tcPrChange: null,
};

interface CommentBoundary {
  id: number;
  kind: 'start' | 'end';
  offset: number;
}

interface BookmarkBoundary extends CommentBoundary {
  name?: string;
  colFirst?: number;
  colLast?: number;
}

interface OriginalRunBoundary {
  text: string;
  noteMarks?: string[];
  /** Flow breaks the run held, at their offsets into `text`. */
  breaks?: Array<{ offset: number; type: 'page' | 'column' }>;
  marksKey?: string;
  formatting?: TextFormatting;
  propertyChanges?: Run['propertyChanges'];
}

interface LiveParagraphBreaks {
  leading: readonly BlockBreakKind[];
  trailing: readonly BlockBreakKind[];
  leadingPage: boolean;
}

interface TableCellPayload {
  tcPr?: Attrs;
  story?: string;
}

interface TableRowPayload {
  trPr?: Attrs;
  cells?: TableCellPayload[];
}

interface TablePayload extends Attrs {
  tblPr?: Attrs;
  grid?: unknown[];
  rows?: TableRowPayload[];
}

const BOOLEAN_MARKS = new Set([
  'bold',
  'italic',
  'superscript',
  'subscript',
  'allCaps',
  'smallCaps',
  'emboss',
  'imprint',
  'textShadow',
  'textOutline',
  'hidden',
  'rtl',
]);

function asObject(value: unknown): Attrs | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Attrs)
    : undefined;
}

function asFiniteNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined;
}

function dropNulls(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(dropNulls);
  const object = asObject(value);
  if (!object) return value;
  const result: Attrs = {};
  for (const [key, entry] of Object.entries(object)) {
    if (entry === null || entry === undefined) continue;
    result[key] = dropNulls(entry);
  }
  return result;
}

function stableStringify(value: unknown): string {
  if (value === null || value === undefined) return 'null';
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
  const object = asObject(value);
  if (object) {
    return `{${Object.keys(object)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(object[key])}`)
      .join(',')}}`;
  }
  return JSON.stringify(value);
}

function formattingAttrs(attributes: Attrs): Attrs {
  const result = { ...attributes };
  delete result.hyperlink;
  delete result.ins;
  delete result.del;
  return result;
}

/** Converts attributes without manufacturing mark objects. */
function attrsToTextFormatting(attributes: Attrs): TextFormatting {
  const formatting: TextFormatting = {};

  if (attributes.bold) formatting.bold = true;
  if (attributes.italic) formatting.italic = true;
  if (attributes.boldCs) formatting.boldCs = true;
  if (attributes.italicCs) formatting.italicCs = true;

  const underline = asObject(attributes.underline);
  if (underline && underline.inheritedHyperlink !== true) {
    formatting.underline = {
      style: (asString(underline.style) || 'single') as NonNullable<
        TextFormatting['underline']
      >['style'],
      color: (asObject(underline.color) ?? null) as TextFormatting['color'],
    };
  }

  const strike = asObject(attributes.strike);
  if (strike) {
    if (strike.double) formatting.doubleStrike = true;
    else formatting.strike = true;
  }

  const textColor = asObject(attributes.textColor);
  if (textColor && textColor.inheritedHyperlink !== true) {
    formatting.color = {
      rgb: (asString(textColor.rgb) ?? null) as string | undefined,
      themeColor: (textColor.themeColor ?? null) as NonNullable<
        TextFormatting['color']
      >['themeColor'],
      themeTint: (asString(textColor.themeTint) ?? null) as string | undefined,
      themeShade: (asString(textColor.themeShade) ?? null) as string | undefined,
    };
  }

  const highlight = asString(attributes.highlight) ?? asString(asObject(attributes.highlight)?.color);
  if (highlight !== undefined) formatting.highlight = highlight as TextFormatting['highlight'];

  const fontSize = asObject(attributes.fontSize);
  if (fontSize) {
    const size = asFiniteNumber(fontSize.size);
    const sizeCs = asFiniteNumber(fontSize.sizeCs);
    if (size !== undefined) formatting.fontSize = size;
    if (sizeCs !== undefined) formatting.fontSizeCs = sizeCs;
  }

  const fontFamily = asObject(attributes.fontFamily);
  if (fontFamily) {
    const ascii = asString(fontFamily.ascii);
    formatting.fontFamily = {
      ascii: (ascii ?? null) as string | undefined,
      hAnsi: (asString(fontFamily.hAnsi) ?? null) as string | undefined,
      eastAsia: asString(fontFamily.eastAsia),
      cs: asString(fontFamily.cs),
      asciiTheme: (fontFamily.asciiTheme ?? null) as NonNullable<
        TextFormatting['fontFamily']
      >['asciiTheme'],
      hAnsiTheme: asString(fontFamily.hAnsiTheme),
      eastAsiaTheme: asString(fontFamily.eastAsiaTheme),
      csTheme: asString(fontFamily.csTheme),
    };
  }

  if (attributes.superscript) formatting.vertAlign = 'superscript';
  if (attributes.subscript) formatting.vertAlign = 'subscript';
  if (attributes.allCaps) formatting.allCaps = true;
  if (attributes.smallCaps) formatting.smallCaps = true;

  const spacing = asObject(attributes.characterSpacing);
  if (spacing) {
    const charSpacing = asFiniteNumber(spacing.spacing);
    const position = asFiniteNumber(spacing.position);
    const scale = asFiniteNumber(spacing.scale);
    const kerning = asFiniteNumber(spacing.kerning);
    if (charSpacing !== undefined) formatting.spacing = charSpacing;
    if (position !== undefined) formatting.position = position;
    if (scale !== undefined) formatting.scale = scale;
    if (kerning !== undefined) formatting.kerning = kerning;
  }

  if (attributes.emboss) formatting.emboss = true;
  if (attributes.imprint) formatting.imprint = true;
  if (attributes.textShadow) formatting.shadow = true;

  const emphasis = asObject(attributes.emphasisMark);
  if (emphasis) {
    formatting.emphasisMark = (asString(emphasis.type) || 'dot') as NonNullable<
      TextFormatting['emphasisMark']
    >;
  }
  if (attributes.textOutline) formatting.outline = true;
  if (attributes.hidden) formatting.hidden = true;
  if (attributes.rtl) formatting.rtl = true;

  const effect = asObject(attributes.textEffect);
  if (effect) {
    formatting.effect = (asString(effect.effect) || 'blinkBackground') as NonNullable<
      TextFormatting['effect']
    >;
  }

  const modern = asObject(attributes.modernTextEffects);
  if (modern?.effects) formatting.modernEffects = modern.effects as TextFormatting['modernEffects'];

  const runStyle = asObject(attributes.runStyle);
  const styleId = asString(runStyle?.styleId);
  if (styleId) formatting.styleId = styleId;

  return formatting;
}

/** The run properties a run with these attributes saves with. */
type RunFormatter = (attributes: Attrs) => TextFormatting;

const directFormatting: RunFormatter = (attributes) =>
  attrsToTextFormatting(formattingAttrs(attributes));

const ON_OFF_RUN_PROPERTIES = new Set<string>([
  'bold', 'italic', 'allCaps', 'smallCaps', 'strike', 'doubleStrike',
  'outline', 'shadow', 'emboss', 'imprint', 'hidden', 'rtl', 'boldCs', 'italicCs',
]);

const SOURCE_RUN_PROPERTIES = ['bold', 'italic', 'fontSize', 'boldCs', 'italicCs', 'fontSizeCs', 'language'] as const;

const lentByLinkStyle = (mark: unknown): boolean => asObject(mark)?.inheritedHyperlink === true;

/** Writes differences from inherited run formatting, including explicit resets. */
function relativeFormatting(attributes: Attrs, inherited: TextFormatting): TextFormatting {
  const formatting = directFormatting(attributes);
  const result: Attrs = {};
  for (const key of new Set([...Object.keys(formatting), ...Object.keys(inherited)])) {
    if ((key === 'boldCs' || key === 'italicCs') && typeof attributes[key] !== 'boolean') continue;
    const value = formatting[key as keyof TextFormatting];
    if (sameJson(value, inherited[key as keyof TextFormatting])) continue;
    if (value !== undefined) result[key] = value;
    else if (ON_OFF_RUN_PROPERTIES.has(key)) result[key] = false;
    else if (key === 'color' && !lentByLinkStyle(attributes.textColor)) result[key] = { auto: true };
    else if (key === 'underline' && !lentByLinkStyle(attributes.underline)) {
      result[key] = { style: 'none' };
    }
  }
  const source = asObject(attributes._sourceRunFormatting);
  for (const key of SOURCE_RUN_PROPERTIES) {
    const value = source?.[key];
    if (value === undefined) continue;
    const live = ON_OFF_RUN_PROPERTIES.has(key) ? formatting[key] ?? false : formatting[key];
    // Runs hold no live language, so text keeps the language of the source run it maps to.
    if (key === 'language' || sameJson(value, live) ||
      ((key === 'boldCs' || key === 'italicCs') && attributes[key] === undefined)) {
      result[key] = value;
    }
  }
  return result as TextFormatting;
}

function runContentForText(text: string, formatting: TextFormatting): RunContent[] {
  const content: RunContent[] = [];
  let plainText = '';
  const flushText = (): void => {
    if (!plainText) return;
    content.push({ type: 'text', text: plainText });
    plainText = '';
  };

  const fonts = formatting.fontFamily
    ? [
        formatting.fontFamily.ascii,
        formatting.fontFamily.hAnsi,
        formatting.fontFamily.eastAsia,
        formatting.fontFamily.cs,
      ].filter((font): font is string => Boolean(font))
    : [];
  const symbolFont = fonts.length > 0 && fonts.every((font) => font === fonts[0]) ? fonts[0] : null;

  for (const char of text) {
    const codePoint = char.codePointAt(0) ?? 0;
    if (char === '\u00ad') {
      flushText();
      content.push({ type: 'softHyphen' });
    } else if (char === '\u2011') {
      flushText();
      content.push({ type: 'noBreakHyphen' });
    } else if (symbolFont && codePoint >= 0xf000 && codePoint <= 0xf8ff) {
      flushText();
      content.push({
        type: 'symbol',
        font: symbolFont,
        char: codePoint.toString(16).toUpperCase().padStart(4, '0'),
      });
    } else {
      plainText += char;
    }
  }
  flushText();
  return content;
}

function createTextRun(text: string, attributes: Attrs, format: RunFormatter): Run {
  const formatting = format(attributes);
  return {
    type: 'run',
    formatting: Object.keys(formatting).length > 0 ? formatting : undefined,
    content: runContentForText(text, directFormatting(attributes)),
  };
}

function appendTextRun(target: Run, source: Run): void {
  for (const item of source.content) {
    const previous = target.content[target.content.length - 1];
    if (previous?.type === 'text' && item.type === 'text') previous.text += item.text;
    else target.content.push(item);
  }
}

function fnv53(value: string): number {
  let hash = 0xcbf29ce484222325n;
  for (const byte of new TextEncoder().encode(value)) {
    hash = (hash ^ BigInt(byte)) * 0x100000001b3n;
    hash &= 0xffffffffffffffffn;
  }
  return Number(hash & ((1n << 53n) - 1n));
}

function revisionId(value: unknown): number {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string') {
    const parsed = Number(value);
    if (Number.isFinite(parsed)) return parsed;
    return fnv53(value);
  }
  return 0;
}

function trackedInfo(raw: unknown, _pmShape = false, id?: number): TrackedChangeInfo | null {
  const value = asObject(raw);
  if (!value) return null;
  const author = asString(value.author) || 'Unknown';
  const date = asString(value.date);
  const source = value.revisionId ?? value.id;
  const info = { id: id ?? revisionId(source), author, ...(date ? { date } : {}) };
  return markRevisionInfo(
    info,
    id === undefined && typeof source === 'string' && !Number.isFinite(Number(source))
  );
}

function propertyChanges(changes: unknown[]): unknown[] {
  return changes.map((change) => {
    const value = asObject(change);
    const info = asObject(value?.info);
    const id = info?.revisionId ?? info?.id;
    if (!value || typeof id !== 'string' || Number.isFinite(Number(id))) return change;
    return { ...value, info: trackedInfo(info) };
  });
}

/** Serialized ids for a session revision's deletion and insertion. */
export interface RevisionNumbers {
  insertion?: number;
  deletion?: number;
}

type RevisionIds = ReadonlyMap<string, RevisionNumbers>;

function mappedRevisionId(
  raw: unknown,
  kind: keyof RevisionNumbers,
  revisionIds: RevisionIds | undefined
): number | undefined {
  const value = asObject(raw);
  const id = value?.revisionId ?? value?.id;
  return typeof id === 'string' ? revisionIds?.get(id)?.[kind] : undefined;
}

function createHyperlink(attributes: Attrs): Hyperlink | null {
  const link = asObject(attributes.hyperlink);
  if (!link) return null;
  const href = asString(link.href) || '';
  if (href.startsWith('#')) {
    return {
      type: 'hyperlink',
      anchor: href.slice(1),
      tooltip: asString(link.tooltip) || undefined,
      children: [],
    };
  }
  return {
    type: 'hyperlink',
    href,
    tooltip: asString(link.tooltip) || undefined,
    rId: asString(link.rId) || undefined,
    children: [],
  };
}

function hyperlinkKey(attributes: Attrs): string | null {
  const link = asObject(attributes.hyperlink);
  return link ? asString(link.href) || '' : null;
}

function fieldFromPayload(
  payload: Attrs,
  attributes: Attrs,
  format: RunFormatter
): SimpleField | ComplexField {
  const fieldData = asString(payload.fieldData);
  if (fieldData && fieldData.length <= 2_000_000) {
    try {
      const stored = JSON.parse(fieldData) as SimpleField | ComplexField;
      if (fieldData.includes('"media:')) resolveMediaSrcs(stored);
      const children = stored.type === 'simpleField' ? stored.content : stored.fieldResult;
      if (
        (stored.type === 'simpleField' || stored.type === 'complexField') &&
        typeof stored.instruction === 'string' &&
        Array.isArray(children)
      ) {
        const displayMode = asString(payload.displayMode) as 'result' | 'code' | undefined;
        if (stored.fieldTree && displayMode) stored.fieldTree.displayMode = displayMode;
        return stored;
      }
    } catch {
      // Malformed editor cache: rebuild the same minimal inert field as the save path.
    }
  }

  const formatting = format(attributes);
  const displayRun: Run = {
    type: 'run',
    content: [{ type: 'text', text: asString(payload.displayText) || '' }],
    ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
  };
  const instruction = asString(payload.instruction) || '';
  const fieldType = (asString(payload.fieldType) || 'UNKNOWN') as FieldType;
  const fldLock = payload.fldLock === true || undefined;
  const dirty = payload.dirty === true || undefined;
  const displayMode = (asString(payload.displayMode) || 'result') as 'result' | 'code';
  if (payload.fieldKind === 'complex') {
    return {
      type: 'complexField',
      instruction,
      fieldType,
      fieldCode: [],
      fieldResult: [displayRun],
      fldLock,
      dirty,
      fieldTree: { version: 1, displayMode },
    };
  }
  return {
    type: 'simpleField',
    instruction,
    fieldType,
    content: [displayRun],
    fldLock,
    dirty,
    fieldTree: { version: 1, displayMode },
  };
}

function mathFromPayload(payload: Attrs): MathEquation {
  return {
    type: 'mathEquation',
    display: (asString(payload.display) as 'inline' | 'block') || 'inline',
    ommlXml: asString(payload.ommlXml) || '',
    plainText: asString(payload.plainText) || undefined,
  };
}

function horizontalRuleRun(payload: Attrs, attributes: Attrs, format: RunFormatter): Run {
  const value = asObject(payload.rule);
  const height = asFiniteNumber(value?.height);
  if (
    !value || height === undefined ||
    (value.width != null && asFiniteNumber(value.width) === undefined) ||
    (value.widthPercent != null && asFiniteNumber(value.widthPercent) === undefined) ||
    typeof value.alignment !== 'string' || typeof value.noShade !== 'boolean' ||
    typeof value.color !== 'string' || typeof value.xml !== 'string'
  ) throw new Error('Malformed horizontalRule embed payload');
  const rule: HorizontalRuleContent['rule'] = {
    width: asFiniteNumber(value.width) ?? null,
    widthPercent: asFiniteNumber(value.widthPercent) ?? null,
    height,
    alignment: value.alignment,
    noShade: value.noShade,
    color: value.color,
    xml: value.xml,
  };
  const formatting = format(attributes);
  return {
    type: 'run',
    content: [{ type: 'horizontalRule', rule }],
    ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
  };
}

/** Reads the `media:{n}` image sources of the projection under way as `data:` URLs. */
let projectedMedia: ((token: string) => string | null) | null = null;

function mediaSrc(src: string): string {
  return projectedMedia && src.startsWith('media:') ? (projectedMedia(src) ?? src) : src;
}

function resolveMediaSrcs(value: unknown): void {
  if (Array.isArray(value)) {
    for (const item of value) resolveMediaSrcs(item);
    return;
  }
  const object = asObject(value);
  if (!object) return;
  for (const [key, field] of Object.entries(object)) {
    if (key === 'src' && typeof field === 'string') object[key] = mediaSrc(field);
    else resolveMediaSrcs(field);
  }
}

function imageFromPayload(payload: Attrs): Image {
  const attrs = payload as YrsImageAttrs & Attrs;
  const wrap: Image['wrap'] = {
    type: (asString(attrs.wrapType) || 'inline') as Image['wrap']['type'],
  };
  if (attrs.distTop !== undefined) wrap.distT = pixelsToEmu(Number(attrs.distTop));
  if (attrs.distBottom !== undefined) wrap.distB = pixelsToEmu(Number(attrs.distBottom));
  if (attrs.distLeft !== undefined) wrap.distL = pixelsToEmu(Number(attrs.distLeft));
  if (attrs.distRight !== undefined) wrap.distR = pixelsToEmu(Number(attrs.distRight));
  if (attrs.wrapText) wrap.wrapText = attrs.wrapText as Image['wrap']['wrapText'];

  const image: Image = {
    type: 'image',
    rId: asString(attrs.rId) || '',
    src: mediaSrc(asString(attrs.src) || ''),
    alt: asString(attrs.alt) || undefined,
    title: asString(attrs.title) || undefined,
    shapeType: asString(attrs.shapeType) || undefined,
    size: {
      width: pixelsToEmu(Number(attrs.width) || 0),
      height: pixelsToEmu(Number(attrs.height) || 0),
    },
    wrap,
  };

  if (attrs.transform) {
    const transform: NonNullable<Image['transform']> = {};
    const rotation = attrs.transform.match(/rotate\(([-\d.]+)deg\)/)?.[1];
    if (rotation) transform.rotation = Number.parseFloat(rotation);
    if (attrs.transform.includes('scaleX(-1)')) transform.flipH = true;
    if (attrs.transform.includes('scaleY(-1)')) transform.flipV = true;
    if (transform.rotation || transform.flipH || transform.flipV) image.transform = transform;
  }

  if (attrs.position?.horizontal && attrs.position.vertical) {
    image.position = {
      relativeHeight: attrs.position.relativeHeight,
      horizontal: {
        relativeTo: (attrs.position.horizontal.relativeTo || 'column') as NonNullable<
          Image['position']
        >['horizontal']['relativeTo'],
        alignment: attrs.position.horizontal.align as NonNullable<
          Image['position']
        >['horizontal']['alignment'],
        posOffset: attrs.position.horizontal.posOffset,
      },
      vertical: {
        relativeTo: (attrs.position.vertical.relativeTo || 'paragraph') as NonNullable<
          Image['position']
        >['vertical']['relativeTo'],
        alignment: attrs.position.vertical.align as NonNullable<
          Image['position']
        >['vertical']['alignment'],
        posOffset: attrs.position.vertical.posOffset,
      },
    };
  }

  if (attrs.borderWidth && attrs.borderWidth > 0) {
    const styles: Record<string, NonNullable<Image['outline']>['style']> = {
      solid: 'solid',
      dotted: 'dot',
      dashed: 'dash',
      double: 'solid',
      groove: 'solid',
      ridge: 'solid',
      inset: 'solid',
      outset: 'solid',
    };
    image.outline = {
      width: pixelsToEmu(attrs.borderWidth),
      color: attrs.borderColor ? { rgb: attrs.borderColor.replace('#', '') } : undefined,
      style: (attrs.borderStyle && styles[attrs.borderStyle]) || 'solid',
    };
  }
  if (attrs.hlinkHref) image.hlinkHref = attrs.hlinkHref;

  const crop: NonNullable<Image['crop']> = {};
  if (attrs.cropTop != null) crop.top = attrs.cropTop;
  if (attrs.cropRight != null) crop.right = attrs.cropRight;
  if (attrs.cropBottom != null) crop.bottom = attrs.cropBottom;
  if (attrs.cropLeft != null) crop.left = attrs.cropLeft;
  if (Object.keys(crop).length > 0) image.crop = crop;
  if (attrs.opacity != null && attrs.opacity < 1) image.opacity = attrs.opacity;
  if (attrs.layoutInCell != null) image.layoutInCell = attrs.layoutInCell;
  if (attrs.allowOverlap != null) image.allowOverlap = attrs.allowOverlap;

  const padding: NonNullable<Image['padding']> = {};
  if (attrs.effectExtentTop != null) padding.top = pixelsToEmu(attrs.effectExtentTop);
  if (attrs.effectExtentBottom != null) padding.bottom = pixelsToEmu(attrs.effectExtentBottom);
  if (attrs.effectExtentLeft != null) padding.left = pixelsToEmu(attrs.effectExtentLeft);
  if (attrs.effectExtentRight != null) padding.right = pixelsToEmu(attrs.effectExtentRight);
  if (Object.keys(padding).length > 0) image.padding = padding;

  return image;
}

function imageRunFromPayload(payload: Attrs): Run {
  return { type: 'run', content: [{ type: 'drawing', image: imageFromPayload(payload) }] };
}

/** The seeded shape, carrying the text body and everything else no payload field describes. */
function storedShape(value: unknown): Shape | undefined {
  const json = asString(value);
  if (!json || json.length > 2_000_000) return undefined;
  try {
    const parsed = JSON.parse(json) as Shape;
    if (parsed?.type !== 'shape' || typeof parsed.shapeType !== 'string') return undefined;
    if (!asObject(parsed.size)) parsed.size = { width: 0, height: 0 };
    if (json.includes('"media:')) resolveMediaSrcs(parsed);
    return parsed;
  } catch {
    return undefined;
  }
}

function chartRunFromPayload(payload: Attrs, recovery: ChartRecovery): Run | null {
  const json = asString(payload.chartJson);
  if (!json) {
    recovery.warn('chart run carries no chart payload; keeping the run out of the output');
    return null;
  }
  try {
    const chart = JSON.parse(json) as Chart;
    if (json.includes('"media:')) resolveMediaSrcs(chart);
    if (chart?.type !== 'chart' || typeof chart.chartType !== 'string') {
      recovery.warn('chart run carries a malformed chart payload; keeping the run out of the output');
      return null;
    }
    if (!chart.drawingXml) {
      const drawingXml = recovery.drawingFor(chart.rId, chart.path);
      if (!drawingXml) {
        recovery.unrecoverable(chart.rId, chart.path);
        return null;
      }
      chart.drawingXml = drawingXml;
    }
    return { type: 'run', content: [{ type: 'chart', chart }] };
  } catch {
    recovery.warn('chart run carries an unreadable chart payload; keeping the run out of the output');
    return null;
  }
}

function opaqueDrawingRunFromPayload(
  payload: Attrs,
  attributes: Attrs,
  format: RunFormatter
): Run | null {
  const kind = asString(payload.kind);
  const xml = asString(payload.xml);
  if (!kind || !xml) return null;
  const preview = asObject(payload.preview);
  const previews = Array.isArray(payload.previews)
    ? payload.previews.map(asObject).filter((image): image is Attrs => !!image).map(imageFromPayload)
    : [];
  const formatting = format(attributes);
  return {
    type: 'run',
    content: [{
      type: 'opaqueDrawing', kind, xml,
      ...(preview ? { preview: imageFromPayload(preview) } : {}),
      ...(previews.length ? { previews } : {}),
    }],
    ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
  };
}

function shapeRunFromPayload(payload: Attrs): Run {
  const shape: Shape = storedShape(payload.shapeJson) ?? {
    type: 'shape',
    shapeType: 'rect',
    size: { width: 0, height: 0 },
  };
  const shapeType = asString(payload.shapeType);
  if (shapeType) shape.shapeType = shapeType as Shape['shapeType'];
  const shapeId = asString(payload.shapeId);
  if (shapeId) shape.id = shapeId;
  if (payload.width) shape.size = { ...shape.size, width: pixelsToEmu(Number(payload.width)) };
  if (payload.height) shape.size = { ...shape.size, height: pixelsToEmu(Number(payload.height)) };
  if (Array.isArray(payload.geometryPath) && payload.geometryPath.length > 0) {
    shape.geometryPath = payload.geometryPath as NonNullable<Shape['geometryPath']>;
  }

  if (payload.fillType === 'gradient' && typeof payload.gradientStops === 'string') {
    try {
      const stops = JSON.parse(payload.gradientStops) as Array<{
        position: number;
        color: string;
      }>;
      shape.fill = {
        type: 'gradient',
        gradient: {
          type: (asString(payload.gradientType) || 'linear') as NonNullable<
            NonNullable<Shape['fill']>['gradient']
          >['type'],
          angle: asFiniteNumber(payload.gradientAngle) || undefined,
          stops: stops.map((stop) => ({
            position: stop.position,
            color: { rgb: stop.color.replace('#', '') },
          })),
        },
      };
    } catch {
      shape.fill = {
        type: 'solid',
        color: { rgb: (asString(payload.fillColor) || '000000').replace('#', '') },
      };
    }
  } else if (typeof payload.fillColor === 'string') {
    shape.fill = {
      type: (asString(payload.fillType) || 'solid') as 'solid' | 'none',
      color: { rgb: payload.fillColor.replace('#', '') },
    };
  } else if (payload.fillType === 'none') {
    shape.fill = { type: 'none' };
  }

  if (typeof payload.outlineWidth === 'number' && payload.outlineWidth > 0) {
    const styles: Record<string, NonNullable<Shape['outline']>['style']> = {
      solid: 'solid',
      dotted: 'dot',
      dashed: 'dash',
    };
    const outlineStyle = asString(payload.outlineStyle);
    shape.outline = {
      width: pixelsToEmu(payload.outlineWidth),
      color:
        typeof payload.outlineColor === 'string'
          ? { rgb: payload.outlineColor.replace('#', '') }
          : undefined,
      style: (outlineStyle && styles[outlineStyle]) || 'solid',
    };
  }

  const transform: NonNullable<Shape['transform']> = {};
  if (typeof payload.rotation === 'number') transform.rotation = payload.rotation;
  else if (typeof payload.transform === 'string') {
    const rotation = payload.transform.match(/rotate\(([-\d.]+)deg\)/)?.[1];
    if (rotation) transform.rotation = Number.parseFloat(rotation);
  }
  if (payload.flipH || String(payload.transform || '').includes('scaleX(-1)'))
    transform.flipH = true;
  if (payload.flipV || String(payload.transform || '').includes('scaleY(-1)'))
    transform.flipV = true;
  if (transform.rotation || transform.flipH || transform.flipV) shape.transform = transform;

  return { type: 'run', content: [{ type: 'shape', shape }] };
}

function inlineSdtFromPayload(
  payload: Attrs,
  recovery: ChartRecovery,
  format: RunFormatter
): InlineSdt {
  let properties: SdtProperties = sdtAttrsToProps(payload);
  const propertiesJson = asString(payload.propertiesJson);
  if (propertiesJson && propertiesJson.length <= 1_000_000) {
    try {
      const parsed = JSON.parse(propertiesJson) as SdtProperties;
      if (propertiesJson.includes('"media:')) resolveMediaSrcs(parsed);
      if (parsed && typeof parsed === 'object' && typeof parsed.sdtType === 'string') {
        properties = parsed;
      }
    } catch {
      // Keep the individually projected properties.
    }
  }
  const items: InlineItem[] = [];
  for (const raw of Array.isArray(payload.content) ? payload.content : []) {
    const entry = asObject(raw);
    const kind = asString(entry?.kind);
    if (!entry || !kind) continue;
    const attributes = asObject(entry.attrs) ?? {};
    if (kind === 'text') {
      const text = asString(entry.text);
      if (text !== undefined) items.push({ kind: 'text', text, attributes });
      continue;
    }
    items.push({
      kind: 'embed',
      embedKind: kind,
      payload: asObject(entry.payload) ?? {},
      attributes,
    });
  }
  let content = inlineSdtContent(buildParagraphContent(items, recovery, format, undefined, true));
  const authoredValue = contentControlValue(payload.value);
  if (authoredValue) {
    try {
      const applied = applyContentControlValue(properties, authoredValue);
      properties = applied.properties;
      const display = applied.content[0];
      content = inlineSdtContent(display?.type === 'paragraph' ? display.content : []);
    } catch {
      // Preserve the embedded control when a malformed authored value cannot
      // be applied to its captured OOXML properties.
    }
  }
  return { type: 'inlineSdt', properties, content };
}

function inlineSdtContent(content: ParagraphContent[]): InlineSdt['content'] {
  return content.filter(
    (child): child is InlineSdt['content'][number] =>
      child.type === 'run' ||
      child.type === 'hyperlink' ||
      child.type === 'simpleField' ||
      child.type === 'complexField' ||
      child.type === 'inlineSdt' ||
      child.type === 'mathEquation' ||
      child.type === 'insertion' ||
      child.type === 'deletion' ||
      child.type === 'moveFrom' ||
      child.type === 'moveTo'
  );
}

function fieldInlineContent(content: ParagraphContent[]): FieldInlineContent[] {
  return content.filter(
    (child): child is FieldInlineContent =>
      child.type === 'run' ||
      child.type === 'hyperlink' ||
      child.type === 'simpleField' ||
      child.type === 'complexField' ||
      child.type === 'inlineSdt' ||
      child.type === 'mathEquation'
  );
}

function contentControlValue(value: unknown): ContentControlValue | null {
  const authored = asObject(value);
  if (authored?.kind === 'checkbox' && typeof authored.checked === 'boolean') {
    return { kind: 'checkbox', checked: authored.checked };
  }
  if (authored?.kind === 'dropdown' && typeof authored.value === 'string') {
    return { kind: 'dropdown', value: authored.value };
  }
  if (authored?.kind === 'date' && typeof authored.date === 'string') {
    return { kind: 'date', date: authored.date };
  }
  return null;
}

function commentReferenceFromPayload(payload: Attrs): Run | null {
  if (payload.modelKind !== 'commentReference') return null;
  const id = asFiniteNumber(payload.commentId);
  return {
    type: 'run',
    content: [{ type: 'commentReference', ...(id !== undefined ? { id } : {}) }],
  };
}

function ordinaryContentForItem(
  item: InlineItem,
  recovery: ChartRecovery,
  format: RunFormatter
): ParagraphContent | null {
  if (item.kind === 'text') return createTextRun(item.text, item.attributes, format);
  switch (item.embedKind) {
    case 'break':
      return { type: 'run', content: [{ type: 'break', breakType: 'textWrapping' }] };
    case 'tab':
      return tabRun(item.attributes, format);
    case 'image':
      return imageRunFromPayload(item.payload);
    case 'horizontalRule':
      return horizontalRuleRun(item.payload, item.attributes, format);
    case 'shape':
      return shapeRunFromPayload(item.payload);
    case 'chart':
      return chartRunFromPayload(item.payload, recovery);
    case 'opaqueDrawing':
      return opaqueDrawingRunFromPayload(item.payload, item.attributes, format);
    case 'field':
      return (
        commentReferenceFromPayload(item.payload) ??
        fieldFromPayload(item.payload, item.attributes, format)
      );
    case 'math':
      return mathFromPayload(item.payload);
    case 'sdt':
      return inlineSdtFromPayload(item.payload, recovery, format);
    case 'noteRef': {
      const footnote = item.payload.footnoteRefId;
      const endnote = item.payload.endnoteRefId;
      return {
        type: 'run',
        content: [
          footnote !== undefined
            ? { type: 'footnoteRef', id: revisionId(footnote) }
            : { type: 'endnoteRef', id: revisionId(endnote) },
        ],
      };
    }
    default:
      return null;
  }
}

function tabRun(attributes: Attrs, format: RunFormatter): Run {
  const formatting = format(attributes);
  return {
    type: 'run',
    content: [{ type: 'tab' }],
    ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
  };
}

function trackedContentForItem(item: InlineItem, info: TrackedChangeInfo, inControl: boolean, recovery: ChartRecovery, format: RunFormatter): ParagraphContent | null {
  let child: TrackedWrapper['content'][number];
  if (inControl) {
    const ordinary = ordinaryContentForItem(item, recovery, format);
    child = inlineSdtContent(ordinary ? [ordinary] : [])[0] ?? { type: 'run', content: [] };
  } else if (item.kind === 'embed' && item.embedKind === 'sdt') child = inlineSdtFromPayload(item.payload, recovery, format);
  else if (item.kind === 'embed' && item.embedKind === 'image') child = imageRunFromPayload(item.payload);
  else if (item.kind === 'embed' && item.embedKind === 'tab') child = tabRun(item.attributes, format);
  else if (item.kind === 'embed' && item.embedKind === 'horizontalRule')
    child = horizontalRuleRun(item.payload, item.attributes, format);
  else if (item.kind === 'embed' && item.embedKind === 'shape')
    child = shapeRunFromPayload(item.payload);
  else if (item.kind === 'embed' && item.embedKind === 'chart')
    child = chartRunFromPayload(item.payload, recovery) ?? { type: 'run', content: [] };
  else if (item.kind === 'embed' && item.embedKind === 'opaqueDrawing')
    child = opaqueDrawingRunFromPayload(item.payload, item.attributes, format) ?? { type: 'run', content: [] };
  else if (item.kind === 'text') child = createTextRun(item.text, item.attributes, format);
  else child = { type: 'run', content: [] };

  if (item.kind === 'embed' && item.embedKind === 'chart' && child.type === 'run' && child.content.length === 0) return null;
  const hyperlink = createHyperlink(item.attributes);
  if (hyperlink && (child.type === 'run' || child.type === 'mathEquation' || child.type === 'inlineSdt' || child.type === 'simpleField' || child.type === 'complexField')) {
    if (child.type === 'run') hyperlink.children.push(child);
    else hyperlink.structuredChildren = [child];
    child = hyperlink;
  }

  const raw = asObject(item.attributes.ins) ?? asObject(item.attributes.del);
  const isMovePair = raw?.isMovePair === true;
  if (item.attributes.ins) {
    return isMovePair
      ? { type: 'moveTo', info, content: [child] }
      : { type: 'insertion', info, content: [child] };
  }
  return isMovePair
    ? { type: 'moveFrom', info, content: [child] }
    : { type: 'deletion', info, content: [child] };
}

/**
 * Adds an item to a hyperlink: a run to its children, and every child, in
 * order, to its structured children once it has them.
 */
function addToHyperlink(
  hyperlink: Hyperlink,
  item: InlineItem,
  recovery: ChartRecovery,
  format: RunFormatter
): void {
  let child: HyperlinkContent | undefined;
  if (item.kind === 'text') child = createTextRun(item.text, item.attributes, format);
  else if (item.embedKind === 'break') {
    child = { type: 'run', content: [{ type: 'break', breakType: 'textWrapping' }] };
  } else if (item.embedKind === 'tab') child = { type: 'run', content: [{ type: 'tab' }] };
  else if (item.embedKind === 'horizontalRule') {
    child = horizontalRuleRun(item.payload, item.attributes, format);
  } else if (item.embedKind === 'image') child = imageRunFromPayload(item.payload);
  else if (item.embedKind === 'shape') child = shapeRunFromPayload(item.payload);
  else if (item.embedKind === 'chart') child = chartRunFromPayload(item.payload, recovery) ?? undefined;
  else if (item.embedKind === 'opaqueDrawing') child = opaqueDrawingRunFromPayload(item.payload, item.attributes, format) ?? undefined;
  else if (item.embedKind === 'field') {
    child =
      commentReferenceFromPayload(item.payload) ??
      fieldFromPayload(item.payload, item.attributes, format);
  } else if (item.embedKind === 'math') child = mathFromPayload(item.payload);
  else if (item.embedKind === 'sdt') {
    const control = inlineSdtFromPayload(item.payload, recovery, format);
    if (item.attributes.ins || item.attributes.del || hasTrackedControlContent(control)) child = control;
  }
  if (!child) return;
  if (child.type === 'run') hyperlink.children.push(child);
  if (child.type !== 'run' || hyperlink.structuredChildren) {
    (hyperlink.structuredChildren ??= [...hyperlink.children]).push(child);
  }
}

function projectionSignature(items: InlineItem[]): string {
  const normalized: InlineItem[] = [];
  for (const source of items) {
    const attributes = { ...source.attributes };
    delete attributes.fieldResult;
    const item: InlineItem = source.kind === 'embed' && source.embedKind === 'tab'
      ? { kind: 'text', text: '\t', attributes }
      : { ...source, attributes };
    const previous = normalized.at(-1);
    if (item.kind === 'text' && previous?.kind === 'text' && stableStringify(previous.attributes) === stableStringify(attributes)) {
      previous.text += item.text;
    } else normalized.push(item);
  }
  return stableStringify(normalized);
}

function restoreProjectedFieldResults(
  items: InlineItem[],
  recovery: ChartRecovery,
  format: RunFormatter
): InlineItem[] {
  const owners: { id: number; owner: EmbedItem; position: number }[] = [];
  for (const [position, item] of items.entries()) {
    if (item.kind !== 'embed' || item.embedKind !== 'field') continue;
    const id = asFiniteNumber(asObject(item.payload.resultProjection)?.id);
    if (id !== undefined) owners.push({ id, owner: item, position });
  }
  if (owners.length === 0) return items;
  const groups = new Map<EmbedItem, Map<number, InlineItem[]>>();
  const remaining = items.filter((item, position) => {
    const marker = asObject(item.attributes.fieldResult);
    const id = asFiniteNumber(marker?.id);
    const index = asFiniteNumber(marker?.index);
    const owner = owners.find((entry) => entry.id === id && entry.position > position)?.owner;
    if (index === undefined || !owner) return true;
    const children = groups.get(owner) ?? new Map<number, InlineItem[]>();
    const group = children.get(index) ?? [];
    const attributes = { ...item.attributes };
    delete attributes.fieldResult;
    group.push({ ...item, attributes });
    children.set(index, group);
    groups.set(owner, children);
    return false;
  });
  for (const { owner } of owners) {
    const stored = fieldFromPayload(owner.payload, owner.attributes, format);
    if (stored.type !== 'complexField') continue;
    const projection = asObject(owner.payload.resultProjection);
    const originals = Array.isArray(projection?.children) ? projection.children : [];
    const replacements = new Map<number, FieldInlineContent[]>();
    for (const raw of originals) {
      const child = asObject(raw);
      const index = asFiniteNumber(child?.index);
      if (index === undefined || !Array.isArray(child?.items)) continue;
      const current = groups.get(owner)?.get(index) ?? [];
      if (projectionSignature(current) === projectionSignature(child.items as InlineItem[])) continue;
      const rebuilt = fieldInlineContent(buildParagraphContent(current, recovery, format));
      const original = index < 0 ? stored.structuredCode?.inline?.[-index - 1] : stored.structuredResult?.inline?.[index];
      if (original?.type === 'hyperlink' && rebuilt.length === 1 && rebuilt[0]?.type === 'hyperlink') {
        rebuilt[0] = { ...original, ...rebuilt[0], structuredChildren: rebuilt[0].structuredChildren };
      }
      replacements.set(index, rebuilt);
    }
    if (replacements.size === 0) continue;
    const inline = (stored.structuredResult?.inline ?? []).flatMap((child, index) => replacements.get(index) ?? [child]);
    const code = stored.structuredCode?.inline?.flatMap((child, index) => replacements.get(-index - 1) ?? [child]);
    stored.structuredResult = { ...stored.structuredResult, inline };
    if (code) stored.structuredCode = { ...stored.structuredCode, inline: code };
    if (stored.fieldTree) {
      stored.fieldTree.result = { ...stored.fieldTree.result, inline };
      if (code) stored.fieldTree.code = { ...stored.fieldTree.code, inline: code };
    }
    stored.fieldResult = inline.flatMap((child) => child.type === 'run' ? [child]
      : child.type === 'hyperlink' ? child.children.filter((entry): entry is Run => entry.type === 'run') : []);
    owner.payload = { ...owner.payload, fieldData: JSON.stringify(stored) };
  }
  return remaining;
}

type TrackedWrapper = Extract<
  ParagraphContent,
  { type: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo' }
>;

function isTrackedWrapper(content: ParagraphContent | undefined): content is TrackedWrapper {
  return (
    content?.type === 'insertion' ||
    content?.type === 'deletion' ||
    content?.type === 'moveFrom' ||
    content?.type === 'moveTo'
  );
}

function buildParagraphContent(
  items: InlineItem[],
  recovery: ChartRecovery,
  format: RunFormatter,
  revisionIds?: RevisionIds,
  inControl = false
): ParagraphContent[] {
  items = restoreProjectedFieldResults(items, recovery, format);
  const content: ParagraphContent[] = [];
  let currentRun: Run | null = null;
  let currentFormattingKey: string | null = null;
  let currentHyperlink: Hyperlink | null = null;

  const flushRun = (): void => {
    if (currentRun) content.push(currentRun);
    currentRun = null;
    currentFormattingKey = null;
  };
  const flushHyperlink = (): void => {
    if (currentHyperlink) content.push(currentHyperlink);
    currentHyperlink = null;
  };

  for (const item of items) {
    // A note reference is handled before tracked/link marks.
    if (item.kind === 'embed' && item.embedKind === 'noteRef') {
      flushRun();
      flushHyperlink();
      const note = ordinaryContentForItem(item, recovery, format);
      if (note) content.push(note);
      continue;
    }

    const raw = item.attributes.ins ?? item.attributes.del;
    const revision = trackedInfo(
      raw,
      false,
      mappedRevisionId(raw, item.attributes.ins ? 'insertion' : 'deletion', revisionIds)
    );
    if (revision) {
      flushRun();
      flushHyperlink();
      const tracked = trackedContentForItem(item, revision, inControl, recovery, format);
      if (!tracked) continue;
      const previous = content[content.length - 1];
      if (
        isTrackedWrapper(previous) &&
        isTrackedWrapper(tracked) &&
        previous.type === tracked.type &&
        previous.info.id === tracked.info.id &&
        previous.info.author === tracked.info.author &&
        previous.info.date === tracked.info.date
      ) {
        previous.content.push(...tracked.content);
      } else {
        content.push(tracked);
      }
      continue;
    }

    const linkKey = hyperlinkKey(item.attributes);
    if (linkKey !== null) {
      flushRun();
      const currentKey = currentHyperlink
        ? currentHyperlink.href || (currentHyperlink.anchor ? `#${currentHyperlink.anchor}` : '')
        : null;
      if (!currentHyperlink || currentKey !== linkKey) {
        flushHyperlink();
        currentHyperlink = createHyperlink(item.attributes);
      }
      if (currentHyperlink) addToHyperlink(currentHyperlink, item, recovery, format);
      continue;
    }

    flushHyperlink();
    if (item.kind === 'text') {
      const key = stableStringify(formattingAttrs(item.attributes));
      const nextRun = createTextRun(item.text, item.attributes, format);
      if (currentRun && currentFormattingKey === key) appendTextRun(currentRun, nextRun);
      else {
        flushRun();
        currentRun = nextRun;
        currentFormattingKey = key;
      }
      continue;
    }

    flushRun();
    const child = ordinaryContentForItem(item, recovery, format);
    if (child) content.push(child);
  }

  flushRun();
  flushHyperlink();
  return content;
}

function marksKeyToYrsAttrs(marksKey: string | undefined): Attrs | null {
  if (!marksKey) return {};
  const attrs: Attrs = {};
  // Mark attribute JSON does not normally contain `|`; if a custom string
  // does, declining restoration merely coalesces runs and never loses text.
  for (const part of marksKey.split('|')) {
    const colon = part.indexOf(':');
    if (colon <= 0) return null;
    const name = part.slice(0, colon);
    let value: Attrs;
    try {
      value = JSON.parse(part.slice(colon + 1)) as Attrs;
    } catch {
      return null;
    }
    if (name === 'comment' || name === 'footnoteRef') continue;
    if (BOOLEAN_MARKS.has(name)) attrs[name] = true;
    else if (name === 'boldCs' || name === 'italicCs') attrs[name] = value.value;
    else if (name === 'highlight') attrs.highlight = value.color;
    else if (name === 'insertion' || name === 'deletion') {
      attrs[name === 'insertion' ? 'ins' : 'del'] = dropNulls({
        id: value.revisionId,
        author: value.author,
        date: value.date,
        isMovePair: value.isMovePair,
      });
    } else attrs[name] = dropNulls(value);
  }
  return attrs;
}

/** Rebuilds the note number marks a run held; they occupy no story unit. */
function noteMarkContent(boundary: OriginalRunBoundary): RunContent[] | null {
  if (!boundary.noteMarks?.length) return null;
  return boundary.noteMarks.map((mark) => ({
    type: mark === 'endnote' ? 'endnoteRefMark' : 'footnoteRefMark',
  }));
}

/**
 * Rebuilds a run from its recorded boundary. Flow breaks occupy no story unit,
 * so without their recorded offsets an authored page break would not be saved.
 */
function boundaryContent(boundary: OriginalRunBoundary, formatting: TextFormatting): RunContent[] {
  const notes = noteMarkContent(boundary);
  if (notes) return notes;
  if (!boundary.breaks?.length) return runContentForText(boundary.text, formatting);
  const content: RunContent[] = [];
  let cursor = 0;
  for (const entry of boundary.breaks) {
    const at = Math.min(Math.max(entry.offset, cursor), boundary.text.length);
    if (at > cursor) {
      content.push(...runContentForText(boundary.text.slice(cursor, at), formatting));
    }
    content.push({ type: 'break', breakType: entry.type });
    cursor = at;
  }
  if (cursor < boundary.text.length) {
    content.push(...runContentForText(boundary.text.slice(cursor), formatting));
  }
  return content;
}

function sourceBreakBoundaries(paragraph: Paragraph | undefined): OriginalRunBoundary[] {
  const recorded = paragraph && projectedBlocks.get(paragraph)?.sourceRuns;
  if (recorded) return recorded;
  const boundaries: OriginalRunBoundary[] = [];
  for (const child of paragraph?.content ?? []) {
    if (child.type !== 'run') {
      if (
        child.type === 'commentRangeStart' || child.type === 'commentRangeEnd' ||
        child.type === 'bookmarkStart' || child.type === 'bookmarkEnd' || child.type === 'rawXml'
      ) continue;
      return [];
    }
    let text = '';
    const breaks: NonNullable<OriginalRunBoundary['breaks']> = [];
    for (const entry of child.content) {
      if (entry.type === 'text') text += entry.text;
      else if (entry.type === 'commentReference') text += '\uFFFC';
      else if (entry.type === 'break' && (entry.breakType === 'page' || entry.breakType === 'column')) {
        breaks.push({ offset: text.length, type: entry.breakType });
      } else return [];
    }
    boundaries.push({ text, breaks });
  }
  return boundaries;
}

function liveRunBoundaries(
  boundaries: OriginalRunBoundary[],
  live: LiveParagraphBreaks
): OriginalRunBoundary[] {
  const count = boundaries.reduce((sum, boundary) =>
    sum + (noteMarkContent(boundary) ? 0 : (boundary.breaks?.length ?? 0)), 0);
  if (count === 0) return boundaries;
  const source = boundaries.map((boundary): ParagraphContent => ({
    type: 'run', content: boundaryContent(boundary, {}),
  }));
  const hoisted = count - trailingBlockBreaks({ type: 'paragraph', content: source }).length;
  const visible = flowTokens(source).includes('visible');
  let index = 0;
  let leading = 0;
  let trailing = 0;
  return boundaries.map(boundary => {
    if (!boundary.breaks?.length || noteMarkContent(boundary)) return boundary;
    const breaks = boundary.breaks.filter(entry => {
      const before = index++ < hoisted;
      if (before && visible && entry.type === 'page') return live.leadingPage;
      if (before) {
        if (entry.type !== live.leading[leading]) return false;
        leading += 1;
      } else {
        if (entry.type !== live.trailing[trailing]) return false;
        trailing += 1;
      }
      return true;
    });
    return breaks.length === boundary.breaks.length ? boundary : { ...boundary, breaks };
  });
}

function paragraphBreaksAreLive(content: ParagraphContent[], live: LiveParagraphBreaks): boolean {
  const paragraph: Paragraph = { type: 'paragraph', content };
  const tokens = flowTokens(content);
  return (!(tokens[0] === 'page' && tokens.includes('visible')) || live.leadingPage) &&
    leadingBlockBreaks(paragraph).every((kind, index) => kind === live.leading[index]) &&
    trailingBlockBreaks(paragraph).every((kind, index) => kind === live.trailing[index]);
}

function restoreRecordedBreaks(
  content: ParagraphContent[],
  items: InlineItem[],
  boundaries: OriginalRunBoundary[],
  liveBreaks: LiveParagraphBreaks
): ParagraphContent[] {
  const breaks: { offset: number; type: BlockBreakKind }[] = [];
  let start = 0;
  for (const boundary of boundaries) {
    let cursor = 0;
    for (const entry of noteMarkContent(boundary) ? [] : (boundary.breaks ?? [])) {
      cursor = Math.min(Math.max(entry.offset, cursor), boundary.text.length);
      breaks.push({ offset: start + cursor, type: entry.type });
    }
    start += boundary.text.length;
  }
  if (breaks.length === 0) return content;
  const source = boundaries.map((boundary): ParagraphContent => ({
    type: 'run',
    content: boundaryContent(boundary, {}),
  }));
  const trailing = trailingBlockBreaks({ type: 'paragraph', content: source });
  const hoisted = breaks.length - trailing.length;
  const visible = flowTokens(source).includes('visible');
  const kept = breaks.filter(
    (entry, index) => index >= hoisted || (visible && entry.type === 'page')
  );
  const recorded = boundaries.map((boundary) => boundary.text).join('');
  const live = items
    .map((item) => (item.kind === 'text' ? item.text : isBlockBreak(item) ? '' : '\uFFFC'))
    .join('');
  let prefix = 0;
  const shared = Math.min(recorded.length, live.length);
  while (prefix < shared && recorded[prefix] === live[prefix]) prefix += 1;
  let suffix = 0;
  while (
    suffix < shared - prefix &&
    recorded[recorded.length - 1 - suffix] === live[live.length - 1 - suffix]
  ) {
    suffix += 1;
  }
  const length = content.reduce((sum, child) => sum + paragraphContentLength(child), 0);
  const placed = insertBoundaries(
    content,
    kept.map(({ offset }, id) => {
      const mapped = offset <= prefix
        ? offset
        : offset >= recorded.length - suffix
          ? offset + live.length - recorded.length
          : Math.min(offset, live.length - suffix);
      return { id, kind: 'start', offset: Math.min(Math.max(mapped, 0), length) };
    }),
    (boundary) => ({ type: 'run', content: [{ type: 'break', breakType: kept[boundary.id]!.type }] })
  );
  const placedTrailing = trailingBlockBreaks({ type: 'paragraph', content: placed });
  return placedTrailing.length === trailing.length &&
    placedTrailing.every((kind, index) => kind === trailing[index]) &&
    paragraphBreaksAreLive(placed, liveBreaks)
    ? placed
    : content;
}

function restoreOriginalRuns(
  content: ParagraphContent[],
  items: InlineItem[],
  boundaries: OriginalRunBoundary[] | undefined,
  baseParagraph: Paragraph | undefined,
  liveBreaks: LiveParagraphBreaks
): ParagraphContent[] {
  if (!boundaries?.length) {
    return restoreRecordedBreaks(content, items,
      liveRunBoundaries(sourceBreakBoundaries(baseParagraph), liveBreaks), liveBreaks);
  }
  const recorded = liveRunBoundaries(boundaries, liveBreaks);
  const fallback = (): ParagraphContent[] => restoreRecordedBreaks(content, items, recorded, liveBreaks);
  if (
    !content.every(
      (child) => child.type === 'run' && child.content.every((entry) => entry.type === 'text')
    ) ||
    items.some((item) => item.kind !== 'text' || item.attributes.hyperlink) ||
    recorded.some((boundary) => boundary.noteMarks?.length && boundary.text.length > 0)
  ) {
    return fallback();
  }
  const fullText = items.map((item) => (item.kind === 'text' ? item.text : '')).join('');
  if (fullText !== recorded.map((boundary) => boundary.text).join('')) return fallback();

  let itemIndex = 0;
  let itemOffset = 0;
  const restoredAttrs: Attrs[] = [];
  for (const boundary of recorded) {
    const expected = marksKeyToYrsAttrs(boundary.marksKey);
    if (!expected) return fallback();
    restoredAttrs.push(expected);
    let remaining = boundary.text.length;
    while (remaining > 0) {
      const item = items[itemIndex];
      if (item?.kind !== 'text') return fallback();
      if (stableStringify(formattingAttrs(item.attributes)) !== stableStringify(expected)) {
        return fallback();
      }
      const available = item.text.length - itemOffset;
      const consumed = Math.min(remaining, available);
      remaining -= consumed;
      itemOffset += consumed;
      if (itemOffset === item.text.length) {
        itemIndex += 1;
        itemOffset = 0;
      }
    }
  }

  if (itemIndex !== items.length || itemOffset !== 0) return fallback();

  const restored = recorded.map((boundary, index) => {
    const run: Run = {
      type: 'run',
      content: boundaryContent(boundary, attrsToTextFormatting(restoredAttrs[index])),
    };
    const formatting = boundary.formatting;
    if (formatting && Object.keys(formatting).length > 0) run.formatting = formatting;
    if (boundary.propertyChanges?.length) run.propertyChanges = boundary.propertyChanges;
    return run;
  });
  return paragraphBreaksAreLive(restored, liveBreaks) ? restored : content;
}

/** The story units a run entry projects from: one per UTF-16 unit of text, one per inline embed. */
function runContentLength(content: RunContent): number {
  if (content.type === 'text' || content.type === 'instrText') return content.text.length;
  if (content.type === 'symbol') return content.char.length;
  if (
    content.type === 'tab' ||
    content.type === 'softHyphen' ||
    content.type === 'noBreakHyphen' ||
    content.type === 'footnoteRef' ||
    content.type === 'endnoteRef' ||
    content.type === 'horizontalRule' ||
    content.type === 'commentReference' ||
    content.type === 'drawing' ||
    content.type === 'shape' ||
    content.type === 'chart' ||
    content.type === 'opaqueDrawing' ||
    (content.type === 'break' &&
      (content.breakType === undefined || content.breakType === 'textWrapping'))
  ) {
    return 1;
  }
  return 0;
}

function runTextLength(run: Run): number {
  return run.content.reduce((length, content) => length + runContentLength(content), 0);
}

/** The story units a hyperlink child spans: its text, or one for a field or an equation. */
function linkChildLength(child: HyperlinkContent): number {
  if (child.type === 'run') return runTextLength(child);
  return child.type === 'simpleField' ||
    child.type === 'complexField' ||
    child.type === 'inlineSdt' ||
    child.type === 'mathEquation'
    ? 1
    : 0;
}

/**
 * The story units projected content spans. Fields, controls and equations
 * are one embed each; a complex field's hyperlinks and simple fields are
 * seeded ahead of it.
 */
function paragraphContentLength(content: ParagraphContent): number {
  switch (content.type) {
    case 'run':
      return runTextLength(content);
    case 'hyperlink':
      return (content.structuredChildren ?? content.children).reduce(
        (sum, child) => sum + linkChildLength(child),
        0
      );
    case 'complexField':
      return [
        ...(content.structuredCode?.inline ?? []),
        ...(content.structuredResult?.inline ?? []),
      ]
        .filter((child) => child.type === 'hyperlink' || child.type === 'simpleField')
        .reduce((sum, child) => sum + paragraphContentLength(child), 1);
    case 'simpleField':
    case 'inlineSdt':
    case 'mathEquation':
      return 1;
    case 'insertion':
    case 'deletion':
    case 'moveFrom':
    case 'moveTo':
      return content.content.reduce((sum, child) => sum + paragraphContentLength(child), 0);
    default:
      return 0;
  }
}

type Split<T> = [T | null, T | null];

/** Splits `items` `at` story units in, or null when an item there cannot be split. */
function splitItems<T>(
  items: readonly T[],
  at: number,
  length: (item: T) => number,
  split: (item: T, at: number) => Split<T> | null
): [T[], T[]] | null {
  const left: T[] = [];
  const right: T[] = [];
  let offset = 0;
  for (const item of items) {
    const width = length(item);
    if (offset >= at) right.push(item);
    else if (offset + width <= at) left.push(item);
    else {
      const parts = split(item, at - offset);
      if (!parts) return null;
      if (parts[0]) left.push(parts[0]);
      if (parts[1]) right.push(parts[1]);
    }
    offset += width;
  }
  return [left, right];
}

function splitRunContent(content: RunContent, at: number): Split<RunContent> | null {
  if (content.type !== 'text' && content.type !== 'instrText') return null;
  return [
    { ...content, text: content.text.slice(0, at) },
    { ...content, text: content.text.slice(at) },
  ];
}

/** One side of a split hyperlink, holding `children` in order. */
function linkPart(link: Hyperlink, children: HyperlinkContent[]): Hyperlink | null {
  if (children.length === 0) return null;
  const legacy = children.filter(
    (child): child is Hyperlink['children'][number] =>
      child.type === 'run' || child.type === 'bookmarkStart' || child.type === 'bookmarkEnd'
  );
  return link.structuredChildren
    ? { ...link, children: legacy, structuredChildren: children }
    : { ...link, children: legacy };
}

/**
 * Splits content `at` story units in: runs between their units, hyperlinks
 * and tracked changes into two halves. Null for content that holds no
 * position there, such as a field.
 */
function splitContent(content: ParagraphContent, at: number): Split<ParagraphContent> | null {
  if (content.type === 'run') {
    const parts = splitItems(content.content, at, runContentLength, splitRunContent);
    return parts && [
      parts[0].length > 0 ? { ...content, content: parts[0] } : null,
      parts[1].length > 0 ? { ...content, content: parts[1] } : null,
    ];
  }
  if (content.type === 'hyperlink') {
    const children: readonly HyperlinkContent[] = content.structuredChildren ?? content.children;
    const parts = splitItems(children, at, linkChildLength, (child, offset) =>
      child.type === 'run' ? (splitContent(child, offset) as Split<Run> | null) : null
    );
    return parts && [linkPart(content, parts[0]), linkPart(content, parts[1])];
  }
  if (isTrackedWrapper(content)) {
    const parts = splitItems(
      content.content,
      at,
      paragraphContentLength,
      (child, offset) => splitContent(child, offset) as Split<TrackedWrapper['content'][number]> | null
    );
    return parts && [
      parts[0].length > 0 ? { ...content, content: parts[0] } : null,
      parts[1].length > 0 ? { ...content, content: parts[1] } : null,
    ];
  }
  return null;
}

/**
 * Inserts zero-width markers at story offsets into paragraph content,
 * splitting runs, hyperlinks and tracked changes where one falls inside
 * them. Markers inside content that cannot be split go around it: starts
 * before, ends after.
 */
function insertBoundaries(
  content: ParagraphContent[],
  boundaries: CommentBoundary[],
  makeMarker: (boundary: CommentBoundary) => ParagraphContent = (boundary) =>
    boundary.kind === 'start'
      ? { type: 'commentRangeStart', id: boundary.id }
      : { type: 'commentRangeEnd', id: boundary.id }
): ParagraphContent[] {
  if (boundaries.length === 0) return content;
  const carets = new Set(
    boundaries
      .filter(
        (boundary) =>
          boundary.kind === 'end' &&
          boundaries.some(
            (other) =>
              other.kind === 'start' && other.id === boundary.id && other.offset === boundary.offset
          )
      )
      .map((boundary) => `${boundary.offset}:${boundary.id}`)
  );
  // At one offset: ranges closing, then carets opening and closing, then ranges opening.
  const rank = (boundary: CommentBoundary): number =>
    carets.has(`${boundary.offset}:${boundary.id}`) ? 1 : boundary.kind === 'end' ? 0 : 2;
  const sorted = [...boundaries].sort(
    (left, right) =>
      left.offset - right.offset ||
      rank(left) - rank(right) ||
      left.id - right.id ||
      (left.kind === right.kind ? 0 : left.kind === 'start' ? -1 : 1)
  );
  const result: ParagraphContent[] = [];
  let cursor = 0;
  let boundaryIndex = 0;
  const emit = (offset: number): void => {
    while (boundaryIndex < sorted.length && sorted[boundaryIndex].offset <= offset) {
      result.push(makeMarker(sorted[boundaryIndex++]));
    }
  };

  emit(0);
  for (const child of content) {
    const end = cursor + paragraphContentLength(child);
    let item: ParagraphContent | null = child;
    while (item && boundaryIndex < sorted.length && sorted[boundaryIndex].offset < end) {
      const offset = sorted[boundaryIndex].offset;
      const parts = splitContent(item, offset - cursor);
      if (!parts) break;
      if (parts[0]) result.push(parts[0]);
      cursor = offset;
      emit(cursor);
      item = parts[1];
    }
    const inside: CommentBoundary[] = [];
    while (boundaryIndex < sorted.length && sorted[boundaryIndex].offset < end) {
      inside.push(sorted[boundaryIndex++]);
    }
    for (const boundary of inside) if (boundary.kind === 'start') result.push(makeMarker(boundary));
    if (item) result.push(item);
    for (const boundary of inside) if (boundary.kind === 'end') result.push(makeMarker(boundary));
    cursor = end;
    emit(cursor);
  }
  emit(Number.POSITIVE_INFINITY);
  return result;
}

function bookmarkBoundaries(properties: Attrs): BookmarkBoundary[] {
  if (!Array.isArray(properties.bookmarks)) return [];
  const result: CommentBoundary[] = [];
  for (const raw of properties.bookmarks) {
    const bookmark = asObject(raw);
    const id = asFiniteNumber(bookmark?.id);
    if (!bookmark || id === undefined) continue;
    const offset = asFiniteNumber(bookmark.offset) ?? 0;
    const metadata = {
      name: asString(bookmark.name),
      colFirst: asFiniteNumber(bookmark.colFirst),
      colLast: asFiniteNumber(bookmark.colLast),
    };
    if (bookmark.kind === 'start') result.push({ id, kind: 'start', offset, ...metadata });
    else if (bookmark.kind === 'end') result.push({ id, kind: 'end', offset, ...metadata });
    else {
      result.push({ id, kind: 'start', offset: 0, ...metadata });
      result.push({ id, kind: 'end', offset: Number.MAX_SAFE_INTEGER, ...metadata });
    }
  }
  return result;
}

/** Paragraphs nested in inline content, such as text-box paragraphs, in document order. */
function nestedParagraphs(value: unknown, found: Record<string, unknown>[] = []): Record<string, unknown>[] {
  if (Array.isArray(value)) {
    for (const item of value) nestedParagraphs(item, found);
  } else if (value !== null && typeof value === 'object') {
    const record = value as Record<string, unknown>;
    if (record.type === 'paragraph') found.push(record);
    for (const child of Object.values(record)) nestedParagraphs(child, found);
  }
  return found;
}

/**
 * Gives nested paragraphs the source occurrences their base counterparts
 * carry, so a save applies IDs prepared for them; only when the paragraph
 * still holds the same number of nested paragraphs.
 */
function restoreNestedOrdinals(content: ParagraphContent[], base: readonly ParagraphContent[]): void {
  const sources = nestedParagraphs(base);
  if (!sources.some((paragraph) => typeof paragraph.sourceOrdinal === 'number')) return;
  const targets = nestedParagraphs(content);
  if (targets.length !== sources.length) return;
  targets.forEach((target, index) => {
    const ordinal = sources[index]!.sourceOrdinal;
    if (typeof ordinal === 'number') target.sourceOrdinal = ordinal;
  });
}

function restoreRawInlines(content: ParagraphContent[], base: Paragraph | undefined): ParagraphContent[] {
  if (!base) return content;
  const length = content.reduce((sum, child) => sum + paragraphContentLength(child), 0);
  const boundaries: CommentBoundary[] = [];
  let offset = 0;
  for (const [index, child] of base.content.entries()) {
    if (child.type === 'rawXml') boundaries.push({ id: index, kind: 'start', offset: Math.min(offset, length) });
    offset += paragraphContentLength(child);
  }
  return insertBoundaries(content, boundaries, (boundary) => base.content[boundary.id]!);
}

function paragraphAttrs(properties: Attrs): ParagraphSaveAttrs {
  const attrs = { ...PARAGRAPH_ATTR_DEFAULTS, ...properties } as Attrs;
  attrs.styleId = properties.pStyle ?? null;
  attrs._sectionProperties = properties.sectPr ?? null;
  delete attrs.pStyle;
  delete attrs.sectPr;
  return attrs as ParagraphSaveAttrs;
}

function sourceRuns(content: readonly ParagraphContent[]): OriginalRunBoundary[] {
  return content.flatMap((child): OriginalRunBoundary[] => {
    if (child.type === 'run') {
      return [{
        text: child.content.map((entry) => {
          if (entry.type === 'text' || entry.type === 'instrText') return entry.text;
          if (entry.type === 'symbol') return entry.char;
          if (entry.type === 'tab') return '\t';
          if (entry.type === 'softHyphen') return '\u00ad';
          if (entry.type === 'noBreakHyphen') return '\u2011';
          return '\ufffc'.repeat(runContentLength(entry));
        }).join(''),
        formatting: child.formatting,
      }];
    }
    if (child.type === 'hyperlink') return sourceRuns(child.children as ParagraphContent[]);
    if (isTrackedWrapper(child) || child.type === 'inlineSdt') {
      return sourceRuns(child.content as ParagraphContent[]);
    }
    return [{ text: '\ufffc'.repeat(paragraphContentLength(child)) }];
  });
}

/** A midpoint of the shortest edit path, in UTF-16 units. */
function editMidpoint(original: string, current: string): [number, number] | null {
  const distanceLimit = Math.ceil((original.length + current.length) / 2);
  const offset = distanceLimit + 1;
  const forward = new Int32Array(2 * distanceLimit + 3).fill(-1);
  const backward = new Int32Array(forward.length).fill(-1);
  forward[offset + 1] = 0;
  backward[offset + 1] = 0;
  const delta = original.length - current.length;
  const odd = delta % 2 !== 0;
  let forwardStart = 0;
  let forwardEnd = 0;
  let backwardStart = 0;
  let backwardEnd = 0;
  for (let distance = 0; distance <= distanceLimit; distance += 1) {
    for (let diagonal = -distance + forwardStart; diagonal <= distance - forwardEnd; diagonal += 2) {
      const index = offset + diagonal;
      let x = diagonal === -distance ||
        (diagonal !== distance && forward[index - 1]! < forward[index + 1]!)
        ? forward[index + 1]! : forward[index - 1]! + 1;
      let y = x - diagonal;
      while (x < original.length && y < current.length && original[x] === current[y]) {
        x += 1;
        y += 1;
      }
      forward[index] = x;
      if (x > original.length) forwardEnd += 2;
      else if (y > current.length) forwardStart += 2;
      else if (odd && backward[offset + delta - diagonal]! >= 0 &&
        x >= original.length - backward[offset + delta - diagonal]!) return [x, y];
    }
    for (let diagonal = -distance + backwardStart; diagonal <= distance - backwardEnd; diagonal += 2) {
      const index = offset + diagonal;
      let x = diagonal === -distance ||
        (diagonal !== distance && backward[index - 1]! < backward[index + 1]!)
        ? backward[index + 1]! : backward[index - 1]! + 1;
      let y = x - diagonal;
      while (x < original.length && y < current.length &&
        original[original.length - x - 1] === current[current.length - y - 1]) {
        x += 1;
        y += 1;
      }
      backward[index] = x;
      if (x > original.length) backwardEnd += 2;
      else if (y > current.length) backwardStart += 2;
      else if (!odd && forward[offset + delta - diagonal]! >= 0 &&
        forward[offset + delta - diagonal]! >= original.length - x) {
        const start = forward[offset + delta - diagonal]!;
        return [start, start - delta + diagonal];
      }
    }
  }
  return null;
}

interface SourceTextMatch {
  original: number;
  current: number;
  length: number;
}

function sourceTextMatches(original: string, current: string): SourceTextMatch[] {
  const matches: SourceTextMatch[] = [];
  const visit = (before: string, after: string, sourceOffset: number, liveOffset: number): void => {
    let prefix = 0;
    while (prefix < Math.min(before.length, after.length) && before[prefix] === after[prefix]) prefix += 1;
    if (prefix > 0) matches.push({ original: sourceOffset, current: liveOffset, length: prefix });
    let suffix = 0;
    while (suffix < Math.min(before.length, after.length) - prefix &&
      before[before.length - suffix - 1] === after[after.length - suffix - 1]) suffix += 1;
    const oldMiddle = before.slice(prefix, before.length - suffix);
    const newMiddle = after.slice(prefix, after.length - suffix);
    if (oldMiddle.length > 0 && newMiddle.length > 0) {
      const split = editMidpoint(oldMiddle, newMiddle);
      if (split && (split[0] > 0 || split[1] > 0) &&
        (split[0] < oldMiddle.length || split[1] < newMiddle.length)) {
        visit(oldMiddle.slice(0, split[0]), newMiddle.slice(0, split[1]), sourceOffset + prefix, liveOffset + prefix);
        visit(oldMiddle.slice(split[0]), newMiddle.slice(split[1]), sourceOffset + prefix + split[0], liveOffset + prefix + split[1]);
      }
    }
    if (suffix > 0) matches.push({
      original: sourceOffset + before.length - suffix,
      current: liveOffset + after.length - suffix,
      length: suffix,
    });
  };
  visit(original, current, 0, 0);
  return matches;
}

function preserveSourceRunFormatting(items: InlineItem[], sources: OriginalRunBoundary[]): InlineItem[] {
  if (!sources.some((source) => SOURCE_RUN_PROPERTIES.some((key) => source.formatting?.[key] !== undefined))) {
    return items;
  }
  const original = sources.map((source) => source.text).join('');
  const current = items.map((item) => item.kind === 'text' ? item.text
    : item.embedKind === 'tab' ? '\t' : '\ufffc').join('');
  const matches = sourceTextMatches(original, current);
  const mapped = (offset: number): number => {
    if (offset === 0) return 0;
    let sourceEnd = 0;
    let liveEnd = 0;
    for (const match of matches) {
      if (offset < match.original) return offset === sourceEnd ? liveEnd : match.current;
      if (offset <= match.original + match.length) return match.current + offset - match.original;
      sourceEnd = match.original + match.length;
      liveEnd = match.current + match.length;
    }
    return offset === sourceEnd ? liveEnd : current.length;
  };
  let sourceOffset = 0;
  const spans = sources.map((source, index) => {
    const start = mapped(sourceOffset);
    sourceOffset += source.text.length;
    return { start, end: index + 1 === sources.length ? current.length : mapped(sourceOffset), formatting: source.formatting };
  }).filter((span) => span.end > span.start);
  let offset = 0;
  return items.flatMap((item): InlineItem[] => {
    const end = offset + (item.kind === 'text' ? item.text.length : 1);
    const slices: InlineItem[] = [];
    while (offset < end) {
      const span = spans.find((span) => span.start <= offset && offset < span.end);
      const next = Math.min(end, span?.end ?? end);
      const source = Object.fromEntries(SOURCE_RUN_PROPERTIES
        .filter((key) => span?.formatting?.[key] !== undefined)
        .map((key) => [key, span!.formatting![key]]));
      const attributes = { ...item.attributes, _sourceRunFormatting: source };
      slices.push(item.kind === 'text'
        ? { ...item, text: item.text.slice(item.text.length - (end - offset), item.text.length - (end - next)), attributes }
        : { ...item, attributes });
      offset = next;
    }
    return slices;
  });
}

function paragraphFromStory(
  paraId: string,
  properties: Attrs,
  items: InlineItem[],
  commentBoundaries: CommentBoundary[],
  baseParagraph: Paragraph | undefined,
  liveBreaks: LiveParagraphBreaks,
  recovery: ChartRecovery,
  format: RunFormatter,
  revisionIds?: RevisionIds,
  inherited?: ParagraphFormatting
): Paragraph {
  const attrs = paragraphAttrs(properties);
  const sources = baseParagraph ? sourceRuns(baseParagraph.content)
    : Array.isArray(attrs._originalRunBoundaries) ? attrs._originalRunBoundaries as OriginalRunBoundary[] : [];
  let content = buildParagraphContent(preserveSourceRunFormatting(items, sources), recovery, format, revisionIds);
  content = restoreOriginalRuns(
    content,
    items,
    Array.isArray(attrs._originalRunBoundaries)
      ? (attrs._originalRunBoundaries as OriginalRunBoundary[])
      : undefined,
    baseParagraph,
    liveBreaks
  );
  content = restoreRawInlines(content, baseParagraph);
  if (baseParagraph) restoreNestedOrdinals(content, baseParagraph.content);
  content = insertBoundaries(content, commentBoundaries);

  const bookmarks = bookmarkBoundaries(properties).map((boundary) => ({
    ...boundary,
    offset:
      boundary.offset === Number.MAX_SAFE_INTEGER
        ? content.reduce((sum, child) => sum + paragraphContentLength(child), 0)
        : boundary.offset,
  }));
  if (bookmarks.length > 0) {
    content = insertBoundaries(content, bookmarks, (rawBoundary) => {
      const boundary = rawBoundary as BookmarkBoundary;
      return boundary.kind === 'start'
        ? {
            type: 'bookmarkStart',
            id: boundary.id,
            name: boundary.name || '',
            ...(boundary.colFirst !== undefined ? { colFirst: boundary.colFirst } : {}),
            ...(boundary.colLast !== undefined ? { colLast: boundary.colLast } : {}),
            position: { offset: boundary.offset },
          }
        : { type: 'bookmarkEnd', id: boundary.id, position: { offset: boundary.offset } };
    });
  }

  const paragraph: Paragraph = {
    type: 'paragraph',
    paraId: paraId || undefined,
    textId: baseParagraph?.textId,
    ...(baseParagraph?.extraAttributes ? { extraAttributes: baseParagraph.extraAttributes } : {}),
    ...(baseParagraph?.paraIdAttribute ? { paraIdAttribute: baseParagraph.paraIdAttribute } : {}),
    formatting: paragraphAttrsToFormatting(attrs, inherited),
    content,
  };
  if (baseParagraph?.renderedPageBreakBefore) paragraph.renderedPageBreakBefore = true;

  const pPrIns = trackedInfo(properties.pPrIns, true);
  const pPrDel = trackedInfo(properties.pPrDel, true);
  if (pPrIns) paragraph.pPrIns = pPrIns;
  if (pPrDel) paragraph.pPrDel = pPrDel;
  if (Array.isArray(properties.pPrChange) && properties.pPrChange.length > 0) {
    paragraph.propertyChanges = propertyChanges(properties.pPrChange) as Paragraph['propertyChanges'];
  }
  if (properties.sectPr) {
    paragraph.sectionProperties = properties.sectPr as Paragraph['sectionProperties'];
  } else if (properties.sectionBreakType) {
    paragraph.sectionProperties = {
      sectionStart: properties.sectionBreakType as NonNullable<
        Paragraph['sectionProperties']
      >['sectionStart'],
    };
  }
  return paragraph;
}

/**
 * Lifts a `w:tblBorders` back out of the per-cell edges seeding pushed down:
 * outer sides come from the cells owning the table's boundary, `insideH`/
 * `insideV` from an interior edge. Sides no cell authors stay absent.
 */
function inferTableBorders(rows: TableRow[]): TableBorders | undefined {
  const firstRow = rows[0]?.cells;
  const lastRow = rows[rows.length - 1]?.cells;
  const corner = firstRow?.[0]?.formatting?.borders;
  if (!firstRow || !lastRow || !corner) return undefined;
  const borders: TableBorders = {
    top: corner.top,
    left: corner.left,
    bottom: lastRow[0]?.formatting?.borders?.bottom,
    right: firstRow[firstRow.length - 1]?.formatting?.borders?.right,
    insideH: rows.length > 1 ? corner.bottom : undefined,
    insideV: firstRow.length > 1 ? corner.right : undefined,
  };
  const authored = Object.entries(borders).filter(([, value]) => value !== undefined);
  return authored.length > 0 ? (Object.fromEntries(authored) as TableBorders) : undefined;
}

function normalizeVMergeRuns(rows: TableRow[]): void {
  const columns = new Map<number, Array<{ rowIndex: number; cell: TableCell }>>();
  rows.forEach((row, rowIndex) => {
    let column = 0;
    for (const cell of row.cells) {
      const start = column;
      column += cell.formatting?.gridSpan ?? 1;
      if (cell.formatting?.vMerge) {
        const entries = columns.get(start) ?? [];
        entries.push({ rowIndex, cell });
        columns.set(start, entries);
      }
    }
  });

  const clear = (cell: TableCell): void => {
    if (!cell.formatting) return;
    delete cell.formatting.vMerge;
    if (Object.keys(cell.formatting).length === 0) cell.formatting = undefined;
  };

  for (const entries of columns.values()) {
    let start: TableCell | null = null;
    let length = 0;
    let lastRow = -1;
    const close = (): void => {
      if (start && length < 2) clear(start);
      start = null;
      length = 0;
    };
    for (const entry of entries) {
      const marker = entry.cell.formatting?.vMerge;
      if (marker === 'restart') {
        close();
        start = entry.cell;
        length = 1;
        lastRow = entry.rowIndex;
      } else if (marker === 'continue') {
        if (start && entry.rowIndex === lastRow + 1) {
          length += 1;
          lastRow = entry.rowIndex;
        } else {
          close();
          clear(entry.cell);
        }
      }
    }
    close();
  }
}

function tableCellFromPayload(context: SaveContext, payload: TableCellPayload): TableCell {
  const attrs = context.cellAttrs(payload);
  const content =
    payload.story && context.storyIds.has(payload.story)
      ? context.storyToBlocks(payload.story)
      : [];
  const cell: TableCell = {
    type: 'tableCell',
    formatting: tableCellAttrsToFormatting(attrs),
    content,
  };
  const marker = attrs.cellMarker;
  if (marker) {
    const info = trackedInfo(marker.info, true) ?? { id: 0, author: 'Unknown' };
    if (marker.kind === 'ins') cell.structuralChange = { type: 'tableCellInsertion', info };
    else if (marker.kind === 'del') cell.structuralChange = { type: 'tableCellDeletion', info };
    else {
      cell.structuralChange = {
        type: 'tableCellMerge',
        info,
        ...(marker.vMerge ? { vMerge: marker.vMerge } : {}),
        ...(marker.vMergeOrig ? { vMergeOrig: marker.vMergeOrig } : {}),
      };
    }
  }
  if (Array.isArray(attrs.tcPrChange) && attrs.tcPrChange.length > 0) {
    cell.propertyChanges = attrs.tcPrChange;
  }
  return cell;
}

function tableFromPayload(context: SaveContext, payload: TablePayload): Table {
  const rowPayloads = Array.isArray(payload.rows) ? payload.rows : [];
  const occupied: boolean[][] = [];
  const anchors: Array<{
    row: number;
    col: number;
    rowspan: number;
    colspan: number;
    cell: TableCell;
  }> = [];
  let totalColumns = 0;

  rowPayloads.forEach((row, rowIndex) => {
    let column = 0;
    for (const cellPayload of Array.isArray(row.cells) ? row.cells : []) {
      while (occupied[rowIndex]?.[column]) column += 1;
      const tcPr = cellPayload.tcPr ?? {};
      const rowspan = asFiniteNumber(tcPr.rowspan) || 1;
      const colspan = asFiniteNumber(tcPr.colspan) || 1;
      anchors.push({
        row: rowIndex,
        col: column,
        rowspan,
        colspan,
        cell: tableCellFromPayload(context, cellPayload),
      });
      for (let r = rowIndex; r < rowIndex + rowspan; r += 1) {
        occupied[r] ??= [];
        for (let c = column; c < column + colspan; c += 1) occupied[r][c] = true;
      }
      column += colspan;
      totalColumns = Math.max(totalColumns, column);
    }
  });

  const byStart = new Map(anchors.map((anchor) => [`${anchor.row}-${anchor.col}`, anchor]));
  const byCovered = new Map<string, (typeof anchors)[number]>();
  for (const anchor of anchors) {
    for (let row = anchor.row; row < anchor.row + anchor.rowspan; row += 1) {
      for (let col = anchor.col; col < anchor.col + anchor.colspan; col += 1) {
        byCovered.set(`${row}-${col}`, anchor);
      }
    }
  }

  const rows: TableRow[] = rowPayloads.map((rowPayload, rowIndex) => {
    const cells: TableCell[] = [];
    for (let col = 0; col < totalColumns; ) {
      const anchor = byStart.get(`${rowIndex}-${col}`);
      if (anchor) {
        const formatting = { ...(anchor.cell.formatting ?? {}) };
        if (anchor.colspan > 1) formatting.gridSpan = anchor.colspan;
        else delete formatting.gridSpan;
        if (anchor.rowspan > 1) formatting.vMerge = 'restart';
        else if (formatting.vMerge !== 'restart' && formatting.vMerge !== 'continue') {
          delete formatting.vMerge;
        }
        cells.push({
          ...anchor.cell,
          formatting: Object.keys(formatting).length > 0 ? formatting : undefined,
        });
        col += anchor.colspan;
        continue;
      }
      const covering = byCovered.get(`${rowIndex}-${col}`);
      if (!covering) {
        col += 1;
        continue;
      }
      const formatting = { ...(covering.cell.formatting ?? {}) };
      if (covering.colspan > 1) formatting.gridSpan = covering.colspan;
      else delete formatting.gridSpan;
      formatting.vMerge = 'continue';
      cells.push({ ...covering.cell, content: [], formatting });
      col += covering.colspan;
    }

    const attrs = { ...TABLE_ROW_ATTR_DEFAULTS, ...(rowPayload.trPr ?? {}) } as TableRowSaveAttrs;
    const row: TableRow = {
      type: 'tableRow',
      formatting: tableRowAttrsToFormatting(attrs),
      cells,
    };
    const ins = trackedInfo(attrs.trIns, true);
    const del = trackedInfo(attrs.trDel, true);
    if (ins) row.structuralChange = { type: 'tableRowInsertion', info: ins };
    else if (del) row.structuralChange = { type: 'tableRowDeletion', info: del };
    if (Array.isArray(attrs.trPrChange) && attrs.trPrChange.length > 0) {
      row.propertyChanges = attrs.trPrChange;
    }
    return row;
  });

  normalizeVMergeRuns(rows);
  const grid = Array.isArray(payload.grid)
    ? payload.grid.filter((width): width is number => typeof width === 'number')
    : [];
  const attrs = {
    ...TABLE_ATTR_DEFAULTS,
    ...(payload.tblPr ?? {}),
    columnWidths: grid.length > 0 ? grid : undefined,
  } as TableSaveAttrs;
  context.applyTableBaseline(payload, attrs);
  let formatting = tableAttrsToFormatting(attrs);
  if (!formatting?.borders && !attrs._originalFormatting) {
    const borders = inferTableBorders(rows);
    if (borders) formatting = { ...(formatting ?? {}), borders };
  }
  const table: Table = {
    type: 'table',
    columnWidths: attrs.columnWidths || undefined,
    formatting,
    rows,
  };
  if (Array.isArray(attrs.tblPrChange) && attrs.tblPrChange.length > 0) {
    table.propertyChanges = attrs.tblPrChange;
  }
  return table;
}

function pageBreakParagraph(): Paragraph {
  return {
    type: 'paragraph',
    content: [{ type: 'run', content: [{ type: 'break', breakType: 'page' }] }],
  };
}

/** A note number mark run, or a tracked-change wrapper holding only those. */
function isNoteMark(content: ParagraphContent): boolean {
  if (
    content.type === 'insertion' ||
    content.type === 'deletion' ||
    content.type === 'moveFrom' ||
    content.type === 'moveTo'
  ) {
    return content.content.length > 0 && content.content.every(isNoteMark);
  }
  return (
    content.type === 'run' &&
    content.content.length > 0 &&
    content.content.every(
      (entry) => entry.type === 'footnoteRefMark' || entry.type === 'endnoteRefMark'
    )
  );
}

function hasNoteMark(blocks: readonly BlockContent[]): boolean {
  return blocks.some((block) => block.type === 'paragraph' && block.content.some(isNoteMark));
}

/**
 * Reinstates the `w:footnoteRef` / `w:endnoteRef` number mark on a projected
 * note. The mark carries no story unit, so an edit that invalidates the run
 * boundary cache would otherwise drop it; the source note is the only record.
 */
function restoreNoteMarks(
  projected: BlockContent[],
  base: readonly BlockContent[]
): BlockContent[] {
  if (hasNoteMark(projected)) return projected;
  const opening = base.find((block) => block.type === 'paragraph')?.content ?? [];
  const end = opening.findIndex((child) => !isNoteMark(child));
  const marks = opening.slice(0, end < 0 ? opening.length : end);
  if (marks.length === 0) return projected;
  const index = projected.findIndex((block) => block.type === 'paragraph');
  if (index < 0) return projected;
  const target = projected[index] as Paragraph;
  const restored = [...projected];
  restored[index] = cloneProjectedParagraph(target, { content: [...marks, ...target.content] });
  return restored;
}

/**
 * Base paragraphs by session key: the key a projection recorded, else the key
 * seeding gives a source paragraph (its `paraId` unless an earlier paragraph
 * holds it, else its story position).
 */
function collectBaseParagraphs(
  stories: ReadonlyMap<string, readonly BlockContent[]>
): Map<string, Paragraph> {
  const paragraphs = new Map<string, Paragraph>();
  for (const [storyId, blocks] of stories) {
    let index = 0;
    for (const block of blocks) {
      if (block.type !== 'paragraph') continue;
      const key =
        projectedBlocks.get(block)?.sessionKey ??
        ((!block.repeatedParaId && block.paraId) || `${storyId}:p${index}`);
      index += 1;
      if (!paragraphs.has(key)) paragraphs.set(key, block);
    }
  }
  return paragraphs;
}

function collectBaseStories(
  session: YrsSession,
  document: Document
): Map<string, readonly BlockContent[]> {
  const stories = new Map<string, readonly BlockContent[]>();
  const visit = (storyId: string, blocks: readonly BlockContent[]): void => {
    if (stories.has(storyId)) return;
    stories.set(storyId, blocks);
    let tableIndex = 0;
    let sdtIndex = 0;
    for (const block of blocks) {
      if (block.type === 'blockSdt') {
        visit(`${storyId}:sdt${sdtIndex++}`, block.content);
        continue;
      }
      if (block.type !== 'table') continue;
      const currentTableIndex = tableIndex++;
      block.rows.forEach((row, rowIndex) => {
        row.cells.forEach((cell, cellIndex) => {
          visit(`${storyId}:t${currentTableIndex}:r${rowIndex}c${cellIndex}`, cell.content);
        });
      });
    }
  };

  visit('body', document.package.document.content);
  for (const [rId, part] of document.package.headers ?? []) {
    visit(headerFooterStory(session, rId), part.content);
  }
  for (const [rId, part] of document.package.footers ?? []) {
    visit(headerFooterStory(session, rId), part.content);
  }
  for (const note of document.package.footnotes ?? []) visit(`fn:${note.id}`, note.content);
  for (const note of document.package.endnotes ?? []) visit(`en:${note.id}`, note.content);
  return stories;
}

function commentRanges(
  session: YrsSession,
  comments: readonly Comment[] | undefined,
  sessionIds: ReadonlyMap<number, string> | undefined
): Map<string, Array<{ id: number; start: number; end: number }>> {
  const byStory = new Map<string, Array<{ id: number; start: number; end: number }>>();
  for (const comment of comments ?? []) {
    let anchors: ReturnType<YrsSession['resolveComment']>;
    try {
      anchors = session.resolveComment(sessionIds?.get(comment.id) ?? String(comment.id));
    } catch {
      continue;
    }
    const storyGroups = new Map<string, typeof anchors>();
    for (const anchor of anchors) {
      const group = storyGroups.get(anchor.story) ?? [];
      group.push(anchor);
      storyGroups.set(anchor.story, group);
    }
    for (const [story, group] of storyGroups) {
      if (group.length === 0) continue;
      const start = Math.min(...group.map((anchor) => anchor.start));
      const end = Math.max(...group.map((anchor) => anchor.end));
      const ranges = byStory.get(story) ?? [];
      ranges.push({ id: comment.id, start, end });
      byStory.set(story, ranges);
    }
  }
  return byStory;
}

/** Chart placements recoverable from the base document, keyed per story. */
interface ChartRecovery {
  drawingFor(rId: string | undefined, path: string | undefined): string | undefined;
  unrecoverable(rId: string | undefined, path: string | undefined): void;
  warn(message: string): void;
}

function collectChartDrawings(blocks: readonly BlockContent[], into: Map<string, string>): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const run of paragraphChartRuns(block.content)) {
        for (const content of run.content) {
          if (content.type !== 'chart' || !content.chart.drawingXml) continue;
          if (content.chart.rId !== undefined) {
            const key = `rId:${content.chart.rId}`;
            if (!into.has(key)) into.set(key, content.chart.drawingXml);
          }
          if (content.chart.path !== undefined) {
            const key = `path:${content.chart.path}`;
            if (!into.has(key)) into.set(key, content.chart.drawingXml);
          }
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) collectChartDrawings(cell.content, into);
    } else if (block.type === 'blockSdt') {
      collectChartDrawings(block.content, into);
    }
  }
}

function paragraphChartRuns(content: readonly ParagraphContent[]): Run[] {
  const runs: Run[] = [];
  const visitInline = (child: ParagraphContent | HyperlinkContent): void => {
    if (child.type === 'run') runs.push(child);
    else if (child.type === 'hyperlink') {
      for (const inline of child.structuredChildren ?? child.children) visitInline(inline);
    } else if (child.type === 'inlineSdt') {
      for (const inline of child.content) visitInline(inline);
    } else if (child.type === 'simpleField') {
      for (const run of child.content) if (run.type === 'run') runs.push(run);
    } else if (child.type === 'complexField') {
      for (const run of [...child.fieldCode, ...child.fieldResult]) runs.push(run);
    } else if (
      child.type === 'insertion' || child.type === 'deletion' ||
      child.type === 'moveFrom' || child.type === 'moveTo'
    ) {
      for (const inline of child.content) visitInline(inline as ParagraphContent);
    }
  };
  for (const child of content) visitInline(child);
  return runs;
}

function isBlockBookmark(block: BlockContent): boolean {
  return isRawXml(block) && /^<w:bookmark(?:Start|End)(?=[\s/>])/.test(block.xml);
}

type BookmarkKind = 'start' | 'end';

function rawBlockBookmark(xml: string): { kind: BookmarkKind; id?: string } | undefined {
  const kind = /^<w:bookmark(Start|End)(?=[\s/>])/.exec(xml)?.[1];
  if (!kind) return undefined;
  let value: string | undefined;
  for (const attribute of xml.matchAll(/([^\s=<>/]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
    if (attribute[1] !== 'w:id') continue;
    value = (attribute[2] ?? attribute[3]!).replace(/&#(?:x([0-9a-f]+)|([0-9]+));/gi, (entity, hex, decimal) => {
      const code = Number.parseInt(hex ?? decimal, hex === undefined ? 10 : 16);
      return code <= 0x10ffff ? String.fromCodePoint(code) : entity;
    });
    break;
  }
  const id = value === undefined ? undefined
    : Number.isSafeInteger(Number(value)) ? String(Number(value)) : value;
  return { kind: kind === 'Start' ? 'start' : 'end', id };
}

function storyBlockKeys(blocks: readonly BlockContent[], storyId: string): Array<string | undefined> {
  let paragraph = 0;
  let table = 0;
  let sdt = 0;
  return blocks.map((block) => {
    const memo = projectedBlocks.get(block);
    if (block.type === 'paragraph') {
      const key = memo?.sessionKey ?? ((!block.repeatedParaId && block.paraId) || `${storyId}:p${paragraph}`);
      paragraph += 1;
      return `P${key}`;
    }
    if (block.type === 'table') {
      const key = memo?.tableSourcePrefix ?? `${storyId}:t${table}`;
      table += 1;
      return `T${key}`;
    }
    if (block.type === 'blockSdt') {
      const key = memo?.key ?? `S${storyId}:sdt${sdt}`;
      sdt += 1;
      return key;
    }
    return undefined;
  });
}

function projectedBookmarkPositions(blocks: readonly BlockContent[]): Map<string, Partial<Record<BookmarkKind, number>>> {
  const positions = new Map<string, Partial<Record<BookmarkKind, number>>>();
  const visited = new WeakSet<object>();
  const visit = (value: unknown, position: number): void => {
    if (value === null || typeof value !== 'object' || visited.has(value)) return;
    visited.add(value);
    if (Array.isArray(value)) {
      for (const child of value) visit(child, position);
      return;
    }
    const record = value as Record<string, unknown>;
    const marker = record.type === 'rawXml' && typeof record.xml === 'string'
      ? rawBlockBookmark(record.xml)
      : record.type === 'bookmarkStart' || record.type === 'bookmarkEnd'
        ? { kind: record.type === 'bookmarkStart' ? 'start' as const : 'end' as const, id: String(record.id) }
        : undefined;
    if (marker?.id !== undefined) {
      let pair = positions.get(marker.id);
      if (!pair) positions.set(marker.id, pair = {});
      pair[marker.kind] ??= position;
      return;
    }
    for (const child of Object.values(record)) visit(child, position);
  };
  blocks.forEach((block, index) => visit(block, index + 0.5));
  return positions;
}

interface BlockBookmarkPlacement {
  block: BlockContent;
  index: number;
  position: number;
  omit?: boolean;
  following?: BlockContent;
}

function restoreBlockBookmarks(
  projected: BlockContent[],
  base: readonly BlockContent[],
  storyId: string
): BlockContent[] {
  if (!base.some(isBlockBookmark)) return projected;
  const keys = storyBlockKeys(base, storyId);
  const live = new Map<string, number>();
  storyBlockKeys(projected, storyId).forEach((key, index) => {
    if (key !== undefined) live.set(key, index);
  });
  const nextPositions = new Array<number | undefined>(base.length);
  let next: number | undefined;
  for (let index = base.length - 1; index >= 0; index -= 1) {
    nextPositions[index] = next;
    const key = keys[index];
    const position = key === undefined ? undefined : live.get(key);
    if (position !== undefined) next = position;
  }
  const present = projectedBookmarkPositions(projected);
  const pairs = new Map<string, Partial<Record<BookmarkKind, BlockBookmarkPlacement>>>();
  const boundaries: BlockBookmarkPlacement[] = [];
  let previous: number | undefined;
  for (let index = 0; index < base.length; index += 1) {
    const block = base[index]!;
    const marker = isRawXml(block) ? rawBlockBookmark(block.xml) : undefined;
    if (marker) {
      const id = marker.id ?? `\0${index}`;
      let pair = pairs.get(id);
      if (!pair) pairs.set(id, pair = {});
      if (!pair[marker.kind]) {
        const boundary = {
          block, index,
          position: nextPositions[index] ?? (previous === undefined ? 0 : previous + 1),
          omit: present.get(id)?.[marker.kind] !== undefined,
        };
        pair[marker.kind] = boundary;
        boundaries.push(boundary);
      }
    }
    const key = keys[index];
    const position = key === undefined ? undefined : live.get(key);
    if (position !== undefined) previous = position;
  }
  for (const [id, pair] of pairs) {
    const { start, end } = pair;
    const existing = present.get(id);
    if (start && !start.omit && existing?.end !== undefined) {
      start.position = Math.min(start.position, Math.floor(existing.end));
    }
    if (end && !end.omit && existing?.start !== undefined) {
      end.position = Math.max(end.position, Math.ceil(existing.start));
    }
    if (start && end && !start.omit && !end.omit && (
      end.position < start.position || (end.position === start.position && end.index < start.index)
    )) {
      start.following = end.block;
      end.omit = true;
    }
  }
  const insertions = new Map<number, BlockContent[]>();
  for (const boundary of boundaries) {
    if (boundary.omit) continue;
    let bucket = insertions.get(boundary.position);
    if (!bucket) insertions.set(boundary.position, bucket = []);
    bucket.push(boundary.block);
    if (boundary.following) bucket.push(boundary.following);
  }
  const restored: BlockContent[] = [];
  for (let index = 0; index <= projected.length; index += 1) {
    for (const block of insertions.get(index) ?? []) restored.push(block);
    if (index < projected.length) restored.push(projected[index]!);
  }
  return restored;
}

function restoreRawBlocks(projected: BlockContent[], base: readonly BlockContent[], storyId: string): BlockContent[] {
  projected = restoreBlockBookmarks(projected, base, storyId);
  let offset = 0;
  for (let index = 0; index < base.length; index += 1) {
    const block = base[index]!;
    if (!isRawXml(block) || isBlockBookmark(block)) continue;
    const following = base.slice(index + 1).find((candidate) =>
      candidate.type === 'paragraph' && candidate.paraId && projected.some((entry) =>
        entry.type === 'paragraph' && entry.paraId === candidate.paraId));
    const anchor = following?.type === 'paragraph'
      ? projected.findIndex((entry) => entry.type === 'paragraph' && entry.paraId === following.paraId)
      : -1;
    const position = anchor >= 0 ? anchor : Math.min(index + offset, projected.length);
    projected.splice(position, 0, block);
    offset = Math.max(0, position - index);
  }
  return projected;
}

/** A projected block plus a fingerprint of the story inputs it was built from. */
interface ProjectedBlockMemo {
  tableBaseline?: YrsTableInheritance;
  tableSourcePrefix?: string;
  /** Bucket key for locating same-container candidates (first cell story / child story). */
  key?: string;
  /** Snapshot of the raw story inputs the block was built from. */
  inputs?: readonly unknown[];
  /** Container cell/child content arrays in payload order (tables, block SDTs). */
  children?: readonly BlockContent[][];
  /** Session key of the pilcrow a paragraph was projected from. */
  sessionKey?: string;
  /** The breaks of a projected paragraph that seeding moves behind it as block embeds. */
  trailing?: readonly BlockBreakKind[];
  sourceRuns?: OriginalRunBoundary[];
}

interface ProjectedStory {
  commentsKey: string;
  blocks: BlockContent[];
}

/** Per-session projection reuse state shared across `yrsToDocument` calls. */
interface SessionProjectionMemo {
  /** An op without story scope ran; every story counts dirty except `clean` ones. */
  wholesale: boolean;
  clean: Set<string>;
  dirty: Set<string>;
  stories: Map<string, ProjectedStory>;
  /** The story each nested story was last projected inside. */
  parents: Map<string, string>;
}

const projectedBlocks = new WeakMap<BlockContent, ProjectedBlockMemo>();

function cloneProjectedParagraph(
  paragraph: Paragraph,
  changes: Partial<Paragraph> = {},
  identity: Paragraph = paragraph
): Paragraph {
  const clone: Paragraph = { ...paragraph, ...changes };
  const memo = projectedBlocks.get(identity);
  if (memo?.sessionKey) {
    projectedBlocks.set(clone, { sessionKey: memo.sessionKey, sourceRuns: memo.sourceRuns });
  }
  return clone;
}

/** The session key of the pilcrow a projected paragraph was built from. @internal */
export function projectedSessionKey(paragraph: Paragraph): string | undefined {
  return projectedBlocks.get(paragraph)?.sessionKey;
}

/**
 * A content-control value's display blocks, each paragraph taking the Word
 * paragraph ID and session key of the child paragraph it replaces in order.
 */
function carryParagraphIdentities(
  replacement: BlockContent[],
  replaced: readonly BlockContent[]
): BlockContent[] {
  const previous = replaced.filter((block): block is Paragraph => block.type === 'paragraph');
  let index = 0;
  return replacement.map((block) => {
    if (block.type !== 'paragraph') return block;
    const source = previous[index++];
    if (!source?.paraId) return block;
    return cloneProjectedParagraph(block, { paraId: source.paraId }, source);
  });
}

/**
 * `blocks` with paragraphs of their own, in table cells too, each still the
 * session paragraph it was projected from: a save edits these, not the
 * projection the next one reuses. @internal
 */
export function ownProjectedParagraphs(blocks: readonly BlockContent[]): BlockContent[] {
  return blocks.map((block) => {
    if (block.type === 'paragraph') {
      return cloneProjectedParagraph(block);
    }
    if (block.type !== 'table') return block;
    const own: Table = {
      ...block,
      rows: block.rows.map((row) => ({
        ...row,
        cells: row.cells.map((cell) => ({ ...cell, content: ownProjectedParagraphs(cell.content) })),
      })),
    };
    const prior = projectedBlocks.get(block);
    if (prior) {
      projectedBlocks.set(own, {
        tableBaseline: prior.tableBaseline,
        tableSourcePrefix: prior.tableSourcePrefix,
      });
    }
    return own;
  });
}
const sessionProjectionMemos = new WeakMap<YrsSession, SessionProjectionMemo>();

function sessionProjectionMemo(session: YrsSession): SessionProjectionMemo {
  let memo = sessionProjectionMemos.get(session);
  if (!memo) {
    memo = {
      wholesale: true,
      clean: new Set(),
      dirty: new Set(),
      stories: new Map(),
      parents: new Map(),
    };
    sessionProjectionMemos.set(session, memo);
  }
  return memo;
}

/** Structural equality over JSON-shaped story inputs — allocation-free sig check. */
function sameJson(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (Array.isArray(a) && Array.isArray(b)) {
    return a.length === b.length && a.every((entry, index) => sameJson(entry, b[index]));
  }
  const left = asObject(a);
  const right = asObject(b);
  if (!left || !right) return false;
  const keys = Object.keys(left);
  return (
    keys.length === Object.keys(right).length &&
    keys.every((key) => key in right && sameJson(left[key], right[key]))
  );
}

const NESTED_STORY_ID = /^(.*?)(?::t\d+:r\d+c\d+(?:~\d+)?|:sdt\d+)$/;

/** Page and column breaks the layout gives a block of their own, ahead of a paragraph. */
const isBlockBreak = (item: InlineItem): boolean =>
  item.kind === 'embed' && blockBreakKind(item.embedKind) !== null;

type BlockBreakKind = 'page' | 'column';

function blockBreakKind(embedKind: string): BlockBreakKind | null {
  return embedKind === 'pageBreak' ? 'page' : embedKind === 'columnBreak' ? 'column' : null;
}

type FlowToken = BlockBreakKind | 'visible';

function runFlowTokens(run: Run, tokens: FlowToken[]): void {
  for (const content of run.content) {
    if (content.type === 'break' && (content.breakType === 'page' || content.breakType === 'column')) {
      tokens.push(content.breakType);
    } else if (content.type !== 'text' || content.text.length > 0) tokens.push('visible');
  }
}

function flowTokens(content: readonly ParagraphContent[], tokens: FlowToken[] = []): FlowToken[] {
  for (const item of content) {
    if (item.type === 'run') runFlowTokens(item, tokens);
    else if (item.type === 'hyperlink') {
      for (const child of item.children) if (child.type === 'run') runFlowTokens(child, tokens);
    } else if (item.type === 'simpleField') {
      for (const child of item.content) if (child.type === 'run') runFlowTokens(child, tokens);
    } else if (item.type === 'complexField') {
      for (const child of [...item.fieldCode, ...item.fieldResult]) runFlowTokens(child, tokens);
    } else if (item.type === 'inlineSdt') flowTokens(item.content as ParagraphContent[], tokens);
    else if (item.type === 'insertion' || item.type === 'deletion' || item.type === 'moveFrom' || item.type === 'moveTo') {
      for (const child of item.content) if (child.type === 'run') runFlowTokens(child, tokens);
    } else if (item.type === 'mathEquation') tokens.push('visible');
  }
  return tokens;
}

function leadingBlockBreaks(paragraph: Paragraph): BlockBreakKind[] {
  const tokens = flowTokens(paragraph.content);
  if (!tokens.includes('visible')) return tokens.slice(0, tokens.lastIndexOf('column') + 1) as BlockBreakKind[];
  return tokens[0] === 'column' ? ['column'] : [];
}

/** The breaks of `paragraph` that seeding moves behind it as block embeds, in order. */
function trailingBlockBreaks(paragraph: Paragraph): BlockBreakKind[] {
  const tokens = flowTokens(paragraph.content);
  if (!tokens.includes('visible')) return tokens.slice(tokens.lastIndexOf('column') + 1) as BlockBreakKind[];
  const trailing: BlockBreakKind[] = [];
  let leading = false;
  let visible = false;
  for (const token of tokens) {
    if (token === 'visible') visible = true;
    else if (visible || leading) trailing.push(token);
    else leading = true;
  }
  return trailing;
}

/** Stories a committed op touched; `all` invalidates the whole session. */
export function noteYrsStoriesDirty(
  session: YrsSession,
  stories: 'all' | string | Iterable<string>
): void {
  const memo = sessionProjectionMemo(session);
  if (stories === 'all') {
    memo.wholesale = true;
    memo.clean.clear();
    memo.dirty.clear();
    memo.stories.clear();
    memo.parents.clear();
    return;
  }
  const queue = typeof stories === 'string' ? [stories] : [...stories];
  const queued = new Set(queue);
  for (let index = 0; index < queue.length; index += 1) {
    const story = queue[index]!;
    memo.dirty.add(story);
    memo.clean.delete(story);
    memo.stories.delete(story);
    const parent = memo.parents.get(story) ?? NESTED_STORY_ID.exec(story)?.[1];
    if (parent && !queued.has(parent)) {
      queued.add(parent);
      queue.push(parent);
    }
  }
}

const EMPTY_BLOCKS: BlockContent[] = [];

class SaveContext {
  readonly storyIds: Set<string>;
  readonly warnings: string[] = [];
  private readonly baseParagraphs: Map<string, Paragraph>;
  private readonly baseStories: Map<string, readonly BlockContent[]>;
  private readonly baseChartDrawings = new Map<string, Map<string, string>>();
  private readonly comments: Map<string, Array<{ id: number; start: number; end: number }>>;
  /** Per story, the comment ranges of it and of the stories nested in it, which key its blocks. */
  private readonly subtreeComments = new Map<string, Map<string, unknown>>();
  private readonly memo: SessionProjectionMemo;
  private readonly styles: StyleResolver;
  private readonly syntheticDefaults: boolean;
  private readonly inherited = new Map<string, ParagraphFormatting | undefined>();
  private readonly styleDefinitions: Document['package']['styles'];
  private runInheritance: YrsRunInheritance | null = null;
  private readonly inheritedRuns = new Map<string, TextFormatting>();
  private readonly sourceTables = new Map<string, Table>();
  private readonly tableBaselines = new Map<Table, YrsTableInheritance>();
  private readonly payloadBaselines = new WeakMap<TablePayload, YrsTableInheritance>();
  private readonly sourceCells = new Map<string, YrsTableInheritance['rows'][number][number]>();
  private readonly cellParagraphs = new Map<string, ParagraphFormatting | undefined>();
  private readonly cellRuns = new Map<string, TextFormatting | undefined>();

  constructor(
    private readonly session: YrsSession,
    base: Document,
    commentIds?: ReadonlyMap<number, string>,
    private readonly revisionIds?: RevisionIds
  ) {
    this.storyIds = new Set(session.storyIds());
    this.baseStories = collectBaseStories(session, base);
    for (const [storyId, blocks] of this.baseStories) {
      let tableIndex = 0;
      for (const block of blocks) {
        if (block.type !== 'table') continue;
        const prefix = projectedBlocks.get(block)?.tableSourcePrefix ?? `${storyId}:t${tableIndex}`;
        this.sourceTables.set(prefix, block);
        tableIndex += 1;
      }
      const scoped = new Map<string, string>();
      collectChartDrawings(blocks, scoped);
      if (scoped.size > 0) this.baseChartDrawings.set(storyId, scoped);
    }
    this.baseParagraphs = collectBaseParagraphs(this.baseStories);
    this.comments = commentRanges(session, base.package.document.comments, commentIds);
    this.memo = sessionProjectionMemo(session);
    this.styleDefinitions = base.package.styles;
    this.styles = createStyleResolver(base.package.styles);
    const defaultStyle = this.styles.getDefaultParagraphStyle();
    this.syntheticDefaults =
      defaultStyle !== undefined && !(base.package.styles?.styles ?? []).includes(defaultStyle);
    for (const [story, ranges] of this.comments) {
      const seen = new Set<string>();
      for (
        let owner: string | undefined = story;
        owner !== undefined && !seen.has(owner);
        owner = this.memo.parents.get(owner) ?? NESTED_STORY_ID.exec(owner)?.[1]
      ) {
        seen.add(owner);
        const nested = this.subtreeComments.get(owner) ?? new Map<string, unknown>();
        nested.set(story, ranges);
        this.subtreeComments.set(owner, nested);
      }
    }
  }

  private storyIsClean(storyId: string): boolean {
    return this.memo.wholesale
      ? this.memo.clean.has(storyId)
      : !this.memo.dirty.has(storyId);
  }

  private inheritance(): YrsRunInheritance {
    return (this.runInheritance ??= this.session.runInheritance(this.styleDefinitions));
  }

  private prepareTable(payload: TablePayload, parent: string): void {
    const cells = (Array.isArray(payload.rows) ? payload.rows : [])
      .flatMap((row) => Array.isArray(row.cells) ? row.cells : []);
    let source: { prefix: string; table: Table } | undefined;
    const candidates = payload.tblPr?._originalFormatting ? cells : [];
    for (const cell of candidates) {
      const story = asString(cell.tcPr?._sourceCellStory) ?? asString(cell.story);
      const prefix = story?.match(/^(.*:t\d+):r\d+c\d+(?:~\d+)?$/)?.[1];
      const table = prefix ? this.sourceTables.get(prefix) : undefined;
      if (prefix && table) {
        source = { prefix, table };
        break;
      }
    }
    if (source) {
      let baseline = this.tableBaselines.get(source.table);
      if (!baseline) {
        baseline = projectedBlocks.get(source.table)?.tableBaseline ??
          this.inheritance().tableAttributes(source.table);
        this.tableBaselines.set(source.table, baseline);
      }
      this.payloadBaselines.set(payload, baseline);
      const prefix = source.prefix;
      baseline.rows.forEach((row, rowIndex) => {
        row.forEach((cell, cellIndex) => {
          this.sourceCells.set(`${prefix}:r${rowIndex}c${cellIndex}`, cell);
        });
      });
    }
    for (const cell of cells) {
      if (typeof cell.story !== 'string') continue;
      this.memo.parents.set(cell.story, parent);
      const source = asString(cell.tcPr?._sourceCellStory) ?? cell.story;
      this.cellParagraphs.set(cell.story, this.sourceCells.get(source)?.paragraphFormatting ?? undefined);
      this.cellRuns.set(cell.story, this.sourceCells.get(source)?.runFormatting ?? undefined);
    }
  }

  applyTableBaseline(payload: TablePayload, attrs: TableSaveAttrs): void {
    const baseline = this.payloadBaselines.get(payload);
    if (baseline) {
      attrs._originalResolvedMargins = baseline.attributes.cellMargins as TableSaveAttrs['cellMargins'];
    }
  }

  cellAttrs(payload: TableCellPayload): TableCellSaveAttrs {
    const attrs = {
      ...TABLE_CELL_ATTR_DEFAULTS,
      ...(payload.tcPr ?? {}),
    } as unknown as TableCellSaveAttrs;
    const origin = asString(payload.tcPr?._sourceCellStory);
    const source = origin ?? asString(payload.story);
    const baseline = source !== undefined ? this.sourceCells.get(source) : undefined;
    if (baseline) {
      attrs._originalResolvedBorders = baseline.attributes.borders as TableCellSaveAttrs['borders'];
      attrs._originalResolvedMargins = baseline.attributes.margins as TableCellSaveAttrs['margins'];
      const original = asObject(baseline.attributes._originalFormatting);
      if (origin && original && !attrs._originalFormatting) {
        const formatting = { ...original };
        delete formatting.gridSpan;
        delete formatting.vMerge;
        attrs._originalFormatting = formatting as TableCellSaveAttrs['_originalFormatting'];
      }
    }
    return attrs;
  }

  private cellContents(payload: TablePayload, parent: string): BlockContent[][] {
    const contents: BlockContent[][] = [];
    for (const row of Array.isArray(payload.rows) ? payload.rows : []) {
      for (const cell of Array.isArray(row.cells) ? row.cells : []) {
        if (cell.story !== undefined) this.memo.parents.set(cell.story, parent);
        contents.push(
          cell.story !== undefined && this.storyIds.has(cell.story)
            ? this.storyToBlocks(cell.story)
            : EMPTY_BLOCKS
        );
      }
    }
    return contents;
  }

  private inheritedFormatting(storyId: string, properties: Attrs): ParagraphFormatting | undefined {
    const cell = this.cellParagraphs.has(storyId);
    const nested = !cell && NESTED_STORY_ID.test(storyId);
    if (this.syntheticDefaults && !cell && !nested) {
      return undefined;
    }
    const numbered = properties.numPr != null;
    const style = typeof properties.pStyle === 'string' ? properties.pStyle : '';
    const tableFormatting = this.cellParagraphs.get(storyId);
    const key = cell
      ? JSON.stringify([style, tableFormatting, numbered])
      : JSON.stringify([style, numbered, nested]);
    if (!this.inherited.has(key)) {
      const resolved = cell
        ? this.inheritance().paragraphFormatting(style || null, tableFormatting)
        : this.styles.resolveParagraphStyle(style || null).paragraphFormatting;
      if (resolved === undefined || (!numbered && resolved.numPr != null)) {
        this.inherited.set(key, undefined);
      } else if (numbered || nested) {
        const formatting: ParagraphFormatting = {};
        if (resolved.autoSpaceDE !== undefined) formatting.autoSpaceDE = resolved.autoSpaceDE;
        if (resolved.autoSpaceDN !== undefined) formatting.autoSpaceDN = resolved.autoSpaceDN;
        if (resolved.adjustRightInd !== undefined) {
          formatting.adjustRightInd = resolved.adjustRightInd;
        }
        this.inherited.set(key, formatting);
      } else {
        const formatting = { ...resolved };
        delete formatting.indentLeft;
        delete formatting.indentRight;
        this.inherited.set(key, formatting);
      }
    }
    return this.inherited.get(key);
  }

  /**
   * How the runs of a paragraph with pilcrow `properties` save: relative to
   * what seeding merged into their marks under the styles the package saves.
   */
  private runFormatter(story: string, properties: Attrs): RunFormatter {
    const paragraphStyle = typeof properties.pStyle === 'string' ? properties.pStyle : null;
    return (attributes) =>
      relativeFormatting(
        attributes,
        this.inheritedRun(paragraphStyle, asString(asObject(attributes.runStyle)?.styleId) || null, this.cellRuns.get(story))
      );
  }

  private inheritedRun(paragraphStyle: string | null, runStyle: string | null, tableRun: TextFormatting | undefined): TextFormatting {
    const key = JSON.stringify([paragraphStyle, runStyle, tableRun]);
    let inherited = this.inheritedRuns.get(key);
    if (inherited === undefined) {
      const inheritance = this.inheritance();
      inherited = attrsToTextFormatting(inheritance.attributes(paragraphStyle, runStyle, tableRun));
      this.inheritedRuns.set(key, inherited);
    }
    return inherited;
  }

  dispose(): void {
    this.runInheritance?.dispose();
    this.runInheritance = null;
  }

  recovery(storyId: string): ChartRecovery {
    const scoped = this.baseChartDrawings.get(storyId);
    return {
      drawingFor: (rId, path) =>
        (rId !== undefined ? scoped?.get(`rId:${rId}`) : undefined) ??
        (path !== undefined ? scoped?.get(`path:${path}`) : undefined),
      unrecoverable: (rId, path) => {
        const identity = rId !== undefined ? ` (rId ${rId})` : path !== undefined ? ` (path ${path})` : '';
        this.warnings.push(
          `chart run carries no drawing to replay${identity}; keeping the run out of the output`
        );
      },
      warn: (message) => {
        this.warnings.push(message);
      },
    };
  }

  storyToBlocks(storyId: string): BlockContent[] {
    const baseBlocks = this.baseStories.get(storyId);
    const storyComments = this.comments.get(storyId) ?? [];
    const nestedComments = this.subtreeComments.get(storyId);
    const commentsKey = nestedComments ? stableStringify(Object.fromEntries(nestedComments)) : '';
    const priorStory = this.memo.stories.get(storyId);
    // Nested stories are mapped positionally, so a different array means
    // drift, not divergence — mutation paths mark them dirty anyway. A root
    // story on a different array is genuine divergence and forces a fresh
    // projection.
    if (
      priorStory !== undefined &&
      priorStory.commentsKey === commentsKey &&
      this.storyIsClean(storyId) &&
      (NESTED_STORY_ID.test(storyId) ||
        baseBlocks === undefined ||
        baseBlocks === priorStory.blocks)
    ) {
      return priorStory.blocks;
    }
    const blocks: BlockContent[] = [];
    const paragraphIds = this.session.storyParagraphIds(storyId);
    const segments = this.session.storySegments(storyId);
    const trailingByParagraph = new Map<string, BlockBreakKind[]>();
    let closing: string | undefined;
    for (const segment of segments) {
      if (segment.kind === 'pilcrow') {
        closing = segment.paraId;
        trailingByParagraph.set(closing, []);
      } else {
        const kind = segment.kind === 'embed' ? blockBreakKind(segment.embedKind) : null;
        if (kind && closing !== undefined) trailingByParagraph.get(closing)!.push(kind);
        else closing = undefined;
      }
    }
    let items: InlineItem[] = [];
    let paragraphStart = 0;
    let inlineStart = 0;
    let paragraphIndex = 0;
    let storyOffset = 0;
    let candidatesByKey: Map<string, BlockContent[]> | null = null;

    const candidatesFor = (key: string): BlockContent[] => {
      if (!candidatesByKey) {
        candidatesByKey = new Map();
        for (const block of baseBlocks ?? []) {
          const blockKey = projectedBlocks.get(block)?.key;
          if (blockKey === undefined) continue;
          const bucket = candidatesByKey.get(blockKey);
          if (bucket) bucket.push(block);
          else candidatesByKey.set(blockKey, [block]);
        }
      }
      return candidatesByKey.get(key) ?? [];
    };

    const reuseContainer = (
      key: string,
      inputs: readonly unknown[],
      children: readonly BlockContent[][]
    ): BlockContent | undefined =>
      candidatesFor(key).find((block) => {
        const memo = projectedBlocks.get(block);
        return (
          memo?.inputs !== undefined &&
          memo.children !== undefined &&
          sameJson(memo.inputs, inputs) &&
          memo.children.length === children.length &&
          memo.children.every((child, index) => child === children[index])
        );
      });

    const paragraphCommentBoundaries = (end: number): CommentBoundary[] => {
      const boundaries: CommentBoundary[] = [];
      for (const range of storyComments) {
        if (range.start >= paragraphStart && range.start <= end) {
          const offset = Math.max(0, range.start - inlineStart);
          boundaries.push({ id: range.id, kind: 'start', offset });
        }
        if (range.end >= paragraphStart && range.end <= end) {
          boundaries.push({ id: range.id, kind: 'end', offset: Math.max(0, range.end - inlineStart) });
        }
      }
      return boundaries;
    };

    const pushText = (text: string, attributes: Attrs): void => {
      let cursor = 0;
      for (let index = 0; index < text.length; index += 1) {
        if (text[index] !== '\t') continue;
        if (index > cursor)
          items.push({ kind: 'text', text: text.slice(cursor, index), attributes });
        items.push({ kind: 'embed', embedKind: 'tab', payload: {}, attributes });
        cursor = index + 1;
      }
      if (cursor < text.length) items.push({ kind: 'text', text: text.slice(cursor), attributes });
    };

    // Block breaks the previous paragraph's own content carries, and those only the stream does.
    let witnesses: BlockBreakKind[] = [];
    let breaks: BlockBreakKind[] = [];
    for (const segment of segments) {
      engineProgressTick();
      if (segment.kind === 'text') {
        pushText(segment.text, segment.attributes);
        storyOffset += segment.text.length;
        witnesses = [];
        continue;
      }
      if (segment.kind === 'pilcrow') {
        const liveTrailing = trailingByParagraph.get(segment.paraId)!;
        const savedParaId = paragraphIds[paragraphIndex] ?? '';
        const baseParagraph = this.baseParagraphs.get(segment.paraId);
        const boundaries = paragraphCommentBoundaries(storyOffset);
        const source = sourceBreakBoundaries(baseParagraph);
        const sourceRuns = source.some(boundary => boundary.breaks?.length) ? source : undefined;
        const inputs = [
          segment.paraId,
          savedParaId,
          segment.properties,
          segment.attributes,
          items,
          boundaries,
          breaks,
          liveTrailing,
          sourceRuns,
        ] as const;
        const priorMemo =
          baseParagraph !== undefined ? projectedBlocks.get(baseParagraph) : undefined;
        let paragraph: Paragraph;
        if (
          baseParagraph !== undefined &&
          priorMemo?.inputs !== undefined &&
          sameJson(priorMemo.inputs, inputs)
        ) {
          paragraph = baseParagraph;
          witnesses = [...(priorMemo.trailing ?? trailingBlockBreaks(baseParagraph))];
        } else {
          // Projection mutates `item.payload`; pin the pre-mutation inputs.
          const snapshot = inputs.map((input, index) =>
            index === 4 ? (input as InlineItem[]).map((item) => ({ ...item })) : input
          );
          paragraph = paragraphFromStory(
            savedParaId,
            segment.properties,
            items,
            boundaries,
            baseParagraph,
            { leading: breaks, trailing: liveTrailing, leadingPage: segment.properties.pageBreakBeforeRun === true },
            this.recovery(storyId),
            this.runFormatter(storyId, segment.properties),
            this.revisionIds,
            this.inheritedFormatting(storyId, segment.properties)
          );
          const trailing = trailingBlockBreaks(paragraph);
          const leading = leadingBlockBreaks(paragraph);
          const consumed = leading.every((kind, index) => kind === breaks[index]) ? leading.length : 0;
          const pending = breaks.slice(consumed);
          if (pending.length) {
            paragraph = cloneProjectedParagraph(paragraph, {
              content: [
                ...pending.map((breakType): ParagraphContent => ({
                  type: 'run', content: [{ type: 'break', breakType }],
                })),
                ...paragraph.content,
              ],
            });
          }
          projectedBlocks.set(paragraph, { inputs: snapshot, sessionKey: segment.paraId, trailing, sourceRuns });
          witnesses = [...trailing];
        }
        blocks.push(paragraph);
        items = [];
        breaks = [];
        paragraphIndex += 1;
        storyOffset += 1;
        paragraphStart = storyOffset;
        inlineStart = storyOffset;
        continue;
      }

      if (segment.embedKind === 'table') {
        const payload = segment.payload as TablePayload;
        this.prepareTable(payload, storyId);
        const inputs = [segment.attributes, payload] as const;
        const firstCell = Array.isArray(payload.rows) ? payload.rows[0]?.cells?.[0] : undefined;
        const key = `T${firstCell?.story ?? ''}`;
        const contents =
          candidatesFor(key).length > 0 ? this.cellContents(payload, storyId) : undefined;
        const reused =
          contents !== undefined
            ? (reuseContainer(key, inputs, contents) as Table | undefined)
            : undefined;
        let table: Table;
        if (reused) {
          table = reused;
        } else {
          table = tableFromPayload(this, payload);
          projectedBlocks.set(table, {
            key,
            inputs,
            tableBaseline: this.payloadBaselines.get(payload),
            tableSourcePrefix: asString(firstCell?.story)?.match(/^(.*:t\d+):r\d+c\d+(?:~\d+)?$/)?.[1],
            children: contents ?? this.cellContents(payload, storyId),
          });
        }
        blocks.push(table);
      } else if (segment.embedKind === 'blockSdt') {
        const childStory = asString(segment.payload.story);
        if (childStory) this.memo.parents.set(childStory, storyId);
        const childContent =
          childStory && this.storyIds.has(childStory)
            ? this.storyToBlocks(childStory)
            : EMPTY_BLOCKS;
        const inputs = [segment.attributes, segment.payload] as const;
        const reused = reuseContainer(`S${childStory ?? ''}`, inputs, [childContent]);
        if (reused) {
          blocks.push(reused);
        } else {
          let properties = sdtAttrsToProps(segment.payload);
          let content = childContent;
          const authoredValue = contentControlValue(segment.payload.value);
          if (authoredValue) {
            try {
              const applied = applyContentControlValue(properties, authoredValue);
              properties = applied.properties;
              content = carryParagraphIdentities(applied.content, childContent);
            } catch {
              // Retain the child story if the authored value is invalid.
            }
          }
          const block: BlockContent = {
            type: 'blockSdt',
            properties,
            content: content === EMPTY_BLOCKS ? [] : content,
          };
          // Rebind the child's entry to the embedded array for the next projection.
          if (content !== childContent && childStory !== undefined) {
            const childEntry = this.memo.stories.get(childStory);
            if (childEntry) childEntry.blocks = block.content;
          }
          projectedBlocks.set(block, {
            key: `S${childStory ?? ''}`,
            inputs,
            children: [block.content],
          });
          blocks.push(block);
        }
      } else if (segment.embedKind === 'opaque') {
        const blob = asObject(segment.payload.blob);
        if (blob?.type === 'pageBreak') blocks.push(pageBreakParagraph());
        else {
          // Carry a same-position base block while block SDTs stay opaque.
          const baseBlock = baseBlocks?.[blocks.length];
          if (blob?.type === 'blockSdt' && baseBlock?.type === 'blockSdt') {
            blocks.push(baseBlock);
          }
          // Standalone text-box blobs do not have a one-to-one base block (the
          // base model stores them inside paragraph runs), so they remain the
          // documented opaque carry gap.
        }
      } else {
        items.push(segment as EmbedItem);
      }
      const kind = blockBreakKind(segment.embedKind);
      if (!kind) witnesses = [];
      else if (witnesses.includes(kind)) witnesses.splice(witnesses.indexOf(kind), 1);
      else breaks.push(kind);
      storyOffset += 1;
      if (items.every(isBlockBreak)) inlineStart = storyOffset;
    }

    const pendingBreaks = breaks.map((breakType): ParagraphContent => ({
      type: 'run', content: [{ type: 'break', breakType }],
    }));
    // Defensive recovery for malformed/legacy stories without a final pilcrow.
    if (items.length > 0) {
      const trailing = buildParagraphContent(
        items,
        this.recovery(storyId),
        directFormatting,
        this.revisionIds
      );
      if (trailing.length > 0) {
        blocks.push({ type: 'paragraph', content: breaks.length ? [...pendingBreaks, ...trailing] : trailing });
        breaks = [];
      }
    }
    if (breaks.length > 0) {
      let index = blocks.length - 1;
      while (index >= 0 && blocks[index]!.type !== 'paragraph') index -= 1;
      const previous = blocks[index];
      if (previous?.type === 'paragraph') {
        blocks[index] = cloneProjectedParagraph(previous, {
          content: [...previous.content, ...pendingBreaks],
        });
      } else {
        blocks.push({ type: 'paragraph', content: pendingBreaks });
      }
    }
    const projected = restoreRawBlocks(blocks, baseBlocks ?? [], storyId);
    this.memo.stories.set(storyId, { commentsKey, blocks: projected });
    this.memo.dirty.delete(storyId);
    this.memo.clean.add(storyId);
    return projected;
  }
}

/**
 * Rebuild all yrs-owned editable stories into the serializer-facing Document
 * while preserving every package part the editor does not own.
 */
export interface YrsToDocumentOptions {
  /**
   * Root stories to project. Omit to rebuild every editable story. A `body`
   * projection includes its nested table/content-control stories.
   */
  storyIds?: ReadonlySet<string>;
  /**
   * The session comment id of each comment in `base` whose OOXML id is not
   * that id, such as one added with {@link YrsSession.addComment}. @internal
   */
  commentIds?: ReadonlyMap<number, string>;
}

export function yrsToDocument(
  session: YrsSession,
  base: Document,
  options: YrsToDocumentOptions = {}
): Document {
  return projectDocument(session, base, options);
}

/**
 * {@link yrsToDocument} of the body alone, writing each session revision with the serialized
 * ids `revisionIds` gives its deletion and insertion.
 */
export function yrsBodyToDocumentWithRevisionIds(
  session: YrsSession,
  base: Document,
  revisionIds: ReadonlyMap<string, RevisionNumbers>
): Document {
  return projectDocument(session, base, { storyIds: new Set(['body']) }, revisionIds);
}

function projectDocument(
  session: YrsSession,
  base: Document,
  options: YrsToDocumentOptions,
  revisionIds?: RevisionIds
): Document {
  const previous = projectedMedia;
  projectedMedia = (token) => session.mediaDataUrl?.(token) ?? null;
  let context: SaveContext | undefined;
  try {
    context = new SaveContext(session, base, options.commentIds, revisionIds);
    return projectStories(context, session, base, options);
  } finally {
    context?.dispose();
    projectedMedia = previous;
  }
}

function projectStories(
  context: SaveContext,
  session: YrsSession,
  base: Document,
  options: YrsToDocumentOptions
): Document {
  const shouldProject = (storyId: string): boolean =>
    options.storyIds === undefined || options.storyIds.has(storyId);
  const bodyContent = context.storyIds.has('body') && shouldProject('body')
    ? context.storyToBlocks('body')
    : base.package.document.content;

  const selectedHeaderFooterStories = new Set(
    [...(base.package.headers?.keys() ?? []), ...(base.package.footers?.keys() ?? [])]
      .filter((rId) => shouldProject(`hf:${rId}`))
      .map((rId) => headerFooterStory(session, rId))
  );
  const shouldProjectHeaderFooter = (storyId: string): boolean =>
    shouldProject(storyId) || selectedHeaderFooterStories.has(storyId);

  const headerFooterContents = new Map<string, BlockContent[]>();
  const projectHeaderFooter = (storyId: string): BlockContent[] => {
    let content = headerFooterContents.get(storyId);
    if (!content) {
      content = context.storyToBlocks(storyId);
      headerFooterContents.set(storyId, content);
    }
    return content;
  };

  const canonicalFor = (rId: string, storyId: string): string | undefined =>
    storyId === `hf:${rId}` ? undefined : storyId.slice(3);
  const projectHeaderFooters = (
    entries: Map<string, HeaderFooter> | undefined
  ): Map<string, HeaderFooter> | undefined => {
    if (!entries) return entries;
    const routes = [...entries].map(
      ([rId, part]) => [rId, part, headerFooterStory(session, rId)] as const
    );
    const project =
      options.storyIds === undefined ||
      routes.some(([, , storyId]) => shouldProjectHeaderFooter(storyId));
    const remark = routes.some(
      ([rId, part, storyId]) => headerFooterAliasOf(part) !== canonicalFor(rId, storyId)
    );
    if (!project && !remark) return entries;
    return new Map(
      routes.map(([rId, part, storyId]) => {
        const canonical = canonicalFor(rId, storyId);
        const next =
          context.storyIds.has(storyId) && shouldProjectHeaderFooter(storyId)
            ? { ...part, content: projectHeaderFooter(storyId) }
            : headerFooterAliasOf(part) === canonical
              ? part
              : { ...part };
        if (canonical !== undefined) markHeaderFooterAlias(next, canonical);
        return [rId, next];
      })
    );
  };
  const headers = projectHeaderFooters(base.package.headers);
  const footers = projectHeaderFooters(base.package.footers);

  const projectNotes = <T extends Footnote | Endnote>(
    notes: T[] | undefined,
    prefix: string
  ): T[] | undefined => {
    const shouldProjectNotes =
      options.storyIds === undefined ||
      notes?.some((note) => shouldProject(`${prefix}${note.id}`));
    if (!shouldProjectNotes) return notes;
    return notes?.map((note) => {
      const storyId = `${prefix}${note.id}`;
      return context.storyIds.has(storyId) && shouldProject(storyId)
        ? {
            ...note,
            content: restoreNoteMarks(context.storyToBlocks(storyId), note.content),
            verbatimXml: undefined,
          }
        : note;
    });
  };
  const footnotes = projectNotes(base.package.footnotes, 'fn:');
  const endnotes = projectNotes(base.package.endnotes, 'en:');

  return {
    ...base,
    ...(context.warnings.length > 0
      ? { warnings: [...(base.warnings ?? []), ...context.warnings] }
      : {}),
    package: {
      ...base.package,
      document: {
        ...base.package.document,
        content: bodyContent,
      },
      ...(headers ? { headers } : {}),
      ...(footers ? { footers } : {}),
      ...(footnotes ? { footnotes } : {}),
      ...(endnotes ? { endnotes } : {}),
    },
  };
}
