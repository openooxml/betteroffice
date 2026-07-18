/**
 * yrs -> OOXML Document save projection.
 *
 * The yrs coexistence model stores each editable container as a flat story:
 * text carries inline formatting attributes, pilcrows carry paragraph
 * properties, and structural/inline atoms are map embeds. This module reads
 * that representation directly and rebuilds the serializer-facing Document.
 */

/* eslint-disable max-lines -- the inverse mapping stays co-located with its save orchestrator */

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
  Run,
  RunContent,
  TextFormatting,
  Hyperlink,
  TrackedChangeInfo,
  Table,
  TableRow,
  TableCell,
  TableBorders,
  SimpleField,
  ComplexField,
  FieldType,
  MathEquation,
  Image,
  Shape,
  InlineSdt,
  SdtProperties,
  Comment,
} from '../types/document';
import type { YrsSession } from './index';

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
    horizontal?: { relativeTo?: string; posOffset?: number; align?: string };
    vertical?: { relativeTo?: string; posOffset?: number; align?: string };
  };
  borderWidth?: number;
  borderColor?: string;
  borderStyle?: string;
  wrapText?: string;
  hlinkHref?: string;
  cropTop?: number;
  cropRight?: number;
  cropBottom?: number;
  cropLeft?: number;
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
  contextualSpacing: null,
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
  marksKey?: string;
  formatting?: TextFormatting;
  propertyChanges?: Run['propertyChanges'];
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

/** Mirrors fromProseDoc/marks.ts without manufacturing PM Mark objects. */
function attrsToTextFormatting(attributes: Attrs): TextFormatting {
  const formatting: TextFormatting = {};

  if (attributes.bold) {
    formatting.bold = true;
    formatting.boldCs = true;
  }
  if (attributes.italic) {
    formatting.italic = true;
    formatting.italicCs = true;
  }

  const underline = asObject(attributes.underline);
  if (underline) {
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
  if (textColor) {
    formatting.color = {
      rgb: (asString(textColor.rgb) ?? null) as string | undefined,
      themeColor: (textColor.themeColor ?? null) as NonNullable<
        TextFormatting['color']
      >['themeColor'],
      themeTint: (asString(textColor.themeTint) ?? null) as string | undefined,
      themeShade: (asString(textColor.themeShade) ?? null) as string | undefined,
    };
  }

  if (typeof attributes.highlight === 'string') {
    formatting.highlight = attributes.highlight as TextFormatting['highlight'];
  }

  const fontSize = asObject(attributes.fontSize);
  if (fontSize) {
    const size = asFiniteNumber(fontSize.size);
    const sizeCs = asFiniteNumber(fontSize.sizeCs) ?? size;
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
      cs: asString(fontFamily.cs) || ascii,
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

function createTextRun(text: string, attributes: Attrs): Run {
  const formatting = attrsToTextFormatting(formattingAttrs(attributes));
  return {
    type: 'run',
    formatting: Object.keys(formatting).length > 0 ? formatting : undefined,
    content: runContentForText(text, formatting),
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

function trackedInfo(raw: unknown, _pmShape = false): TrackedChangeInfo | null {
  const value = asObject(raw);
  if (!value) return null;
  const author = asString(value.author) || 'Unknown';
  const date = asString(value.date);
  return {
    id: revisionId(value.revisionId ?? value.id),
    author,
    ...(date ? { date } : {}),
  };
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

function fieldFromPayload(payload: Attrs, attributes: Attrs): SimpleField | ComplexField {
  const fieldData = asString(payload.fieldData);
  if (fieldData && fieldData.length <= 2_000_000) {
    try {
      const stored = JSON.parse(fieldData) as SimpleField | ComplexField;
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
      // Malformed editor cache: rebuild the same minimal inert field as PM save.
    }
  }

  const formatting = attrsToTextFormatting(formattingAttrs(attributes));
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

function imageRunFromPayload(payload: Attrs): Run {
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
    src: asString(attrs.src) || '',
    alt: asString(attrs.alt) || undefined,
    title: asString(attrs.title) || undefined,
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
  if (attrs.effectExtentTop) padding.top = pixelsToEmu(attrs.effectExtentTop);
  if (attrs.effectExtentBottom) padding.bottom = pixelsToEmu(attrs.effectExtentBottom);
  if (attrs.effectExtentLeft) padding.left = pixelsToEmu(attrs.effectExtentLeft);
  if (attrs.effectExtentRight) padding.right = pixelsToEmu(attrs.effectExtentRight);
  if (Object.keys(padding).length > 0) image.padding = padding;

  return { type: 'run', content: [{ type: 'drawing', image }] };
}

function shapeRunFromPayload(payload: Attrs): Run {
  const shape: Shape = {
    type: 'shape',
    shapeType: (asString(payload.shapeType) || 'rect') as Shape['shapeType'],
    id: asString(payload.shapeId) || undefined,
    size: {
      width: payload.width ? pixelsToEmu(Number(payload.width)) : 0,
      height: payload.height ? pixelsToEmu(Number(payload.height)) : 0,
    },
  };
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

function inlineSdtFromPayload(payload: Attrs): InlineSdt {
  let properties: SdtProperties = sdtAttrsToProps(payload);
  const propertiesJson = asString(payload.propertiesJson);
  if (propertiesJson && propertiesJson.length <= 1_000_000) {
    try {
      const parsed = JSON.parse(propertiesJson) as SdtProperties;
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
  let content = inlineSdtContent(buildParagraphContent(items));
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

function ordinaryContentForItem(item: InlineItem): ParagraphContent | null {
  if (item.kind === 'text') return createTextRun(item.text, item.attributes);
  switch (item.embedKind) {
    case 'break':
      return { type: 'run', content: [{ type: 'break', breakType: 'textWrapping' }] };
    case 'tab':
      return { type: 'run', content: [{ type: 'tab' }] };
    case 'image':
      return imageRunFromPayload(item.payload);
    case 'shape':
      return shapeRunFromPayload(item.payload);
    case 'field':
      return (
        commentReferenceFromPayload(item.payload) ?? fieldFromPayload(item.payload, item.attributes)
      );
    case 'math':
      return mathFromPayload(item.payload);
    case 'sdt':
      return inlineSdtFromPayload(item.payload);
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

function trackedContentForItem(item: InlineItem, info: TrackedChangeInfo): ParagraphContent {
  let run: Run;
  if (item.kind === 'embed' && item.embedKind === 'image') run = imageRunFromPayload(item.payload);
  else if (item.kind === 'embed' && item.embedKind === 'shape')
    run = shapeRunFromPayload(item.payload);
  else if (item.kind === 'text') {
    const formatting = attrsToTextFormatting(formattingAttrs(item.attributes));
    run = {
      type: 'run',
      content: [{ type: 'text', text: item.text }],
      ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
    };
  } else run = { type: 'run', content: [] };

  const raw = asObject(item.attributes.ins) ?? asObject(item.attributes.del);
  const isMovePair = raw?.isMovePair === true;
  if (item.attributes.ins) {
    return isMovePair
      ? { type: 'moveTo', info, content: [run] }
      : { type: 'insertion', info, content: [run] };
  }
  return isMovePair
    ? { type: 'moveFrom', info, content: [run] }
    : { type: 'deletion', info, content: [run] };
}

function addToHyperlink(hyperlink: Hyperlink, item: InlineItem): void {
  if (item.kind === 'text') {
    hyperlink.children.push(createTextRun(item.text, item.attributes));
    return;
  }
  if (item.embedKind === 'break') {
    hyperlink.children.push({
      type: 'run',
      content: [{ type: 'break', breakType: 'textWrapping' }],
    });
  } else if (item.embedKind === 'tab') {
    hyperlink.children.push({ type: 'run', content: [{ type: 'tab' }] });
  } else if (item.embedKind === 'field') {
    const child =
      commentReferenceFromPayload(item.payload) ?? fieldFromPayload(item.payload, item.attributes);
    if (child.type === 'run') hyperlink.children.push(child);
    else (hyperlink.structuredChildren ??= [...hyperlink.children]).push(child);
  } else if (item.embedKind === 'math') {
    (hyperlink.structuredChildren ??= [...hyperlink.children]).push(mathFromPayload(item.payload));
  }
}

function buildParagraphContent(items: InlineItem[]): ParagraphContent[] {
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
    // A note reference is handled before tracked/link marks by fromProseDoc.
    if (item.kind === 'embed' && item.embedKind === 'noteRef') {
      flushRun();
      flushHyperlink();
      const note = ordinaryContentForItem(item);
      if (note) content.push(note);
      continue;
    }

    const revision = trackedInfo(item.attributes.ins ?? item.attributes.del);
    if (revision) {
      flushRun();
      flushHyperlink();
      content.push(trackedContentForItem(item, revision));
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
      if (currentHyperlink) addToHyperlink(currentHyperlink, item);
      continue;
    }

    flushHyperlink();
    if (item.kind === 'text') {
      const key = stableStringify(formattingAttrs(item.attributes));
      const nextRun = createTextRun(item.text, item.attributes);
      if (currentRun && currentFormattingKey === key) appendTextRun(currentRun, nextRun);
      else {
        flushRun();
        currentRun = nextRun;
        currentFormattingKey = key;
      }
      continue;
    }

    flushRun();
    const child = ordinaryContentForItem(item);
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

function restoreOriginalRuns(
  content: ParagraphContent[],
  items: InlineItem[],
  boundaries: OriginalRunBoundary[] | undefined
): ParagraphContent[] {
  if (
    !boundaries?.length ||
    !content.every(
      (child) => child.type === 'run' && child.content.every((entry) => entry.type === 'text')
    ) ||
    items.some((item) => item.kind !== 'text' || item.attributes.hyperlink)
  ) {
    return content;
  }
  const fullText = items.map((item) => (item.kind === 'text' ? item.text : '')).join('');
  if (fullText !== boundaries.map((boundary) => boundary.text).join('')) return content;

  let itemIndex = 0;
  let itemOffset = 0;
  const restoredAttrs: Attrs[] = [];
  for (const boundary of boundaries) {
    const expected = marksKeyToYrsAttrs(boundary.marksKey);
    if (!expected) return content;
    restoredAttrs.push(expected);
    let remaining = boundary.text.length;
    while (remaining > 0) {
      const item = items[itemIndex] as TextItem | undefined;
      if (!item) return content;
      if (stableStringify(formattingAttrs(item.attributes)) !== stableStringify(expected)) {
        return content;
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

  return boundaries.map((boundary, index) => {
    // fromProseDoc restores the original segmentation/property-change cache,
    // but non-empty runs keep formatting reconstructed from their live marks.
    // Only empty runs have no PM node and therefore take cached formatting.
    const formatting =
      boundary.text.length === 0
        ? boundary.formatting
        : attrsToTextFormatting(restoredAttrs[index]);
    const run: Run = {
      type: 'run',
      content: runContentForText(boundary.text, formatting ?? {}),
    };
    if (formatting && Object.keys(formatting).length > 0) run.formatting = formatting;
    if (boundary.propertyChanges?.length) run.propertyChanges = boundary.propertyChanges;
    return run;
  });
}

function runTextLength(run: Run): number {
  return run.content.reduce((length, content) => {
    if (content.type === 'text' || content.type === 'instrText')
      return length + content.text.length;
    if (content.type === 'symbol') return length + content.char.length;
    if (
      content.type === 'tab' ||
      content.type === 'softHyphen' ||
      content.type === 'noBreakHyphen'
    ) {
      return length + 1;
    }
    return length;
  }, 0);
}

function paragraphContentLength(content: ParagraphContent): number {
  switch (content.type) {
    case 'run':
      return runTextLength(content);
    case 'hyperlink':
      return content.children.reduce(
        (sum, child) => sum + (child.type === 'run' ? runTextLength(child) : 0),
        0
      );
    case 'simpleField':
      return content.content.reduce((sum, child) => sum + paragraphContentLength(child), 0);
    case 'complexField':
      return content.fieldResult.reduce((sum, run) => sum + runTextLength(run), 0);
    case 'inlineSdt':
      return content.content.reduce((sum, child) => sum + paragraphContentLength(child), 0);
    case 'insertion':
    case 'deletion':
    case 'moveFrom':
    case 'moveTo':
      return content.content.reduce(
        (sum, child) => sum + (child.type === 'run' ? runTextLength(child) : 0),
        0
      );
    case 'mathEquation':
      return content.plainText?.length ?? 0;
    default:
      return 0;
  }
}

function splitTextRun(run: Run, offset: number): [Run | null, Run | null] {
  if (!run.content.every((content) => content.type === 'text')) return [run, null];
  const text = run.content.map((content) => (content.type === 'text' ? content.text : '')).join('');
  const make = (part: string): Run | null =>
    part
      ? {
          type: 'run',
          ...(run.formatting ? { formatting: run.formatting } : {}),
          ...(run.propertyChanges ? { propertyChanges: run.propertyChanges } : {}),
          content: [{ type: 'text', text: part }],
        }
      : null;
  return [make(text.slice(0, offset)), make(text.slice(offset))];
}

function insertBoundaries(
  content: ParagraphContent[],
  boundaries: CommentBoundary[],
  makeMarker: (boundary: CommentBoundary) => ParagraphContent = (boundary) =>
    boundary.kind === 'start'
      ? { type: 'commentRangeStart', id: boundary.id }
      : { type: 'commentRangeEnd', id: boundary.id }
): ParagraphContent[] {
  if (boundaries.length === 0) return content;
  const sorted = [...boundaries].sort(
    (left, right) =>
      left.offset - right.offset ||
      (left.kind === right.kind ? left.id - right.id : left.kind === 'end' ? -1 : 1)
  );
  const result: ParagraphContent[] = [];
  let cursor = 0;
  let boundaryIndex = 0;
  const emit = (offset: number): void => {
    while (boundaryIndex < sorted.length && sorted[boundaryIndex].offset === offset) {
      const boundary = sorted[boundaryIndex++];
      result.push(makeMarker(boundary));
    }
  };

  emit(0);
  for (const item of content) {
    const length = paragraphContentLength(item);
    const inside = sorted
      .slice(boundaryIndex)
      .map((boundary) => boundary.offset)
      .filter((offset) => offset > cursor && offset < cursor + length);
    if (item.type === 'run' && item.content.every((entry) => entry.type === 'text')) {
      let remaining: Run | null = item;
      let localCursor = 0;
      for (const absolute of inside) {
        if (!remaining) break;
        const [left, right] = splitTextRun(remaining, absolute - cursor - localCursor);
        if (left) result.push(left);
        emit(absolute);
        remaining = right;
        localCursor = absolute - cursor;
      }
      if (remaining) result.push(remaining);
    } else {
      result.push(item);
    }
    cursor += length;
    emit(cursor);
  }
  emit(cursor);
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

function paragraphAttrs(properties: Attrs): ParagraphSaveAttrs {
  const attrs = { ...PARAGRAPH_ATTR_DEFAULTS, ...properties } as Attrs;
  attrs.styleId = properties.pStyle ?? null;
  attrs._sectionProperties = properties.sectPr ?? null;
  delete attrs.pStyle;
  delete attrs.sectPr;
  return attrs as ParagraphSaveAttrs;
}

function paragraphFromStory(
  paraId: string,
  properties: Attrs,
  items: InlineItem[],
  commentBoundaries: CommentBoundary[],
  baseParagraph: Paragraph | undefined
): Paragraph {
  const attrs = paragraphAttrs(properties);
  let content = buildParagraphContent(items);
  content = restoreOriginalRuns(
    content,
    items,
    Array.isArray(attrs._originalRunBoundaries)
      ? (attrs._originalRunBoundaries as OriginalRunBoundary[])
      : undefined
  );
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
    formatting: paragraphAttrsToFormatting(attrs),
    content,
  };
  if (baseParagraph?.renderedPageBreakBefore) paragraph.renderedPageBreakBefore = true;

  const pPrIns = trackedInfo(properties.pPrIns, true);
  const pPrDel = trackedInfo(properties.pPrDel, true);
  if (pPrIns) paragraph.pPrIns = pPrIns;
  if (pPrDel) paragraph.pPrDel = pPrDel;
  if (Array.isArray(properties.pPrChange) && properties.pPrChange.length > 0) {
    paragraph.propertyChanges = properties.pPrChange as Paragraph['propertyChanges'];
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

function inferTableBorders(rows: TableRow[]): TableBorders | undefined {
  for (const row of rows) {
    for (const cell of row.cells) {
      const borders = cell.formatting?.borders;
      if (!borders) continue;
      const base =
        borders.top ||
        borders.left ||
        borders.right ||
        borders.bottom ||
        borders.insideH ||
        borders.insideV;
      if (!base) return undefined;
      return {
        top: borders.top ?? base,
        bottom: borders.bottom ?? base,
        left: borders.left ?? base,
        right: borders.right ?? base,
        insideH: borders.insideH ?? borders.bottom ?? base,
        insideV: borders.insideV ?? borders.right ?? base,
      };
    }
  }
  return undefined;
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
  const attrs = {
    ...TABLE_CELL_ATTR_DEFAULTS,
    ...(payload.tcPr ?? {}),
  } as unknown as TableCellSaveAttrs;
  const content =
    payload.story && context.storyIds.has(payload.story)
      ? context
          .storyToBlocks(payload.story)
          .filter(
            (block): block is Paragraph | Table =>
              block.type === 'paragraph' || block.type === 'table'
          )
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
  let formatting = tableAttrsToFormatting(attrs);
  if (!formatting?.borders) {
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

function collectBaseParagraphs(document: Document): Map<string, Paragraph> {
  const paragraphs = new Map<string, Paragraph>();
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        if (block.paraId && !paragraphs.has(block.paraId)) paragraphs.set(block.paraId, block);
      } else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.content);
      } else {
        visit(block.content);
      }
    }
  };
  visit(document.package.document.content);
  for (const part of document.package.headers?.values() ?? []) visit(part.content);
  for (const part of document.package.footers?.values() ?? []) visit(part.content);
  for (const note of document.package.footnotes ?? []) visit(note.content);
  for (const note of document.package.endnotes ?? []) visit(note.content);
  return paragraphs;
}

function collectBaseStories(document: Document): Map<string, readonly BlockContent[]> {
  const stories = new Map<string, readonly BlockContent[]>();
  const visit = (storyId: string, blocks: readonly BlockContent[]): void => {
    stories.set(storyId, blocks);
    let tableIndex = 0;
    for (const block of blocks) {
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
  for (const [rId, part] of document.package.headers ?? []) visit(`hf:${rId}`, part.content);
  for (const [rId, part] of document.package.footers ?? []) visit(`hf:${rId}`, part.content);
  for (const note of document.package.footnotes ?? []) visit(`fn:${note.id}`, note.content);
  return stories;
}

function commentRanges(
  session: YrsSession,
  comments: readonly Comment[] | undefined
): Map<string, Array<{ id: number; start: number; end: number }>> {
  const byStory = new Map<string, Array<{ id: number; start: number; end: number }>>();
  for (const comment of comments ?? []) {
    let anchors: ReturnType<YrsSession['resolveComment']>;
    try {
      anchors = session.resolveComment(String(comment.id));
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

class SaveContext {
  readonly storyIds: Set<string>;
  private readonly baseParagraphs: Map<string, Paragraph>;
  private readonly baseStories: Map<string, readonly BlockContent[]>;
  private readonly comments: Map<string, Array<{ id: number; start: number; end: number }>>;

  constructor(
    private readonly session: YrsSession,
    base: Document
  ) {
    this.storyIds = new Set(session.storyIds());
    this.baseParagraphs = collectBaseParagraphs(base);
    this.baseStories = collectBaseStories(base);
    this.comments = commentRanges(session, base.package.document.comments);
  }

  storyToBlocks(storyId: string): BlockContent[] {
    const blocks: BlockContent[] = [];
    const baseBlocks = this.baseStories.get(storyId);
    const segments = this.session.storySegments(storyId);
    const storyComments = this.comments.get(storyId) ?? [];
    let items: InlineItem[] = [];
    let paragraphStart = 0;
    let paragraphIndex = 0;
    let storyOffset = 0;

    const paragraphCommentBoundaries = (end: number): CommentBoundary[] => {
      const boundaries: CommentBoundary[] = [];
      for (const range of storyComments) {
        if (range.start >= paragraphStart && range.start <= end) {
          boundaries.push({ id: range.id, kind: 'start', offset: range.start - paragraphStart });
        }
        if (range.end >= paragraphStart && range.end <= end) {
          boundaries.push({ id: range.id, kind: 'end', offset: range.end - paragraphStart });
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

    for (const segment of segments) {
      if (segment.kind === 'text') {
        pushText(segment.text, segment.attributes);
        storyOffset += segment.text.length;
        continue;
      }
      if (segment.kind === 'pilcrow') {
        const generatedId = `${storyId}:p${paragraphIndex}`;
        const savedParaId =
          segment.paraId === generatedId && !this.baseParagraphs.has(segment.paraId)
            ? ''
            : segment.paraId;
        blocks.push(
          paragraphFromStory(
            savedParaId,
            segment.properties,
            items,
            paragraphCommentBoundaries(storyOffset),
            this.baseParagraphs.get(segment.paraId)
          )
        );
        items = [];
        paragraphIndex += 1;
        storyOffset += 1;
        paragraphStart = storyOffset;
        continue;
      }

      if (segment.embedKind === 'table') {
        blocks.push(tableFromPayload(this, segment.payload as TablePayload));
      } else if (segment.embedKind === 'blockSdt') {
        const childStory = asString(segment.payload.story);
        let properties = sdtAttrsToProps(segment.payload);
        let content =
          childStory && this.storyIds.has(childStory) ? this.storyToBlocks(childStory) : [];
        const authoredValue = contentControlValue(segment.payload.value);
        if (authoredValue) {
          try {
            const applied = applyContentControlValue(properties, authoredValue);
            properties = applied.properties;
            content = applied.content;
          } catch {
            // Retain the child story if the authored value is invalid.
          }
        }
        blocks.push({
          type: 'blockSdt',
          properties,
          content,
        });
      } else if (segment.embedKind === 'opaque') {
        const blob = asObject(segment.payload.blob);
        if (blob?.type === 'pageBreak') blocks.push(pageBreakParagraph());
        else {
          // Opaque authored blocks have no editable yrs sub-story yet. Carry a
          // same-position base block when it is still structurally compatible;
          // this keeps block SDTs lossless until they become native.
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
      storyOffset += 1;
    }

    // Defensive recovery for malformed/legacy stories without a final pilcrow.
    if (items.length > 0) {
      blocks.push({ type: 'paragraph', content: buildParagraphContent(items) });
    }
    return blocks;
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
}

export function yrsToDocument(
  session: YrsSession,
  base: Document,
  options: YrsToDocumentOptions = {}
): Document {
  const context = new SaveContext(session, base);
  const shouldProject = (storyId: string): boolean =>
    options.storyIds === undefined || options.storyIds.has(storyId);
  const bodyContent = context.storyIds.has('body') && shouldProject('body')
    ? context.storyToBlocks('body')
    : base.package.document.content;

  let headers = base.package.headers;
  if (
    headers &&
    (options.storyIds === undefined || [...headers.keys()].some((rId) => shouldProject(`hf:${rId}`)))
  ) {
    headers = new Map(
      [...headers].map(([rId, part]) => {
        const storyId = `hf:${rId}`;
        return [
          rId,
          context.storyIds.has(storyId) && shouldProject(storyId)
            ? { ...part, content: context.storyToBlocks(storyId) }
            : part,
        ];
      })
    );
  }

  let footers = base.package.footers;
  if (
    footers &&
    (options.storyIds === undefined || [...footers.keys()].some((rId) => shouldProject(`hf:${rId}`)))
  ) {
    footers = new Map(
      [...footers].map(([rId, part]) => {
        const storyId = `hf:${rId}`;
        return [
          rId,
          context.storyIds.has(storyId) && shouldProject(storyId)
            ? { ...part, content: context.storyToBlocks(storyId) }
            : part,
        ];
      })
    );
  }

  const shouldProjectFootnotes =
    options.storyIds === undefined ||
    base.package.footnotes?.some((note) => shouldProject(`fn:${note.id}`));
  const footnotes = shouldProjectFootnotes
    ? base.package.footnotes?.map((note) => {
        const storyId = `fn:${note.id}`;
        return context.storyIds.has(storyId) && shouldProject(storyId)
          ? { ...note, content: context.storyToBlocks(storyId), verbatimXml: undefined }
          : note;
      })
    : base.package.footnotes;

  return {
    ...base,
    package: {
      ...base.package,
      document: {
        ...base.package.document,
        content: bodyContent,
      },
      ...(headers ? { headers } : {}),
      ...(footers ? { footers } : {}),
      ...(footnotes ? { footnotes } : {}),
    },
  };
}
