/** Public XmlElement text-box compatibility adapter and pure model helpers. */

import type {
  ImagePosition,
  ImageSize,
  ImageWrap,
  MediaFile,
  Paragraph,
  RelationshipMap,
  ShapeTextBodyProperties,
  Table,
  TextBox,
  Theme,
} from '../types/document';
import type { NumberingMap } from './numberingParser';
import type { StyleMap } from './styleParser';
import {
  findByFullName,
  findChildrenByLocalName,
  getAttribute,
  getChildElements,
  getLocalName,
  parseNumericAttribute,
  type XmlElement,
} from './xmlParser';
import {
  parseAnchorPosition,
  parseAnchorWrap,
  parseFill,
  parseOutline,
  resolveColorValueToHex,
} from './drawingUtils';
import { rotToDegrees } from '@betteroffice/drawingml';
import { emuToPixels } from '../utils/units';

export { emuToPixels } from '../utils/units';

const DEFAULT_MARGIN_EMU = 91440;

function directChild(parent: XmlElement | null, localName: string): XmlElement | null {
  return (
    getChildElements(parent).find((child) => getLocalName(child.name ?? '') === localName) ?? null
  );
}

function numberAttribute(element: XmlElement | null, name: string): number | undefined {
  const value = parseNumericAttribute(element, null, name);
  return value !== undefined && Number.isFinite(value) ? value : undefined;
}

function parseBodyProperties(bodyPr: XmlElement | null): {
  margins?: TextBox['margins'];
  bodyProperties?: ShapeTextBodyProperties;
} {
  if (!bodyPr) return {};
  const margins = {
    left: numberAttribute(bodyPr, 'lIns'),
    right: numberAttribute(bodyPr, 'rIns'),
    top: numberAttribute(bodyPr, 'tIns'),
    bottom: numberAttribute(bodyPr, 'bIns'),
  };
  const hasMargins = Object.values(margins).some((value) => value !== undefined);
  const verticalMap: Record<string, NonNullable<ShapeTextBodyProperties['vertical']>> = {
    horz: 'horizontal',
    vert: 'vertical',
    vert270: 'vertical270',
    wordArtVert: 'wordArtVertical',
    wordArtVertRtl: 'wordArtVertical',
    eaVert: 'eastAsianVertical',
    mongolianVert: 'mongolianVertical',
  };
  const anchorMap: Record<string, NonNullable<ShapeTextBodyProperties['anchor']>> = {
    t: 'top',
    ctr: 'middle',
    b: 'bottom',
    dist: 'distributed',
    just: 'justified',
  };
  const vertical = getAttribute(bodyPr, null, 'vert');
  const anchor = getAttribute(bodyPr, null, 'anchor');
  const normAutofit = directChild(bodyPr, 'normAutofit');
  const bodyProperties: ShapeTextBodyProperties = {
    vertical: vertical ? verticalMap[vertical] : undefined,
    rotation: rotToDegrees(getAttribute(bodyPr, null, 'rot')),
    upright: getAttribute(bodyPr, null, 'upright') === '1' || undefined,
    anchor: anchor ? anchorMap[anchor] : undefined,
    anchorCenter: getAttribute(bodyPr, null, 'anchorCtr') === '1' || undefined,
    columns: numberAttribute(bodyPr, 'numCol'),
    columnSpacing: numberAttribute(bodyPr, 'spcCol'),
    margins: hasMargins ? margins : undefined,
    autoFit: directChild(bodyPr, 'noAutofit')
      ? 'none'
      : normAutofit
        ? 'normal'
        : directChild(bodyPr, 'spAutoFit')
          ? 'shape'
          : undefined,
    fontScale: numberAttribute(normAutofit, 'fontScale'),
    lineSpacingReduction: numberAttribute(normAutofit, 'lnSpcReduction'),
  };
  return {
    ...(hasMargins ? { margins } : {}),
    ...(Object.values(bodyProperties).some((value) => value !== undefined)
      ? { bodyProperties }
      : {}),
  };
}

export function extractTextBoxContentElements(txbxContent: XmlElement | null): {
  paragraphElements: XmlElement[];
  tableElements: XmlElement[];
} {
  return {
    paragraphElements: findChildrenByLocalName(txbxContent, 'p'),
    tableElements: findChildrenByLocalName(txbxContent, 'tbl'),
  };
}

export type ParagraphParserFn = (
  node: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels?: RelationshipMap | null
) => Paragraph;

export type TableParserFn = (
  node: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels?: RelationshipMap | null,
  media?: Map<string, MediaFile>
) => Table;

export function parseTextBoxContent(
  txbxContent: XmlElement | null,
  parseParagraph: ParagraphParserFn,
  parseTable: TableParserFn | null,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels?: RelationshipMap | null,
  _media?: Map<string, MediaFile>
): Paragraph[] {
  if (!txbxContent) return [];
  const paragraphs: Paragraph[] = [];
  for (const child of getChildElements(txbxContent)) {
    const name = getLocalName(child.name ?? '');
    if (name === 'p') paragraphs.push(parseParagraph(child, styles, theme, numbering, rels));
    else if (name === 'tbl' && parseTable) {
      parseTable(child, styles, theme, numbering, rels, _media);
    }
  }
  return paragraphs;
}

function drawingParts(drawingEl: XmlElement): {
  container: XmlElement | null;
  shape: XmlElement | null;
} {
  const container =
    getChildElements(drawingEl).find((child) =>
      ['inline', 'anchor'].includes(getLocalName(child.name ?? ''))
    ) ?? null;
  const graphic = container ? findByFullName(container, 'a:graphic') : null;
  const graphicData = graphic ? findByFullName(graphic, 'a:graphicData') : null;
  return { container, shape: graphicData ? findByFullName(graphicData, 'wps:wsp') : null };
}

export function isTextBoxDrawing(drawingEl: XmlElement): boolean {
  const { shape } = drawingParts(drawingEl);
  return shape ? isShapeTextBox(shape) : false;
}

export function isShapeTextBox(wsp: XmlElement): boolean {
  return findByFullName(wsp, 'wps:txbx') !== null;
}

function buildTextBox(
  wsp: XmlElement,
  size: ImageSize,
  position?: ImagePosition,
  wrap?: ImageWrap
): TextBox | null {
  if (!isShapeTextBox(wsp)) return null;
  const spPr = directChild(wsp, 'spPr');
  const bodyPr = directChild(wsp, 'bodyPr');
  const cNvPr = directChild(wsp, 'cNvPr');
  const body = parseBodyProperties(bodyPr);
  const textBox: TextBox = { type: 'textBox', size, content: [] };
  textBox.id = getAttribute(cNvPr, null, 'id') ?? undefined;
  const fill = parseFill(spPr);
  if (fill) textBox.fill = fill;
  const outline = parseOutline(spPr);
  if (outline) textBox.outline = outline;
  if (body.margins) textBox.margins = body.margins;
  if (body.bodyProperties) textBox.bodyProperties = body.bodyProperties;
  if (position) textBox.position = position;
  if (wrap) textBox.wrap = wrap;
  return textBox;
}

export function parseTextBox(drawingEl: XmlElement): TextBox | null {
  const { container, shape } = drawingParts(drawingEl);
  if (!container || !shape) return null;
  const extent = findByFullName(container, 'wp:extent');
  const size = {
    width: numberAttribute(extent, 'cx') ?? 0,
    height: numberAttribute(extent, 'cy') ?? 0,
  };
  const anchored = getLocalName(container.name ?? '') === 'anchor';
  const textBox = buildTextBox(
    shape,
    size,
    anchored ? parseAnchorPosition(container) : undefined,
    anchored ? parseAnchorWrap(container) : undefined
  );
  if (textBox) {
    textBox.id = getAttribute(findByFullName(container, 'wp:docPr'), null, 'id') ?? textBox.id;
  }
  return textBox;
}

export function getTextBoxContentElement(wsp: XmlElement): XmlElement | null {
  return findByFullName(findByFullName(wsp, 'wps:txbx'), 'w:txbxContent');
}

export function parseTextBoxFromShape(
  wsp: XmlElement,
  size: ImageSize,
  position?: ImagePosition,
  wrap?: ImageWrap
): TextBox | null {
  return buildTextBox(wsp, size, position, wrap);
}

export function getTextBoxWidthPx(textBox: TextBox): number {
  return emuToPixels(textBox.size.width);
}

export function getTextBoxHeightPx(textBox: TextBox): number {
  return emuToPixels(textBox.size.height);
}

export function getTextBoxDimensionsPx(textBox: TextBox): { width: number; height: number } {
  return { width: getTextBoxWidthPx(textBox), height: getTextBoxHeightPx(textBox) };
}

export function getTextBoxMarginsPx(textBox: TextBox): {
  top: number;
  bottom: number;
  left: number;
  right: number;
} {
  return {
    top: emuToPixels(textBox.margins?.top ?? DEFAULT_MARGIN_EMU),
    bottom: emuToPixels(textBox.margins?.bottom ?? DEFAULT_MARGIN_EMU),
    left: emuToPixels(textBox.margins?.left ?? DEFAULT_MARGIN_EMU),
    right: emuToPixels(textBox.margins?.right ?? DEFAULT_MARGIN_EMU),
  };
}

export function isFloatingTextBox(textBox: TextBox): boolean {
  return textBox.position !== undefined || textBox.wrap !== undefined;
}

export function hasTextBoxFill(textBox: TextBox): boolean {
  return textBox.fill !== undefined && textBox.fill.type !== 'none';
}

export function hasTextBoxOutline(textBox: TextBox): boolean {
  return textBox.outline !== undefined;
}

export function hasTextBoxContent(textBox: TextBox): boolean {
  return textBox.content.length > 0;
}

export function getTextBoxText(textBox: TextBox): string {
  return textBox.content
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

export function resolveTextBoxFillColor(textBox: TextBox): string | undefined {
  return textBox.fill?.type === 'solid'
    ? resolveColorValueToHex(textBox.fill.color)
    : undefined;
}

export function resolveTextBoxOutlineColor(textBox: TextBox): string | undefined {
  return textBox.outline?.color ? resolveColorValueToHex(textBox.outline.color) : undefined;
}

export function getTextBoxOutlineWidthPx(textBox: TextBox): number {
  return textBox.outline?.width ? emuToPixels(textBox.outline.width) : 0;
}
