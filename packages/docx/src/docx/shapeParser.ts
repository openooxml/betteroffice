/**
 * Public XmlElement shape compatibility adapter.
 *
 * Rust S9 owns package drawing parsing. These small helpers remain because
 * the published ./docx surface accepts xml-js elements directly.
 */

import type { Paragraph, Shape, ShapeTextBody, ShapeType } from '../types/document';
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
import { parseShapeType, parseTransform, rotToDegrees } from '@betteroffice/drawingml';
import { emuToPixels } from '../utils/units';

export { emuToPixels } from '../utils/units';

function directChild(parent: XmlElement | null, localName: string): XmlElement | null {
  return (
    getChildElements(parent).find((child) => getLocalName(child.name ?? '') === localName) ?? null
  );
}

function firstDescendant(
  parent: XmlElement | null,
  localName: string,
  depth = 0
): XmlElement | null {
  if (!parent || depth > 32) return null;
  for (const child of getChildElements(parent)) {
    if (getLocalName(child.name ?? '') === localName) return child;
    const nested = firstDescendant(child, localName, depth + 1);
    if (nested) return nested;
  }
  return null;
}

function booleanAttribute(element: XmlElement | null, name: string): boolean | undefined {
  const value = getAttribute(element, null, name);
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return undefined;
}

function numberAttribute(element: XmlElement | null, name: string): number | undefined {
  const value = parseNumericAttribute(element, null, name);
  return value !== undefined && Number.isFinite(value) ? value : undefined;
}

function placeholderParagraphs(txbxContent: XmlElement | null): Paragraph[] {
  if (!txbxContent) return [];
  return findChildrenByLocalName(txbxContent, 'p').map(() => ({
    type: 'paragraph',
    formatting: {},
    content: [],
  }));
}

function parseShapeTextBody(bodyPr: XmlElement | null, txbxContent: XmlElement | null): ShapeTextBody {
  const anchorMap: Record<string, ShapeTextBody['anchor']> = {
    t: 'top',
    ctr: 'middle',
    b: 'bottom',
    dist: 'distributed',
    just: 'justified',
  };
  const anchor = getAttribute(bodyPr, null, 'anchor');
  const vertical = getAttribute(bodyPr, null, 'vert');
  const margins = {
    left: numberAttribute(bodyPr, 'lIns'),
    right: numberAttribute(bodyPr, 'rIns'),
    top: numberAttribute(bodyPr, 'tIns'),
    bottom: numberAttribute(bodyPr, 'bIns'),
  };
  const body: ShapeTextBody = { content: placeholderParagraphs(txbxContent) };
  if (vertical && vertical !== 'horz') body.vertical = true;
  const rotation = rotToDegrees(getAttribute(bodyPr, null, 'rot'));
  if (rotation !== undefined) body.rotation = rotation;
  if (anchor && anchorMap[anchor]) body.anchor = anchorMap[anchor];
  const anchorCenter = booleanAttribute(bodyPr, 'anchorCtr');
  if (anchorCenter !== undefined) body.anchorCenter = anchorCenter;
  if (directChild(bodyPr, 'noAutofit')) body.autoFit = 'none';
  else if (directChild(bodyPr, 'normAutofit')) body.autoFit = 'normal';
  else if (directChild(bodyPr, 'spAutoFit')) body.autoFit = 'shape';
  if (Object.values(margins).some((value) => value !== undefined)) body.margins = margins;
  return body;
}

export function parseShape(node: XmlElement): Shape {
  const spPr = directChild(node, 'spPr') ?? firstDescendant(node, 'spPr');
  const cNvPr = firstDescendant(node, 'cNvPr');
  const bodyPr = directChild(node, 'bodyPr') ?? firstDescendant(directChild(node, 'txBody'), 'bodyPr');
  const txbxContent = firstDescendant(node, 'txbxContent');
  const parsedTransform = parseTransform(spPr ? findByFullName(spPr, 'a:xfrm') : null);
  const shape: Shape = {
    type: 'shape',
    shapeType: parseShapeType(spPr),
    size: parsedTransform.size,
  };
  if (parsedTransform.offset) shape.offset = parsedTransform.offset;
  if (parsedTransform.transform) shape.transform = parsedTransform.transform;
  const fill = parseFill(spPr);
  if (fill) shape.fill = fill;
  const outline = parseOutline(spPr);
  if (outline) shape.outline = outline;
  if (cNvPr) {
    shape.id = getAttribute(cNvPr, null, 'id') ?? undefined;
    shape.name = getAttribute(cNvPr, null, 'name') ?? undefined;
    shape.title = getAttribute(cNvPr, null, 'title') ?? undefined;
    shape.description = getAttribute(cNvPr, null, 'descr') ?? undefined;
    shape.hidden = booleanAttribute(cNvPr, 'hidden');
    shape.decorative = booleanAttribute(cNvPr, 'decorative');
  }
  if (txbxContent || bodyPr) shape.textBody = parseShapeTextBody(bodyPr, txbxContent);
  return shape;
}

function drawingContainer(drawingEl: XmlElement): XmlElement | null {
  return (
    getChildElements(drawingEl).find((child) => {
      const name = getLocalName(child.name ?? '');
      return name === 'inline' || name === 'anchor';
    }) ?? null
  );
}

export function parseShapeFromDrawing(drawingEl: XmlElement): Shape | null {
  const container = drawingContainer(drawingEl);
  if (!container) return null;
  const graphicData = findByFullName(container, 'a:graphicData');
  if (!graphicData) return null;
  const root =
    getChildElements(graphicData).find((child) =>
      ['wsp', 'sp', 'cxnSp'].includes(getLocalName(child.name ?? ''))
    ) ?? firstDescendant(graphicData, 'wsp');
  if (!root) return null;
  const shape = parseShape(root);
  const extent = findByFullName(container, 'wp:extent');
  if (extent) {
    shape.size = {
      width: numberAttribute(extent, 'cx') ?? 0,
      height: numberAttribute(extent, 'cy') ?? 0,
    };
  }
  if (getLocalName(container.name ?? '') === 'anchor') {
    shape.position = parseAnchorPosition(container);
    shape.wrap = parseAnchorWrap(container);
  }
  const docPr = findByFullName(container, 'wp:docPr');
  if (docPr) {
    shape.id = getAttribute(docPr, null, 'id') ?? shape.id;
    shape.name = getAttribute(docPr, null, 'name') ?? shape.name;
    shape.title = getAttribute(docPr, null, 'title') ?? shape.title;
    shape.description = getAttribute(docPr, null, 'descr') ?? shape.description;
    shape.hidden = booleanAttribute(docPr, 'hidden') ?? shape.hidden;
    shape.decorative = booleanAttribute(docPr, 'decorative') ?? shape.decorative;
  }
  const effectExtent = findByFullName(container, 'wp:effectExtent');
  if (effectExtent) {
    shape.effectExtent = {
      left: numberAttribute(effectExtent, 'l'),
      top: numberAttribute(effectExtent, 't'),
      right: numberAttribute(effectExtent, 'r'),
      bottom: numberAttribute(effectExtent, 'b'),
    };
  }
  shape.relativeHeight = shape.position?.relativeHeight;
  return shape;
}

export function isShapeDrawing(drawingEl: XmlElement): boolean {
  const container = drawingContainer(drawingEl);
  const graphicData = container ? findByFullName(container, 'a:graphicData') : null;
  if (!graphicData) return false;
  return getChildElements(graphicData).some((child) =>
    ['wsp', 'wgp', 'wpc', 'spTree', 'grpSp', 'sp', 'cxnSp'].includes(
      getLocalName(child.name ?? '')
    )
  );
}

export function isLineShape(shape: Shape): boolean {
  const lineTypes: ShapeType[] = [
    'line',
    'straightConnector1',
    'bentConnector2',
    'bentConnector3',
    'bentConnector4',
    'bentConnector5',
    'curvedConnector2',
    'curvedConnector3',
    'curvedConnector4',
    'curvedConnector5',
  ];
  return lineTypes.includes(shape.shapeType);
}

export function isTextBoxShape(shape: Shape): boolean {
  return shape.shapeType === 'textBox' || Boolean(shape.textBody?.content.length);
}

export function hasTextContent(shape: Shape): boolean {
  return Boolean(shape.textBody?.content.length);
}

export function getShapeWidthPx(shape: Shape): number {
  return emuToPixels(shape.size.width);
}

export function getShapeHeightPx(shape: Shape): number {
  return emuToPixels(shape.size.height);
}

export function getShapeDimensionsPx(shape: Shape): { width: number; height: number } {
  return { width: getShapeWidthPx(shape), height: getShapeHeightPx(shape) };
}

export function isFloatingShape(shape: Shape): boolean {
  return shape.position !== undefined || shape.wrap !== undefined;
}

export function hasFill(shape: Shape): boolean {
  return shape.fill !== undefined && shape.fill.type !== 'none';
}

export function hasOutline(shape: Shape): boolean {
  return shape.outline !== undefined;
}

export function getOutlineWidthPx(shape: Shape): number {
  return shape.outline?.width ? emuToPixels(shape.outline.width) : 0;
}

export function resolveFillColor(shape: Shape): string | undefined {
  return shape.fill?.type === 'solid' ? resolveColorValueToHex(shape.fill.color) : undefined;
}

export function resolveOutlineColor(shape: Shape): string | undefined {
  return shape.outline?.color ? resolveColorValueToHex(shape.outline.color) : undefined;
}
