/**
 * Public XmlElement image compatibility adapter.
 *
 * Rust S9 owns package drawing/media parsing. This adapter remains for the
 * published leaf API and pure image-model helpers.
 */

import type {
  Image,
  ImageCrop,
  ImagePadding,
  ImageSize,
  ImageTransform,
  ImageWrap,
  MediaFile,
  RelationshipMap,
} from '../types/document';
import {
  findByFullName,
  getAttribute,
  getChildElements,
  getLocalName,
  parseNumericAttribute,
  type XmlElement,
} from './xmlParser';
import { parseAnchorPosition, parseAnchorWrap, parseOutline } from './drawingUtils';
import { resolveTarget } from './relsParser';
import { sanitizeHref } from '../utils/sanitizeHref';
import { emuToPixels } from '../utils/units';

export { emuToPixels, pixelsToEmu } from '../utils/units';

function numberAttribute(element: XmlElement | null, name: string): number | undefined {
  const value = parseNumericAttribute(element, null, name);
  return value !== undefined && Number.isFinite(value) ? value : undefined;
}

function booleanAttribute(element: XmlElement | null, name: string): boolean | undefined {
  const value = getAttribute(element, null, name);
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return undefined;
}

function drawingContainer(drawingEl: XmlElement): XmlElement | null {
  return (
    getChildElements(drawingEl).find((child) => {
      const name = getLocalName(child.name ?? '');
      return name === 'inline' || name === 'anchor';
    }) ?? null
  );
}

function pictureParts(container: XmlElement): {
  picture: XmlElement | null;
  blipFill: XmlElement | null;
  blip: XmlElement | null;
} {
  const graphic = findByFullName(container, 'a:graphic');
  const graphicData = graphic ? findByFullName(graphic, 'a:graphicData') : null;
  const picture = graphicData ? findByFullName(graphicData, 'pic:pic') : null;
  const blipFill = picture ? findByFullName(picture, 'pic:blipFill') : null;
  return { picture, blipFill, blip: blipFill ? findByFullName(blipFill, 'a:blip') : null };
}

function parseExtent(container: XmlElement): ImageSize {
  const extent = findByFullName(container, 'wp:extent');
  return {
    width: numberAttribute(extent, 'cx') ?? 0,
    height: numberAttribute(extent, 'cy') ?? 0,
  };
}

function parsePadding(container: XmlElement): ImagePadding | undefined {
  const extent = findByFullName(container, 'wp:effectExtent');
  if (!extent) return undefined;
  const padding: ImagePadding = {
    left: numberAttribute(extent, 'l') ?? 0,
    top: numberAttribute(extent, 't') ?? 0,
    right: numberAttribute(extent, 'r') ?? 0,
    bottom: numberAttribute(extent, 'b') ?? 0,
  };
  return Object.values(padding).some(Boolean) ? padding : undefined;
}

function parseTransform(container: XmlElement): ImageTransform | undefined {
  const { picture } = pictureParts(container);
  const spPr = picture ? findByFullName(picture, 'pic:spPr') : null;
  const xfrm = spPr ? findByFullName(spPr, 'a:xfrm') : null;
  if (!xfrm) return undefined;
  const rawRotation = getAttribute(xfrm, null, 'rot');
  const rotation = rawRotation === null ? undefined : Number(rawRotation) / 60000;
  const flipH = booleanAttribute(xfrm, 'flipH');
  const flipV = booleanAttribute(xfrm, 'flipV');
  if (!Number.isFinite(rotation) && flipH !== true && flipV !== true) return undefined;
  return {
    ...(Number.isFinite(rotation) ? { rotation } : {}),
    ...(flipH === true ? { flipH: true } : {}),
    ...(flipV === true ? { flipV: true } : {}),
  };
}

function parseCrop(blipFill: XmlElement): ImageCrop | undefined {
  const srcRect = findByFullName(blipFill, 'a:srcRect');
  if (!srcRect) return undefined;
  const fraction = (name: string): number | undefined => {
    const value = numberAttribute(srcRect, name);
    return value === undefined || value === 0 ? undefined : value / 100000;
  };
  const crop: ImageCrop = {
    left: fraction('l'),
    top: fraction('t'),
    right: fraction('r'),
    bottom: fraction('b'),
  };
  return Object.values(crop).some((value) => value !== undefined) ? crop : undefined;
}

function parseOpacity(blip: XmlElement): number | undefined {
  const alpha = findByFullName(blip, 'a:alphaModFix');
  const amount = numberAttribute(alpha, 'amt');
  if (amount === undefined || amount >= 100000) return undefined;
  return Math.max(0, Math.min(1, amount / 100000));
}

function parseInlineWrap(container: XmlElement): ImageWrap {
  const wrap: ImageWrap = { type: 'inline' };
  for (const [attribute, key] of [
    ['distT', 'distT'],
    ['distB', 'distB'],
    ['distL', 'distL'],
    ['distR', 'distR'],
  ] as const) {
    const value = numberAttribute(container, attribute);
    if (value !== undefined) wrap[key] = value;
  }
  return wrap;
}

function normalizeMediaPath(target: string): string {
  const path = target.replace(/^\/+/, '');
  if (path.startsWith('word/')) return path;
  return path.startsWith('media/') ? `word/${path}` : `word/${path}`;
}

function mediaForTarget(
  target: string,
  media: Map<string, MediaFile> | undefined
): MediaFile | undefined {
  if (!media) return undefined;
  const candidates = new Set([
    target.replace(/^\/+/, ''),
    normalizeMediaPath(target),
    `word/${target.replace(/^\/+/, '')}`,
  ]);
  for (const [key, value] of media) {
    if ([...candidates].some((candidate) => candidate.toLowerCase() === key.toLowerCase())) {
      return value;
    }
  }
  return undefined;
}

function mimeTypeFor(path: string): string {
  const extension = path.split('.').pop()?.toLowerCase();
  return (
    {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      tif: 'image/tiff',
      tiff: 'image/tiff',
      webp: 'image/webp',
      svg: 'image/svg+xml',
      emf: 'image/x-emf',
      wmf: 'image/x-wmf',
    }[extension ?? ''] ?? 'application/octet-stream'
  );
}

function applyRelationshipData(
  image: Image,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined
): void {
  const relationship = rels?.get(image.rId);
  if (!relationship?.target) return;
  const file = mediaForTarget(relationship.target, media);
  image.filename = relationship.target.split('/').pop();
  image.mimeType = file?.mimeType ?? mimeTypeFor(relationship.target);
  image.src = file?.dataUrl ?? file?.base64;
}

function applyDocumentProperties(
  image: Image,
  container: XmlElement,
  rels: RelationshipMap | undefined
): void {
  const docPr = findByFullName(container, 'wp:docPr');
  if (!docPr) return;
  image.id = getAttribute(docPr, null, 'id') ?? undefined;
  image.alt = getAttribute(docPr, null, 'descr') ?? undefined;
  image.title = getAttribute(docPr, null, 'title') ?? undefined;
  if (booleanAttribute(docPr, 'decorative') || findByFullName(docPr, 'adec:decorative')) {
    image.decorative = true;
  }
  const hyperlink = findByFullName(docPr, 'a:hlinkClick');
  const hyperlinkId = getAttribute(hyperlink, 'r', 'id');
  if (hyperlinkId && rels) {
    const href = sanitizeHref(resolveTarget(rels, hyperlinkId));
    if (href) {
      image.hlinkHref = href;
      image.hyperlink = {
        href,
        tooltip: getAttribute(hyperlink, null, 'tooltip') ?? undefined,
        target: getAttribute(hyperlink, null, 'tgtFrame') ?? undefined,
        history: getAttribute(hyperlink, null, 'history') !== '0',
      };
    }
  }
}

function parsePicture(
  container: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined
): Image | null {
  const { picture, blipFill, blip } = pictureParts(container);
  if (!picture || !blipFill || !blip) return null;
  const rId =
    getAttribute(blip, 'r', 'embed') ??
    getAttribute(blip, null, 'embed') ??
    getAttribute(blip, 'r', 'link') ??
    '';
  const anchored = getLocalName(container.name ?? '') === 'anchor';
  const image: Image = {
    type: 'image',
    rId,
    size: parseExtent(container),
    wrap: anchored ? (parseAnchorWrap(container) ?? { type: 'inFront' }) : parseInlineWrap(container),
  };
  if (anchored) {
    image.position = parseAnchorPosition(container);
    image.layoutInCell = booleanAttribute(container, 'layoutInCell');
    image.allowOverlap = booleanAttribute(container, 'allowOverlap');
  }
  const padding = parsePadding(container);
  if (padding) {
    image.padding = padding;
    image.effectExtent = { ...padding };
  }
  const transform = parseTransform(container);
  if (transform) image.transform = transform;
  const crop = parseCrop(blipFill);
  if (crop) image.crop = crop;
  const opacity = parseOpacity(blip);
  if (opacity !== undefined) image.opacity = opacity;
  const spPr = findByFullName(picture, 'pic:spPr');
  const outline = parseOutline(spPr);
  if (outline) image.outline = outline;
  applyRelationshipData(image, rels, media);
  applyDocumentProperties(image, container, rels);
  return image;
}

export function parseDrawing(
  drawingEl: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined
): Image | null {
  const container = drawingContainer(drawingEl);
  return container ? parsePicture(container, rels, media) : null;
}

export function parseImage(
  node: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined
): Image | null {
  return parseDrawing(node, rels, media);
}

export function isInlineImage(image: Image): boolean {
  return image.wrap.type === 'inline';
}

export function isFloatingImage(image: Image): boolean {
  return !isInlineImage(image);
}

export function isBehindText(image: Image): boolean {
  return image.wrap.type === 'behind';
}

export function isInFrontOfText(image: Image): boolean {
  return image.wrap.type === 'inFront';
}

export function getImageWidthPx(image: Image): number {
  return emuToPixels(image.size.width);
}

export function getImageHeightPx(image: Image): number {
  return emuToPixels(image.size.height);
}

export function getImageDimensionsPx(image: Image): { width: number; height: number } {
  return { width: getImageWidthPx(image), height: getImageHeightPx(image) };
}

export function hasAltText(image: Image): boolean {
  return Boolean(image.alt?.trim());
}

export function isDecorativeImage(image: Image): boolean {
  return image.decorative === true;
}

export function getWrapDistancesPx(image: Image): {
  top: number;
  bottom: number;
  left: number;
  right: number;
} {
  return {
    top: emuToPixels(image.wrap.distT),
    bottom: emuToPixels(image.wrap.distB),
    left: emuToPixels(image.wrap.distL),
    right: emuToPixels(image.wrap.distR),
  };
}

export function needsTextWrapping(image: Image): boolean {
  return ['square', 'tight', 'through', 'topAndBottom'].includes(image.wrap.type);
}
