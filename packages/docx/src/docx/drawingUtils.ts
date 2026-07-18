/**
 * Shared Drawing Parsing Utilities
 *
 * The pure DrawingML (a: namespace) half — color/fill/outline parsing and
 * the theme-less color fallback — lives in @betteroffice/drawingml and is
 * re-exported here. This module keeps the WordprocessingML drawing (wp:)
 * half: positions and text wrapping for anchored drawings.
 */

import type { ImagePosition, ImageWrap } from '../types/document';
import {
  getChildElements,
  getAttribute,
  getTextContent,
  parseNumericAttribute,
  findByFullName,
  findChildrenByLocalName,
  type XmlElement,
} from './xmlParser';

export {
  parseColorElement,
  parseFill,
  parseOutline,
  resolveColorValueToHex,
} from '@betteroffice/drawingml';

// ============================================================================
// POSITION PARSING
// ============================================================================

/**
 * Parse horizontal position from wp:positionH element.
 */
export function parsePositionH(posH: XmlElement | null): ImagePosition['horizontal'] | undefined {
  if (!posH) return undefined;

  const relativeTo = getAttribute(posH, null, 'relativeFrom') ?? 'column';

  const alignEl = findByFullName(posH, 'wp:align');
  if (alignEl) {
    const text = getTextContent(alignEl);
    return {
      relativeTo: relativeTo as ImagePosition['horizontal']['relativeTo'],
      alignment: text as ImagePosition['horizontal']['alignment'],
    };
  }

  const posOffsetEl = findByFullName(posH, 'wp:posOffset');
  if (posOffsetEl) {
    const text = getTextContent(posOffsetEl);
    const posOffset = parseInt(text, 10);
    return {
      relativeTo: relativeTo as ImagePosition['horizontal']['relativeTo'],
      posOffset: isNaN(posOffset) ? 0 : posOffset,
    };
  }

  return {
    relativeTo: relativeTo as ImagePosition['horizontal']['relativeTo'],
  };
}

/**
 * Parse vertical position from wp:positionV element.
 */
export function parsePositionV(posV: XmlElement | null): ImagePosition['vertical'] | undefined {
  if (!posV) return undefined;

  const relativeTo = getAttribute(posV, null, 'relativeFrom') ?? 'paragraph';

  const alignEl = findByFullName(posV, 'wp:align');
  if (alignEl) {
    const text = getTextContent(alignEl);
    return {
      relativeTo: relativeTo as ImagePosition['vertical']['relativeTo'],
      alignment: text as ImagePosition['vertical']['alignment'],
    };
  }

  const posOffsetEl = findByFullName(posV, 'wp:posOffset');
  if (posOffsetEl) {
    const text = getTextContent(posOffsetEl);
    const posOffset = parseInt(text, 10);
    return {
      relativeTo: relativeTo as ImagePosition['vertical']['relativeTo'],
      posOffset: isNaN(posOffset) ? 0 : posOffset,
    };
  }

  return {
    relativeTo: relativeTo as ImagePosition['vertical']['relativeTo'],
  };
}

/**
 * Parse position for anchored drawings (combines positionH + positionV).
 */
export function parseAnchorPosition(anchor: XmlElement): ImagePosition | undefined {
  const positionH = findByFullName(anchor, 'wp:positionH');
  const positionV = findByFullName(anchor, 'wp:positionV');
  const simplePos = findByFullName(anchor, 'wp:simplePos');
  const useSimplePos = getAttribute(anchor, null, 'simplePos') === '1';
  const relativeHeight = parseNumericAttribute(anchor, null, 'relativeHeight');
  const behindDoc = getAttribute(anchor, null, 'behindDoc') === '1';
  const hiddenRaw = getAttribute(anchor, null, 'hidden');
  const lockedRaw = getAttribute(anchor, null, 'locked');

  if (
    !positionH &&
    !positionV &&
    !simplePos &&
    !useSimplePos &&
    relativeHeight === undefined &&
    !behindDoc &&
    hiddenRaw === null &&
    lockedRaw === null
  ) {
    return undefined;
  }

  return {
    useSimplePos: useSimplePos || undefined,
    simplePos:
      simplePos && useSimplePos
        ? {
            x: parseNumericAttribute(simplePos, null, 'x') ?? undefined,
            y: parseNumericAttribute(simplePos, null, 'y') ?? undefined,
          }
        : undefined,
    relativeHeight:
      relativeHeight !== undefined && Number.isFinite(relativeHeight) ? relativeHeight : undefined,
    behindDoc: behindDoc || undefined,
    hidden: hiddenRaw === '1' || hiddenRaw === 'true' || undefined,
    locked: lockedRaw === '1' || lockedRaw === 'true' || undefined,
    horizontal: parsePositionH(positionH) ?? { relativeTo: 'column' },
    vertical: parsePositionV(positionV) ?? { relativeTo: 'paragraph' },
  };
}

// ============================================================================
// WRAP PARSING
// ============================================================================

/** Known wrap element names */
export const WRAP_ELEMENT_NAMES = [
  'wp:wrapNone',
  'wp:wrapSquare',
  'wp:wrapTight',
  'wp:wrapThrough',
  'wp:wrapTopAndBottom',
];

/**
 * Parse wrap settings from a wrap element.
 *
 * Distance attributes (distT/distB/distL/distR) can appear on both
 * the anchor element and the wrap child. Wrap child values take priority;
 * anchor-level values are used as fallbacks.
 */
export function parseWrapElement(
  wrapEl: XmlElement | null,
  behindDoc: boolean,
  anchorDistances?: { distT?: number; distB?: number; distL?: number; distR?: number }
): ImageWrap {
  if (!wrapEl) {
    const wrap: ImageWrap = { type: behindDoc ? 'behind' : 'inFront' };
    if (anchorDistances?.distT !== undefined) wrap.distT = anchorDistances.distT;
    if (anchorDistances?.distB !== undefined) wrap.distB = anchorDistances.distB;
    if (anchorDistances?.distL !== undefined) wrap.distL = anchorDistances.distL;
    if (anchorDistances?.distR !== undefined) wrap.distR = anchorDistances.distR;
    return wrap;
  }

  const wrapName = wrapEl.name || '';
  const wrapType = wrapName.replace('wp:', '');

  let type: ImageWrap['type'];
  switch (wrapType) {
    case 'wrapNone':
      type = behindDoc ? 'behind' : 'inFront';
      break;
    case 'wrapSquare':
      type = 'square';
      break;
    case 'wrapTight':
      type = 'tight';
      break;
    case 'wrapThrough':
      type = 'through';
      break;
    case 'wrapTopAndBottom':
      type = 'topAndBottom';
      break;
    default:
      type = 'square';
  }

  const wrap: ImageWrap = { type };

  const wrapText = getAttribute(wrapEl, null, 'wrapText');
  if (wrapText) wrap.wrapText = wrapText as ImageWrap['wrapText'];

  // Wrap child distances take priority, then anchor-level
  const distT = parseNumericAttribute(wrapEl, null, 'distT') ?? anchorDistances?.distT;
  const distB = parseNumericAttribute(wrapEl, null, 'distB') ?? anchorDistances?.distB;
  const distL = parseNumericAttribute(wrapEl, null, 'distL') ?? anchorDistances?.distL;
  const distR = parseNumericAttribute(wrapEl, null, 'distR') ?? anchorDistances?.distR;

  if (distT !== undefined) wrap.distT = distT;
  if (distB !== undefined) wrap.distB = distB;
  if (distL !== undefined) wrap.distL = distL;
  if (distR !== undefined) wrap.distR = distR;

  const polygon = findByFullName(wrapEl, 'wp:wrapPolygon');
  if (polygon) {
    const points = [
      ...findChildrenByLocalName(polygon, 'start'),
      ...findChildrenByLocalName(polygon, 'lineTo'),
    ]
      .slice(0, 2048)
      .map((point) => ({
        x: parseNumericAttribute(point, null, 'x') ?? undefined,
        y: parseNumericAttribute(point, null, 'y') ?? undefined,
      }))
      .filter(
        (point) =>
          point.x !== undefined &&
          point.y !== undefined &&
          Number.isFinite(point.x) &&
          Number.isFinite(point.y) &&
          Math.abs(point.x) <= 1_000_000_000 &&
          Math.abs(point.y) <= 1_000_000_000
      );
    if (points.length > 1) wrap.polygon = points;
  }

  return wrap;
}

/**
 * Parse wrap from an anchor element (finds wrap child internally).
 */
export function parseAnchorWrap(anchor: XmlElement): ImageWrap | undefined {
  const children = getChildElements(anchor);
  const behindDoc = getAttribute(anchor, null, 'behindDoc') === '1';

  const wrapEl = children.find((el) => WRAP_ELEMENT_NAMES.includes(el.name ?? ''));

  // Read anchor-level distance fallbacks
  const anchorDistances = {
    distT: parseNumericAttribute(anchor, null, 'distT') ?? undefined,
    distB: parseNumericAttribute(anchor, null, 'distB') ?? undefined,
    distL: parseNumericAttribute(anchor, null, 'distL') ?? undefined,
    distR: parseNumericAttribute(anchor, null, 'distR') ?? undefined,
  };

  return parseWrapElement(wrapEl ?? null, behindDoc, anchorDistances);
}
