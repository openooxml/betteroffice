import type { DisplayList, DisplayPage, DisplayPrimitive } from './displayList';
import type { DisplayListTableRegion } from './displayListTables';
import { glyphRunRect, textRunRect, type GeoRect } from './displayListGeometry';

export interface DisplayListHyperlinkHit {
  pageIndex: number;
  region: 'body' | 'header' | 'footer';
  rId?: string;
  href: string;
  tooltip?: string;
  displayText: string;
  rect: GeoRect;
}

interface RegionPrimitives {
  kind: 'body' | 'header' | 'footer';
  rId?: string;
  primitives: DisplayPrimitive[];
}

function primitivesForRegion(
  page: DisplayPage,
  region: DisplayListTableRegion
): RegionPrimitives | null {
  const kind = region.kind ?? 'body';
  if (kind === 'body') return { kind, primitives: page.primitives };
  const hf = kind === 'header' ? page.header : page.footer;
  if (!hf || hf.kind !== kind) return null;
  if (region.rId && hf.rId !== region.rId) return null;
  return { kind, rId: hf.rId, primitives: hf.primitives };
}

function primitiveRect(p: DisplayPrimitive): GeoRect {
  switch (p.kind) {
    case 'text':
      return textRunRect(p);
    case 'glyphRun':
      return glyphRunRect(p);
    case 'rect':
    case 'image':
    case 'shape':
    case 'decoration':
      return { x: p.x, y: p.y, w: p.w, h: p.h };
    case 'line':
      return {
        x: Math.min(p.x1, p.x2),
        y: Math.min(p.y1, p.y2),
        w: Math.abs(p.x2 - p.x1),
        h: Math.abs(p.y2 - p.y1),
      };
  }
}

function clippedPrimitiveRect(p: DisplayPrimitive): GeoRect | null {
  const rect = primitiveRect(p);
  const clip = p.clipGroup?.clip;
  if (!clip) return rect;
  const clipX = clip.x ?? 0;
  const clipY = clip.y ?? 0;
  const clipRight = clipX + Math.max(0, clip.w ?? 0);
  const clipBottom = clipY + Math.max(0, clip.h ?? 0);
  const x = Math.max(rect.x, clipX);
  const y = Math.max(rect.y, clipY);
  const right = Math.min(rect.x + rect.w, clipRight);
  const bottom = Math.min(rect.y + rect.h, clipBottom);
  if (right < x || bottom < y) return null;
  return { x, y, w: right - x, h: bottom - y };
}

function primitiveText(p: DisplayPrimitive): string {
  if (p.kind === 'text' || p.kind === 'glyphRun') return p.text;
  if (p.kind === 'image') return p.altText ?? '';
  return '';
}

function containsPoint(rect: GeoRect, x: number, y: number): boolean {
  return x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h;
}

export function findDisplayListHyperlinkAtPoint(
  list: DisplayList,
  pageIndex: number,
  x: number,
  y: number,
  region: DisplayListTableRegion = { kind: 'body' }
): DisplayListHyperlinkHit | null {
  const page = list.pages[pageIndex];
  if (!page) return null;
  const regionPrims = primitivesForRegion(page, region);
  if (!regionPrims) return null;

  for (let i = regionPrims.primitives.length - 1; i >= 0; i--) {
    const p = regionPrims.primitives[i];
    if (p.kind === 'line' || !p.href || p.hiddenObject) continue;
    const rect = clippedPrimitiveRect(p);
    if (!rect) continue;
    if (!containsPoint(rect, x, y)) continue;
    return {
      pageIndex,
      region: regionPrims.kind,
      rId: regionPrims.rId,
      href: p.href,
      tooltip: p.linkTitle ?? p.tooltip,
      displayText: primitiveText(p) || p.href,
      rect,
    };
  }

  return null;
}
