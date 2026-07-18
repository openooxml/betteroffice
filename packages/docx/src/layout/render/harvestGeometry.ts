// normalized geometry snapshot over the mirror DOM contract. one schema,
// two producers: harvestPaintedPage walks a rendered mirror page, and
// snapshotFromDisplayPage normalizes a DisplayPage directly. diffGeometry
// compares any two snapshots with sub-pixel tolerance. INTERNAL test
// infrastructure (the mirror-vs-display-list unit gate) — not exported
// from the public render barrel.

import type { DisplayPage, DisplayPrimitive } from './displayList';
import { glyphRunRect, lineRect, textRunRect, type GeoRect } from './displayListGeometry';

// Compatibility re-exports for internal test consumers; surviving
// renderer/query consumers import from displayListGeometry directly.
export {
  fontSizePxFromShorthand,
  glyphRunRect,
  lineRect,
  textRunRect,
  type GeoRect,
} from './displayListGeometry';

export type GeometryNodeKind = 'text' | 'rect' | 'image' | 'shape' | 'decoration' | 'edge';

/** which page region a node belongs to; absent = body content */
export type GeometryRegion = 'header' | 'footer';

/** one dataset-contract node, normalized: rect relative to page origin (px). */
export interface GeometryNode {
  kind: GeometryNodeKind;
  rect: GeoRect;
  /**
   * region scope — header/footer nodes live in a different PM doc than body
   * nodes, so the diff must never match them across regions (absent = body)
   */
  region?: GeometryRegion;
  docStart?: number;
  docEnd?: number;
  /** display-list block ids are numbers; painter fragments may carry string ids */
  blockId?: number | string;
  commentIds?: string[];
  revision?: { author?: string; date?: string; revisionId?: string; kind?: 'ins' | 'del' };
  /** text content — text nodes only */
  text?: string;
  /** decoration variety — decoration nodes only */
  deco?: string;
  /** edge role (border | table-border | table-cut | separator) — edge nodes only */
  role?: string;
  /** image relationship id — image nodes only */
  relId?: string;
}

export interface PageGeometrySnapshot {
  pageIndex?: number;
  width: number;
  height: number;
  nodes: GeometryNode[];
}

/**
 * normalize a DisplayPage into the snapshot schema (the oracle side of a
 * diff). header/footer region primitives normalize like body primitives —
 * they are already in page coordinates — but carry the region tag so the
 * diff stays region-scoped.
 */
export function snapshotFromDisplayPage(page: DisplayPage): PageGeometrySnapshot {
  const nodes: GeometryNode[] = [];
  const push = (p: DisplayPrimitive, region?: GeometryRegion): void => {
    const node = nodeOfPrimitive(p);
    if (region) node.region = region;
    nodes.push(node);
  };
  for (const p of page.primitives) push(p);
  for (const region of [page.header, page.footer]) {
    if (!region) continue;
    for (const p of region.primitives) push(p, region.kind);
  }
  return { pageIndex: page.pageIndex, width: page.width, height: page.height, nodes };
}

function nodeOfPrimitive(p: DisplayPrimitive): GeometryNode {
  switch (p.kind) {
    case 'text':
      return { kind: 'text', rect: textRunRect(p), text: p.text, ...docAttrsOf(p) };
    // a glyph run normalizes to the same 'text' node the mirror emits for it,
    // so painter/mirror geometry stays comparable through the diff harness
    case 'glyphRun':
      return { kind: 'text', rect: glyphRunRect(p), text: p.text, ...docAttrsOf(p) };
    case 'rect':
      return { kind: 'rect', rect: { x: p.x, y: p.y, w: p.w, h: p.h }, ...docAttrsOf(p) };
    case 'line':
      return { kind: 'edge', rect: lineRect(p), role: p.role ?? 'border' };
    case 'image':
      return {
        kind: 'image',
        rect: { x: p.x, y: p.y, w: p.w, h: p.h },
        relId: p.relId,
        ...docAttrsOf(p),
      };
    case 'shape':
      return { kind: 'shape', rect: { x: p.x, y: p.y, w: p.w, h: p.h }, ...docAttrsOf(p) };
    case 'decoration':
      return {
        kind: 'decoration',
        rect: { x: p.x, y: p.y, w: p.w, h: p.h },
        deco: p.deco,
        ...docAttrsOf(p),
      };
  }
}

function docAttrsOf(p: {
  docStart?: number;
  docEnd?: number;
  blockId?: number;
  commentIds?: string[];
  revision?: { author: string; date: string; revisionId: string; kind: 'ins' | 'del' };
}): Partial<GeometryNode> {
  const out: Partial<GeometryNode> = {};
  if (p.docStart !== undefined) out.docStart = p.docStart;
  if (p.docEnd !== undefined) out.docEnd = p.docEnd;
  if (p.blockId !== undefined) out.blockId = p.blockId;
  if (p.commentIds && p.commentIds.length > 0) out.commentIds = [...p.commentIds];
  if (p.revision) out.revision = { ...p.revision };
  return out;
}

export interface HarvestOptions {
  /**
   * how element rects are resolved:
   * - 'layout' — getBoundingClientRect relative to the page element (real browser)
   * - 'styles' — accumulate inline left/top of absolutely positioned ancestors
   *   (mirror output / layout-less DOM like happy-dom)
   * - 'auto' (default) — 'layout' when the page element reports a nonzero
   *   rect, else 'styles'
   */
  rectSource?: 'auto' | 'layout' | 'styles';
}

// class → node kind. painter and mirror both emit these; painter has no
// element-level decoration/rect nodes yet (underline/shading are CSS there),
// so those kinds only appear when harvesting mirror output.
const KIND_SELECTORS: ReadonlyArray<{ selector: string; kind: GeometryNodeKind }> = [
  { selector: '.layout-run-text', kind: 'text' },
  { selector: '.layout-run-image', kind: 'image' },
  { selector: '.layout-run-shape', kind: 'shape' },
  { selector: '.layout-mirror-rect', kind: 'rect' },
  { selector: '.layout-decoration', kind: 'decoration' },
  { selector: '.layout-table-cut-border, .layout-mirror-line', kind: 'edge' },
];

/**
 * walk a rendered page's DOM (painter output or mirror output) and produce a
 * normalized geometry snapshot: one node per dataset-contract element, rects
 * relative to the page origin. the body walk scopes to `.layout-page-content`;
 * header/footer nodes come from a second walk scoped to
 * `.layout-page-header` / `.layout-page-footer` and carry the region tag
 * (their doc ranges refer to the region's HF PM doc, not the body doc).
 */
export function harvestPaintedPage(
  pageEl: Element,
  options: HarvestOptions = {}
): PageGeometrySnapshot {
  const useLayout = resolveRectSource(pageEl, options.rectSource ?? 'auto');
  const pageOrigin = useLayout ? pageEl.getBoundingClientRect() : undefined;
  const bodyScope = pageEl.querySelector('.layout-page-content') ?? pageEl;

  const nodes: GeometryNode[] = [];

  // when the body scope had to fall back to the page element, HF subtrees are
  // inside it — exclude them so their nodes only appear region-tagged below
  const inHfRegion = (el: Element): boolean =>
    bodyScope === pageEl && el.closest('.layout-page-header, .layout-page-footer') !== null;

  nodes.push(...harvestScope(bodyScope, pageEl, pageOrigin, undefined, inHfRegion));
  for (const { selector, region } of [
    { selector: '.layout-page-header', region: 'header' as const },
    { selector: '.layout-page-footer', region: 'footer' as const },
  ]) {
    for (const regionEl of Array.from(pageEl.querySelectorAll<HTMLElement>(selector))) {
      nodes.push(...harvestScope(regionEl, pageEl, pageOrigin, region, () => false));
    }
  }

  const pageRect = useLayout ? pageEl.getBoundingClientRect() : undefined;
  const width = pageRect?.width || styleLengthOf(pageEl, 'width') || 0;
  const height = pageRect?.height || styleLengthOf(pageEl, 'height') || 0;
  const pageIndexAttr = (pageEl as HTMLElement).dataset?.pageIndex;

  return {
    ...(pageIndexAttr !== undefined ? { pageIndex: Number(pageIndexAttr) } : {}),
    width,
    height,
    nodes,
  };
}

// one contract walk over a scope element, nodes sorted in DOM order; the
// exclude predicate lets the body walk skip HF subtrees on fallback scoping
function harvestScope(
  scope: Element,
  pageEl: Element,
  pageOrigin: DOMRect | undefined,
  region: GeometryRegion | undefined,
  exclude: (el: Element) => boolean
): GeometryNode[] {
  const collected: Array<{ node: GeometryNode; order: number }> = [];
  const seen = new Set<Element>();
  const all = Array.from(scope.querySelectorAll<HTMLElement>('*'));

  for (const { selector, kind } of KIND_SELECTORS) {
    for (const el of Array.from(scope.querySelectorAll<HTMLElement>(selector))) {
      if (seen.has(el) || exclude(el)) continue;
      seen.add(el);
      const node = harvestNode(el, kind, pageEl, pageOrigin);
      if (region) node.region = region;
      collected.push({ node, order: all.indexOf(el) });
    }
  }

  collected.sort((a, b) => a.order - b.order);
  return collected.map((n) => n.node);
}

function resolveRectSource(pageEl: Element, source: 'auto' | 'layout' | 'styles'): boolean {
  if (source === 'layout') return true;
  if (source === 'styles') return false;
  return pageEl.getBoundingClientRect().width > 0;
}

function harvestNode(
  el: HTMLElement,
  kind: GeometryNodeKind,
  pageEl: Element,
  pageOrigin: DOMRect | undefined
): GeometryNode {
  const node: GeometryNode = { kind, rect: rectOf(el, pageEl, pageOrigin) };

  const start = ownOrClosest(el, 'data-doc-start');
  const end = ownOrClosest(el, 'data-doc-end');
  if (start !== null) node.docStart = Number(start);
  if (end !== null) node.docEnd = Number(end);

  const blockHost = el.closest('[data-block-id]');
  const blockId = blockHost?.getAttribute('data-block-id');
  if (blockId !== null && blockId !== undefined && blockId !== '') {
    const parsed = Number(blockId);
    node.blockId = Number.isFinite(parsed) ? parsed : blockId;
  }

  const commentId = el.getAttribute('data-comment-id');
  if (commentId) node.commentIds = commentId.split(' ').filter(Boolean);

  const author = el.getAttribute('data-change-author');
  const date = el.getAttribute('data-change-date');
  const revisionId = el.getAttribute('data-revision-id');
  if (author !== null || date !== null || revisionId !== null) {
    node.revision = {
      ...(author !== null ? { author } : {}),
      ...(date !== null ? { date } : {}),
      ...(revisionId !== null ? { revisionId } : {}),
      ...(el.classList.contains('docx-insertion')
        ? { kind: 'ins' as const }
        : el.classList.contains('docx-deletion')
          ? { kind: 'del' as const }
          : {}),
    };
  }

  if (kind === 'text') node.text = el.textContent ?? '';
  if (kind === 'decoration') node.deco = el.getAttribute('data-deco') ?? undefined;
  if (kind === 'edge') {
    node.role = el.classList.contains('layout-table-cut-border')
      ? 'table-cut'
      : (el.getAttribute('data-role') ?? 'border');
  }
  if (kind === 'image') node.relId = el.getAttribute('data-rel-id') ?? undefined;

  return node;
}

// doc range lives on the node itself in the mirror; painter text spans also
// carry it directly (applyPmPositions), so own-attribute wins and closest()
// only backfills structural wrappers.
function ownOrClosest(el: HTMLElement, attr: string): string | null {
  const own = el.getAttribute(attr);
  if (own !== null) return own;
  return el.closest(`[${attr}]`)?.getAttribute(attr) ?? null;
}

function rectOf(el: HTMLElement, pageEl: Element, pageOrigin: DOMRect | undefined): GeoRect {
  if (pageOrigin) {
    const r = el.getBoundingClientRect();
    return { x: r.left - pageOrigin.left, y: r.top - pageOrigin.top, w: r.width, h: r.height };
  }
  // style accumulation: sum inline left/top of every absolutely positioned
  // element from the node up to (not including) the page element. only the
  // mirror's shape is supported here — absolutely positioned, inline px styles.
  let x = 0;
  let y = 0;
  let cur: HTMLElement | null = el;
  while (cur && cur !== pageEl) {
    if (cur.style.position === 'absolute') {
      x += parseFloat(cur.style.left || '0') || 0;
      y += parseFloat(cur.style.top || '0') || 0;
    }
    cur = cur.parentElement;
  }
  return {
    x,
    y,
    w: styleLengthOf(el, 'width') ?? 0,
    h: styleLengthOf(el, 'height') ?? 0,
  };
}

function styleLengthOf(el: Element, prop: 'width' | 'height'): number | undefined {
  const value = (el as HTMLElement).style?.[prop];
  if (!value) return undefined;
  const parsed = parseFloat(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}
