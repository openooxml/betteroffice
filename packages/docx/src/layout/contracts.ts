import type { Document, HeaderFooter, SectionProperties } from '../types/document';
import type { DisplayListQueries, DisplayListRect } from './render';

export const DEFAULT_PAGE_HEIGHT_PX = 1056;

export function resolveHeaderFooter(
  document: Document | null,
  section: SectionProperties | null | undefined
): {
  header: HeaderFooter | null;
  footer: HeaderFooter | null;
  firstHeader: HeaderFooter | null;
  firstFooter: HeaderFooter | null;
} {
  const headers = document?.package.headers;
  const footers = document?.package.footers;
  const headerRef = section?.headerReferences?.find((ref) => ref.type === 'default');
  const firstHeaderRef = section?.headerReferences?.find((ref) => ref.type === 'first');
  const footerRef = section?.footerReferences?.find((ref) => ref.type === 'default');
  const firstFooterRef = section?.footerReferences?.find((ref) => ref.type === 'first');
  let header = headerRef ? (headers?.get(headerRef.rId) ?? null) : null;
  let footer = footerRef ? (footers?.get(footerRef.rId) ?? null) : null;
  const firstHeader = firstHeaderRef ? (headers?.get(firstHeaderRef.rId) ?? null) : null;
  const firstFooter = firstFooterRef ? (footers?.get(firstFooterRef.rId) ?? null) : null;
  if (!section?.titlePg) {
    header ??= firstHeader;
    footer ??= firstFooter;
  }
  return { header, footer, firstHeader, firstFooter };
}

export function computeHfCaretRectsFromDisplayList(
  queries: DisplayListQueries,
  section: 'header' | 'footer',
  rId: string,
  position: number,
  pageIndex?: number
): DisplayListRect[] {
  const rects = queries.hfCaretRects(section, rId, position);
  return pageIndex === undefined ? rects : rects.filter((rect) => rect.pageIndex === pageIndex);
}

export function computeHfSelectionRectsFromDisplayList(
  queries: DisplayListQueries,
  section: 'header' | 'footer',
  rId: string,
  from: number,
  to: number,
  pageIndex?: number
): DisplayListRect[] {
  const rects = queries.hfRangeRects(section, rId, Math.min(from, to), Math.max(from, to));
  return pageIndex === undefined ? rects : rects.filter((rect) => rect.pageIndex === pageIndex);
}
