import type { Document, HeaderFooter, SectionProperties } from '../types/document';

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
