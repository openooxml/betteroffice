/**
 * Header/footer body-margin extension.
 *
 * Word grows the header (or footer) band when its in-flow content is taller
 * than the authored top (or bottom) margin minus the header/footer distance,
 * pushing the body text down (or up). This module owns that computation so the
 * React and Vue adapters share one implementation instead of byte-identical
 * inline copies (the layout pipelines were drifting candidates — see
 * `docx-editor` engine-unification work, issue #696).
 *
 * Two correctness rules live here:
 *
 *  1. The band height is driven by `HeaderFooterContent.flowHeight` (in-flow
 *     content only), NOT `height` / `visualBottom`. A page/margin-anchored
 *     float — e.g. a full-page letterhead anchored in a header — is positioned
 *     on the page and does not push the body in Word. Counting it inflated the
 *     effective top margin past the page on real-world templates, so the
 *     paginator hard-threw "page size and margins yield no content area" and
 *     the document rendered blank (issue #705).
 *
 *  2. A clamp guarantees `top + bottom` never consumes the whole page, so a
 *     pathological in-flow header degrades to a thin content band with a
 *     warning instead of aborting pagination.
 */

import type { LayoutBlock, PageMargins, SectionBreakBlock } from '../pagination/types';
import type { HeaderFooterContent } from '../pagination/types';

/** Word's default `w:header` / `w:footer` distance (0.5in = 48px). */
const DEFAULT_HF_DISTANCE_PX = 48;

/**
 * Floor on the body content area. Even when header/footer content is absurdly
 * tall, leave at least this much height so pagination produces a page instead
 * of throwing. ~one line at the default body font.
 */
const MIN_CONTENT_HEIGHT_PX = 24;

/** In-flow band height for one HF variant (falls back to total height). */
function bandHeight(hf: HeaderFooterContent | undefined): number {
  if (!hf) return 0;
  return hf.flowHeight ?? hf.height;
}

/** @public */
export interface ExtendMarginsForHeaderFooterInput {
  pageSize: { w: number; h: number };
  /** Body fallback margins. */
  margins: PageMargins;
  /** Final-section margins (last `sectPr`). */
  finalMargins: PageMargins;
  /**
   * Body flow blocks. Each `sectionBreak` block's `margins` is extended IN
   * PLACE so multi-section documents paginate with the same band growth (the
   * layout engine prefers `sectionBreak.margins` over the body fallback).
   */
  bodyBlocks?: LayoutBlock[];
  /** Header variants in play this layout (e.g. default + first-page). */
  headers?: Array<HeaderFooterContent | undefined>;
  /** Footer variants in play this layout. */
  footers?: Array<HeaderFooterContent | undefined>;
  /** Optional diagnostic sink for the clamp (adapters pass `console.warn`). */
  warn?: (message: string) => void;
}

/** @public */
export interface ExtendMarginsForHeaderFooterResult {
  margins: PageMargins;
  finalMargins: PageMargins;
}

/**
 * Extend body margins so the body clears the header/footer bands, mirroring
 * Word. Returns new `margins` / `finalMargins`; mutates `sectionBreak.margins`
 * in place. When no extension is needed the original objects are returned
 * unchanged.
 *
 * @public
 */
export function extendMarginsForHeaderFooter(
  input: ExtendMarginsForHeaderFooterInput
): ExtendMarginsForHeaderFooterResult {
  const { pageSize, margins, finalMargins, bodyBlocks, headers, footers, warn } = input;

  const headerDistance = margins.header ?? DEFAULT_HF_DISTANCE_PX;
  const footerDistance = margins.footer ?? DEFAULT_HF_DISTANCE_PX;
  const availableHeaderSpace = margins.top - headerDistance;
  const availableFooterSpace = margins.bottom - footerDistance;

  const headerContentHeight = Math.max(0, ...(headers ?? []).map(bandHeight));
  const footerContentHeight = Math.max(0, ...(footers ?? []).map(bandHeight));

  const extendHeader = headerContentHeight > availableHeaderSpace;
  const extendFooter = footerContentHeight > availableFooterSpace;
  if (!extendHeader && !extendFooter) {
    return { margins, finalMargins };
  }

  const maxMargins = Math.max(0, pageSize.h - MIN_CONTENT_HEIGHT_PX);
  let clamped = false;

  const extend = (m: PageMargins): PageMargins => {
    const out = { ...m };
    if (extendHeader) out.top = Math.max(m.top, headerDistance + headerContentHeight);
    if (extendFooter) out.bottom = Math.max(m.bottom, footerDistance + footerContentHeight);
    // Safety net: never let header + footer consume the whole page. Clamp the
    // footer band first (it sits at the page bottom), then the header band if
    // it alone still overflows, so the body keeps a positive content area.
    if (out.top + out.bottom > maxMargins) {
      clamped = true;
      out.bottom = Math.max(0, Math.min(out.bottom, maxMargins - out.top));
      if (out.top + out.bottom > maxMargins) {
        out.top = Math.max(0, maxMargins - out.bottom);
      }
    }
    return out;
  };

  const extendedMargins = extend(margins);
  const extendedFinal = extend(finalMargins);
  if (bodyBlocks) {
    for (const block of bodyBlocks) {
      if (block.kind !== 'sectionBreak') continue;
      const sb = block as SectionBreakBlock;
      if (sb.margins) sb.margins = extend(sb.margins);
    }
  }

  if (clamped && warn) {
    warn(
      '[layout] header/footer content exceeds page height; clamping margins to ' +
        `preserve a content area. pageHeight=${Math.round(pageSize.h)} ` +
        `headerBand=${Math.round(headerContentHeight)} footerBand=${Math.round(footerContentHeight)} ` +
        `top=${Math.round(extendedMargins.top)} bottom=${Math.round(extendedMargins.bottom)}`
    );
  }

  return { margins: extendedMargins, finalMargins: extendedFinal };
}
