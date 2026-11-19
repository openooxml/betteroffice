import { findDisplayListHyperlinkAtPoint, type DisplayPage } from '@betteroffice/docx/layout/render';
import type { HyperlinkPopupData } from '../../ui/HyperlinkPopup';

/**
 * The popup for a clicked link, with the text and ScreenTip of the whole link the worker reads in
 * place of the hit's. A click hits one painted run, so the hit's text is only part of a link that
 * spans formats or lines. Without a read, or a read that found no link, the hit stands.
 */
export function popupWithLink(hit: HyperlinkPopupData,
  link: { url?: string; displayText?: string; tooltip?: string } | null | undefined): HyperlinkPopupData {
  if (!link?.url) return hit;
  const tooltip = link.tooltip ?? hit.tooltip;
  return { ...hit, displayText: link.displayText || hit.displayText, ...(tooltip ? { tooltip } : {}) };
}

const REGIONS = [{ kind: 'body' }, { kind: 'header' }, { kind: 'footer' }] as const;

/**
 * The ScreenTip of the hyperlink under page point (`x`, `y`), for hover. `null` off a link or on a
 * link without one.
 *
 * The page mirror already carries the ScreenTip as a DOM `title`, but the mirror ignores the
 * pointer, so the browser never shows it; the pages host shows this instead.
 */
export function hyperlinkHoverTitle(page: DisplayPage | null | undefined, x: number, y: number): string | null {
  if (!page) return null;
  for (const region of REGIONS) {
    const hit = findDisplayListHyperlinkAtPoint({ pages: [page] }, 0, x, y, region);
    if (hit) return hit.tooltip?.trim() || null;
  }
  return null;
}
