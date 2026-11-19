import { findDisplayListHyperlinkAtPoint, type DisplayPage } from '@betteroffice/docx/layout/render';
import type { HyperlinkPopupData } from '../../ui/HyperlinkPopup';

/** Prefer the complete link read at the caret over the clicked paint fragment. */
export function popupWithLink(hit: HyperlinkPopupData,
  link: { url?: string; displayText?: string; tooltip?: string } | null | undefined): HyperlinkPopupData {
  if (!link?.url) return hit;
  const tooltip = link.tooltip ?? hit.tooltip;
  return { ...hit, displayText: link.displayText || hit.displayText, ...(tooltip ? { tooltip } : {}) };
}

const REGIONS = [{ kind: 'body' }, { kind: 'header' }, { kind: 'footer' }] as const;

/** The mirror ignores pointer input, so the pages host exposes link ScreenTips. */
export function hyperlinkHoverTitle(page: DisplayPage | null | undefined, x: number, y: number): string | null {
  if (!page) return null;
  for (const region of REGIONS) {
    const hit = findDisplayListHyperlinkAtPoint({ pages: [page] }, 0, x, y, region);
    if (hit) return hit.tooltip?.trim() || null;
  }
  return null;
}
