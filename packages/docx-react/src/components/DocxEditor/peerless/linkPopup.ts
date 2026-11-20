import type { HyperlinkPopupData } from '../../ui/HyperlinkPopup';

/** Prefer the complete link read at the caret over the clicked paint fragment. */
export function popupWithLink(hit: HyperlinkPopupData,
  link: { url?: string; displayText?: string } | null | undefined): HyperlinkPopupData {
  if (!link?.url) return hit;
  return { ...hit, displayText: link.displayText || hit.displayText };
}
