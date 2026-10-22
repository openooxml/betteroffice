// Re-export from core where the canonical constants live.
export {
  SIDEBAR_WIDTH,
  SIDEBAR_PAGE_GAP,
  SIDEBAR_DOCUMENT_SHIFT,
  MIN_CARD_GAP,
} from '@betteroffice/docx/utils/sidebarConstants';

/** Tracked-change card width; the column narrows to it when it holds only review cards. Matches `--doc-review-card-width`. */
export const REVIEW_CARD_WIDTH = 240;
