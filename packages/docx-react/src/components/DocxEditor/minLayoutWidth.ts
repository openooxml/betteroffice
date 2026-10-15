import { SIDEBAR_DOCUMENT_SHIFT } from '../sidebar/constants';

export interface MinLayoutWidthInput {
  outlineLeftAllowance: number;
  /** Widest section page at 100% zoom. */
  maxPageWidthPx: number;
  zoom: number;
  sidebarOpen: boolean;
}

/** Content width below which the scroll container scrolls horizontally. Pages scale with zoom; the outline and sidebar chrome does not. */
export function minLayoutWidth({ outlineLeftAllowance, maxPageWidthPx, zoom, sidebarOpen }: MinLayoutWidthInput): number {
  return (
    2 * outlineLeftAllowance +
    Math.ceil(maxPageWidthPx * zoom) +
    (sidebarOpen ? SIDEBAR_DOCUMENT_SHIFT * 2 : 0)
  );
}
