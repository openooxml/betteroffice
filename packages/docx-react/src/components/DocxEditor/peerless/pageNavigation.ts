import type { PeerlessFrameGeometry } from './viewStore';

export function workerPageAtViewport(pages: PeerlessFrameGeometry['pages'], scrollTop: number, height: number, zoom: number) {
  if (!pages.length) return 1;
  const middle = (scrollTop + height / 2) / zoom;
  let low = 0;
  let high = pages.length - 1;
  while (low < high) {
    const index = (low + high) >> 1;
    const bounds = pages[index]!.bounds;
    if (bounds.y + bounds.height > middle) high = index; else low = index + 1;
  }
  return pages[low]!.page.pageIndex + 1;
}
