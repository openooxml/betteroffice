import type { PeerlessFrameGeometry } from './viewStore';

type PageGeometry = PeerlessFrameGeometry['pages'][number];
type PageIndex = { bounds: PageGeometry['bounds']; page?: PageGeometry; children?: readonly [PageIndex, PageIndex] };
const indexes = new WeakMap<readonly PageGeometry[], PageIndex | null>();

function indexPages(pages: readonly PageGeometry[]): PageIndex | null {
  if (!pages.length) return null;
  if (pages.length === 1) return { bounds: pages[0]!.bounds, page: pages[0]! };
  const sorted = [...pages].sort((a, b) => a.bounds.y - b.bounds.y);
  const middle = Math.floor(sorted.length / 2);
  const left = indexPages(sorted.slice(0, middle))!;
  const right = indexPages(sorted.slice(middle))!;
  const x = Math.min(left.bounds.x, right.bounds.x);
  const y = Math.min(left.bounds.y, right.bounds.y);
  return { bounds: {
    x, y,
    width: Math.max(left.bounds.x + left.bounds.width, right.bounds.x + right.bounds.width) - x,
    height: Math.max(left.bounds.y + left.bounds.height, right.bounds.y + right.bounds.height) - y,
  }, children: [left, right] };
}

function nearestPage(index: PageIndex, x: number, y: number): PageGeometry | undefined {
  const distance = ({ bounds }: PageIndex) => Math.hypot(
    Math.max(bounds.x - x, 0, x - bounds.x - bounds.width),
    Math.max(bounds.y - y, 0, y - bounds.y - bounds.height),
  );
  let nearest: PageGeometry | undefined;
  let best = Infinity;
  const visit = (node: PageIndex) => {
    const gap = distance(node);
    if (gap >= best) return;
    if (node.page) {
      if (node.bounds.width > 0 && node.bounds.height > 0) {
        best = gap;
        nearest = node.page;
      }
      return;
    }
    const [left, right] = node.children!;
    if (distance(left) <= distance(right)) { visit(left); visit(right); }
    else { visit(right); visit(left); }
  };
  visit(index);
  return nearest;
}

export function clientPointToPage(clientX: number, clientY: number, host: Pick<DOMRectReadOnly, 'left' | 'top'>,
  zoom: number, pages: readonly PageGeometry[]) {
  if (!Number.isFinite(zoom) || zoom <= 0) return null;
  const x = (clientX - host.left) / zoom;
  const y = (clientY - host.top) / zoom;
  if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
  let index = indexes.get(pages);
  if (index === undefined) { index = indexPages(pages); indexes.set(pages, index); }
  const page = index && nearestPage(index, x, y);
  return page ? { page, x: x - page.bounds.x, y: y - page.bounds.y } : null;
}
