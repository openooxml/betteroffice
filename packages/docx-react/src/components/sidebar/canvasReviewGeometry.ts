import { useLayoutEffect, useMemo, useState, type RefObject } from 'react';
import { canvasPageTops, displayPageCanvas, type DisplayListQueries, type DisplayListRect } from '@betteroffice/docx/layout/render';
import type { ReviewPage } from './TrackedChangesReview';
import type { ReviewSidebarItem } from './trackedChanges';
import type { TrackedChangeEntry } from './cardUtils';

type Geometry = { pages: ReviewPage[]; rects: [string, DisplayListRect[]][]; zoom: number };

/** Page bounds and change rects for the review layer on the in-thread canvas route, relative to `originRef`. */
export function useCanvasReviewGeometry(queries: DisplayListQueries | null, canvasHostRef: RefObject<HTMLElement | null>,
  originRef: RefObject<HTMLElement | null>, items: readonly ReviewSidebarItem[], zoom: number, epoch: unknown) {
  const [geometry, setGeometry] = useState<Geometry>({ pages: [], rects: [], zoom });
  const ranges = JSON.stringify(items.flatMap(item => {
    const change = item.review?.change as (TrackedChangeEntry & { hfRid?: string }) | undefined;
    return change && !change.hfRid ? [[item.id, change.from, change.to]] : [];
  }));
  useLayoutEffect(() => {
    const host = canvasHostRef.current, origin = originRef.current;
    if (!queries || !host || !origin) return;
    let live = true;
    const measure = () => {
      if (!live) return;
      const base = origin.getBoundingClientRect();
      const tops = canvasPageTops(queries);
      const sizes = Array.from({ length: queries.pageCount() }, (_, index) => queries.pageSize(index));
      const measured = sizes.flatMap((size, index) => {
        const canvas = size && displayPageCanvas(host, index)?.getBoundingClientRect();
        return canvas && size.width > 0 ? [{ index, canvas, scale: canvas.width / size.width }] : [];
      });
      const reference = measured[0];
      const scale = reference?.scale || zoom || 1;
      const pages = sizes.flatMap((size, index) => {
        if (!size) return [];
        const canvas = measured.find(entry => entry.index === index)?.canvas;
        const x = canvas ? (canvas.left - base.left) / scale : reference ? (reference.canvas.left - base.left) / scale : 0;
        const y = canvas ? (canvas.top - base.top) / scale
          : reference ? (reference.canvas.top - base.top) / scale + tops[index]! - tops[reference.index]! : tops[index]!;
        return [{ page: { pageIndex: index }, bounds: { x, y, width: size.width, height: size.height } }];
      });
      const rects = (JSON.parse(ranges) as [string, number, number][]).map(([id, from, to]) => [id, queries.rangeRects(from, to)] as [string, DisplayListRect[]]);
      const next = { pages, rects, zoom: scale };
      setGeometry(previous => JSON.stringify(previous) === JSON.stringify(next) ? previous : next);
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(host);
    observer.observe(origin);
    const view = host.ownerDocument.defaultView;
    view?.addEventListener('resize', measure);
    host.addEventListener('transitionend', measure);
    return () => { live = false; observer.disconnect(); view?.removeEventListener('resize', measure); host.removeEventListener('transitionend', measure); };
  }, [queries, canvasHostRef, originRef, ranges, zoom, epoch]);
  const rects = useMemo(() => new Map(geometry.rects), [geometry.rects]);
  return { pages: geometry.pages, rects, zoom: geometry.zoom };
}
