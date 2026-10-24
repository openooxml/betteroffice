import { Fragment, useCallback, useEffect, useMemo, useRef, useState, type ReactNode, type RefObject } from 'react';
import type { DisplayListRect } from '@betteroffice/docx/layout/render';
import type { DocxEditorCollaborationOptions } from '../DocxEditor/types';
import { useTranslation } from '../../i18n';
import { SIDEBAR_DOCUMENT_SHIFT, SIDEBAR_PAGE_GAP } from './constants';
import { ambiguousReviewIds, resolveReviewAuthor, type ReviewAnchor, ReviewAvatar, reviewAuthorStyle, TrackedChangesContext, useReviewPeers,
  type ReviewSidebarItem, type TrackedChangeAuthor, type TrackedChangesOptions } from './trackedChanges';

export function ReviewSummary({ authors, count, orientation, acceptAll }: {
  authors: TrackedChangeAuthor[]; count: number; orientation: 'horizontal' | 'vertical'; acceptAll(): void;
}) {
  const { t } = useTranslation();
  return <div className="docx-review-summary" data-orientation={orientation}>
    <span className="docx-review-summary__authors">{authors.filter((author, index) => authors.findIndex(candidate => candidate.name === author.name) === index)
      .slice(0, 3).map(author => <ReviewAvatar key={author.name} author={author} />)}</span>
    <span className="docx-review-summary__count" title={t('trackedChanges.changeCount', { count })}>{orientation === 'horizontal' ? t('trackedChanges.changeCount', { count }) : count}</span>
    <button type="button" className="docx-review-accept-all" title={t('trackedChanges.acceptAllChanges')} aria-label={t('trackedChanges.acceptAllChanges')}
      onClick={acceptAll}>{orientation === 'horizontal' ? t('trackedChanges.acceptAll') : <svg className="docx-review-accept-all__icon" viewBox="0 0 16 16" aria-hidden="true"><path d="M1.5 8.6 4.4 11.4 10.2 4.8M7.6 11.2 8 11.4 14.4 4.8" /></svg>}</button>
  </div>;
}

export type ReviewPage = { page: { pageIndex: number }; bounds: { x: number; y: number; width: number; height: number } };
type Point = { x: number; y: number };
type Placement = ReviewAnchor & { x: number; y: number; bottom: number; end: Point };
const noopMeasure = () => {};

export function TrackedChangesReview({ children, items, options, collaboration, pages, rects, zoom, pageWidth, sidebarOpen,
  hostRef, scrollerRef, originRef, activeId, setActive, select, readOnly, acceptAll,
}: {
  children: ReactNode;
  items: ReviewSidebarItem[];
  options?: TrackedChangesOptions;
  collaboration?: DocxEditorCollaborationOptions;
  pages: readonly ReviewPage[];
  rects: ReadonlyMap<string, readonly DisplayListRect[]>;
  zoom: number;
  pageWidth: number;
  sidebarOpen: boolean;
  hostRef: RefObject<HTMLDivElement | null>;
  scrollerRef: RefObject<HTMLDivElement | null>;
  /** Element whose top-left is the origin of `pages`; defaults to the peerless pages host. */
  originRef?: RefObject<HTMLElement | null>;
  activeId: string | null;
  setActive(id: string | null): void;
  select(item: ReviewSidebarItem): void;
  readOnly: boolean;
  acceptAll(): void;
}) {
  const { t } = useTranslation();
  const dotsMode = options?.display === 'dots';
  const peers = useReviewPeers(collaboration?.presence);
  const authors = useMemo(() => items.map(item => resolveReviewAuthor(item.review!.change.author || t('trackedChanges.unknown'), item.review!.kind,
    options, peers, collaboration?.user, collaboration?.clientId)), [items, options, peers, collaboration?.user, collaboration?.clientId, t]);
  const cards = useRef(new Map<string, HTMLDivElement>());
  const dots = useRef(new Map<string, HTMLButtonElement>());
  const balloon = useRef<HTMLDivElement>(null);
  const [measurement, setMeasurement] = useState(0);
  const measureCard = useCallback((id: string, element: HTMLDivElement | null) => {
    if (cards.current.get(id) === element || (!element && !cards.current.has(id))) return;
    if (element) cards.current.set(id, element); else cards.current.delete(id);
    setMeasurement(value => value + 1);
  }, []);
  const navigate = (id: string, direction: number) => {
    const index = items.findIndex(item => item.id === id);
    const next = items[(index + direction + items.length) % items.length];
    if (next) { cancelClose(); setTransient(null); setDismissed(null); setActive(next.id); select(next); }
  };
  const [transient, setTransient] = useState<string | null>(null);
  const [dismissed, setDismissed] = useState<string | null>(null);
  const returningFocus = useRef(false);
  const hovered = useRef<string | null>(null);
  const focused = useRef<string | null>(null);
  const closeTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const cancelClose = () => { clearTimeout(closeTimer.current); };
  const open = (id: string) => { cancelClose(); setDismissed(null); setTransient(id); };
  const closeSoon = () => {
    cancelClose();
    closeTimer.current = setTimeout(() => setTransient(current => current && (hovered.current === current || focused.current === current) ? current : null), 150);
  };
  useEffect(() => () => clearTimeout(closeTimer.current), []);
  useEffect(() => { setDismissed(null); }, [activeId]);
  const shownId = transient ?? (activeId !== dismissed ? activeId : null);
  const activeIndex = items.findIndex(item => item.id === (dotsMode ? shownId : activeId));
  const active = items[activeIndex];
  const [hoveredCard, setHoveredCard] = useState<string | null>(null);
  const linked = useMemo(() => dotsMode ? [] : [...new Set([activeId, hoveredCard])]
    .filter((id): id is string => id !== null && items.some(item => item.id === id)), [dotsMode, activeId, hoveredCard, items]);
  const [layout, setLayout] = useState<{ placements: Placement[]; summary: Point; right: number; linked: Record<string, Point>;
    ambiguous: string[]; width: number; ready: boolean }>({
    placements: [], summary: { x: 0, y: 12 }, right: 0, linked: {}, ambiguous: [], width: 300, ready: false,
  });
  const ambiguous = useMemo(() => new Set(layout.ambiguous), [layout.ambiguous]);
  const laidOut = useRef(false);
  useEffect(() => {
    const host = hostRef.current;
    if (!host || dotsMode) return;
    const over = (event: MouseEvent) => {
      const target = event.target as Node;
      setHoveredCard([...cards.current].find(([, card]) => card.contains(target))?.[0] ?? null);
    };
    const leave = () => setHoveredCard(null);
    host.addEventListener('mouseover', over);
    host.addEventListener('mouseleave', leave);
    return () => { host.removeEventListener('mouseover', over); host.removeEventListener('mouseleave', leave); };
  }, [hostRef, dotsMode]);
  useEffect(() => {
    const host = hostRef.current;
    const scroller = scrollerRef.current;
    if (!host || !scroller) return;
    let live = true;
    const update = () => {
      if (!live) return;
      const parent = host.getBoundingClientRect();
      const viewport = scroller.getBoundingClientRect();
      const origin = (originRef?.current ?? host.querySelector('[data-docx-pages-host]'))?.getBoundingClientRect();
      const ox = origin ? origin.left - parent.left : (host.clientWidth - (sidebarOpen ? SIDEBAR_DOCUMENT_SHIFT * 2 : 0) - pageWidth * zoom) / 2;
      const oy = origin ? origin.top - parent.top : 48;
      const slots = new Map<number, { y: number; count: number }[]>();
      const placements = items.flatMap(item => {
        const rs = rects.get(item.id);
        const first = rs?.[0], last = rs?.[rs.length - 1];
        const page = first && pages.find(page => page.page.pageIndex === first.pageIndex);
        const lastPage = last && pages.find(page => page.page.pageIndex === last.pageIndex);
        if (!first || !last || !page || !lastPage) return [];
        const y = oy + (page.bounds.y + first.y + first.height / 2) * zoom;
        const lines = slots.get(first.pageIndex) ?? [];
        let line = lines.find(line => Math.abs(line.y - y) < 6);
        if (!line) { line = { y, count: 0 }; lines.push(line); slots.set(first.pageIndex, lines); }
        const x = ox + (page.bounds.x + page.bounds.width) * zoom + 6 + line.count++ * 22;
        const lastOnPage = rs.filter(rect => rect.pageIndex === first.pageIndex).pop() ?? first;
        return [{ id: item.id, pageIndex: first.pageIndex, x, y, top: oy + (page.bounds.y + first.y) * zoom, height: first.height * zoom,
          bottom: oy + (page.bounds.y + lastOnPage.y + lastOnPage.height) * zoom,
          lines: rs.flatMap(rect => {
            const bounds = pages.find(candidate => candidate.page.pageIndex === rect.pageIndex)?.bounds;
            return bounds ? [{ pageIndex: rect.pageIndex, top: oy + (bounds.y + rect.y) * zoom, height: rect.height * zoom }] : [];
          }),
          end: { x: ox + (lastPage.bounds.x + last.x + last.width) * zoom, y: oy + (lastPage.bounds.y + last.y + last.height) * zoom } }];
      });
      const visible = pages.find(page => oy + (page.bounds.y + page.bounds.height) * zoom > viewport.top - parent.top) ?? pages[0];
      const column = host.querySelector('.docx-unified-sidebar')?.getBoundingClientRect();
      const card = active && cards.current.get(active.id)?.getBoundingClientRect();
      const cardTops = new Map(dotsMode ? [] : [...cards.current].map(([id, element]) => [id, element.getBoundingClientRect().top - parent.top] as const));
      const linkedCards = Object.fromEntries(linked.flatMap(id => {
        const bounds = cards.current.get(id)?.getBoundingClientRect();
        return bounds ? [[id, { x: bounds.left - parent.left, y: bounds.top - parent.top + 18 }] as const] : [];
      }));
      const next = { placements, right: viewport.right - parent.left, linked: linkedCards, ambiguous: ambiguousReviewIds(placements, cardTops),
        summary: { x: dotsMode ? ox + ((visible?.bounds.x ?? 0) + (visible?.bounds.width ?? pageWidth)) * zoom + 2
          : column ? column.left - parent.left : ox + pageWidth * zoom + SIDEBAR_PAGE_GAP,
        y: Math.max(12, viewport.top - parent.top + 12) },
        width: card?.width || balloon.current?.offsetWidth || 300, ready: true };
      setLayout(previous => JSON.stringify(previous) === JSON.stringify(next) ? previous : next);
    };
    // Scrolls and re-renders lay out in the next frame, never while React commits.
    let frame: number | null = null;
    const schedule = () => { if (frame === null) frame = requestAnimationFrame(() => { frame = null; update(); }); };
    if (laidOut.current) schedule(); else { laidOut.current = true; update(); }
    scroller.addEventListener('scroll', schedule, { passive: true });
    const view = host.ownerDocument.defaultView;
    view?.addEventListener('resize', update);
    host.addEventListener('transitionend', update);
    const observer = new ResizeObserver(update);
    observer.observe(host); observer.observe(scroller);
    for (const card of cards.current.values()) observer.observe(card);
    if (balloon.current) observer.observe(balloon.current);
    const moves = new MutationObserver(update);
    const column = host.querySelector('.docx-unified-sidebar');
    if (column) moves.observe(column, { subtree: true, attributes: true, attributeFilter: ['style'] });
    return () => {
      live = false;
      if (frame !== null) cancelAnimationFrame(frame);
      scroller.removeEventListener('scroll', schedule); view?.removeEventListener('resize', update); host.removeEventListener('transitionend', update);
      observer.disconnect(); moves.disconnect();
    };
  }, [items, rects, pages, zoom, pageWidth, sidebarOpen, hostRef, scrollerRef, originRef, active, dotsMode, measurement, linked]);
  const position = layout.placements.find(placement => placement.id === active?.id);
  const below = !!position && position.x + 34 + layout.width > layout.right - 8;
  const balloonPosition = position ? below ? { x: Math.max(8, position.x + 18 - layout.width), y: Math.max(position.y + 9, position.bottom) + 6 }
    : { x: position.x + 34, y: position.y - 26 } : null;
  const target = dotsMode ? balloonPosition && (below ? { x: position!.x + 9, y: balloonPosition.y } : { x: balloonPosition.x, y: balloonPosition.y + 18 }) : null;
  const leaders = dotsMode ? position && target ? [{ id: active!.id, path: `M${position.end.x} ${position.end.y} L${position.x + 9} ${position.y} L${target.x} ${target.y}` }] : []
    : linked.flatMap(id => {
      const anchor = layout.placements.find(placement => placement.id === id), card = layout.linked[id];
      return anchor && card ? [{ id, path: `M${anchor.end.x} ${anchor.end.y} H${anchor.x - 6} L${card.x} ${card.y}` }] : [];
    });
  const escape = () => {
    cancelClose(); setTransient(null); setDismissed(shownId); setActive(null);
    returningFocus.current = true;
    if (shownId) dots.current.get(shownId)?.focus({ preventScroll: true });
    returningFocus.current = false;
  };
  useEffect(() => {
    if (!dotsMode || activeIndex < 0) return;
    const keydown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); escape(); }
    };
    const host = hostRef.current;
    host?.addEventListener('keydown', keydown, true);
    return () => host?.removeEventListener('keydown', keydown, true);
  });
  return <TrackedChangesContext.Provider value={{ options, peers, user: collaboration?.user, clientId: collaboration?.clientId, items, navigate, measureCard, ambiguous }}>
    {children}
    <div className="docx-review-layer" onMouseDown={event => event.stopPropagation()}>
      {leaders.map(({ id, path }) => <svg key={id} className="docx-review-leader" data-review-item={id} aria-hidden="true"
        style={reviewAuthorStyle(authors[items.findIndex(item => item.id === id)]!)}><path className="docx-review-leader__path" d={path} /></svg>)}
      {items.length > 0 && !readOnly && layout.ready && <div className="docx-review-summary-position" style={{ left: layout.summary.x, top: layout.summary.y }}>
        <ReviewSummary authors={authors} count={items.length} orientation={dotsMode ? 'vertical' : 'horizontal'} acceptAll={acceptAll} />
      </div>}
      {dotsMode && pages.map(page => {
        const placements = layout.placements.filter(placement => placement.pageIndex === page.page.pageIndex);
        const left = placements[0]?.x ?? 0;
        return <div key={page.page.pageIndex} className="docx-review-gutter" style={{ left }}>
        {placements.map(placement => {
          const index = items.findIndex(item => item.id === placement.id);
          const item = items[index], author = authors[index];
          if (!item || !author) return null;
          const expanded = item.id === shownId;
          return <button key={item.id} type="button" className="docx-review-dot" data-kind={author.kind} data-active={expanded}
            data-revision-id={item.review!.change.revisionId} aria-label={t('trackedChanges.changeBy', { author: author.name })} aria-expanded={expanded}
            style={{ left: placement.x - left, top: placement.y - 9, ...reviewAuthorStyle(author) }} ref={element => { if (element) dots.current.set(item.id, element); else dots.current.delete(item.id); }}
            onMouseEnter={() => { hovered.current = item.id; open(item.id); }} onMouseLeave={() => { hovered.current = null; closeSoon(); }}
            onFocus={() => { focused.current = item.id; if (!returningFocus.current) open(item.id); }} onBlur={event => { focused.current = null; if (!balloon.current?.contains(event.relatedTarget as Node)) closeSoon(); }}
            onClick={() => { cancelClose(); setDismissed(null); setTransient(null); setActive(item.id); select(item); }}>{author.icon}</button>;
        })}
      </div>;
      })}
      {dotsMode && active && balloonPosition && <div ref={balloon} className="docx-review-balloon" style={{ left: balloonPosition.x, top: balloonPosition.y }}
        onMouseEnter={() => { hovered.current = active.id; cancelClose(); }} onMouseLeave={() => { hovered.current = null; closeSoon(); }} onFocus={() => { focused.current = active.id; cancelClose(); }}
        onBlur={event => { if (!event.currentTarget.contains(event.relatedTarget as Node)) { focused.current = null; closeSoon(); } }}>
        <Fragment key={active.id}>{active.render({ isExpanded: true, onToggleExpand: () => setActive(active.id), measureRef: noopMeasure })}</Fragment>
      </div>}
    </div>
  </TrackedChangesContext.Provider>;
}
