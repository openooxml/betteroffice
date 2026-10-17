import { useCallback, useEffect, useRef, useState, type RefObject } from 'react';
import type { DocxEditorProps } from '../../DocxEditor';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { WorkerReadSupersededError, type createPeerlessShellController } from './shellController';
import { workerPageAtViewport } from './pageNavigation';
import type { DocxCommandSelection, DocxOutlineHeading, ResidentReviewTarget } from '@betteroffice/docx/yrs';
import { createDisplayListQueries, loadRustDisplayListQueryEngine } from '@betteroffice/docx/layout/render';
import { createCanvasHostProjector } from '@betteroffice/docx/plugin-api/RenderedDomContext';

export function useWorkerNavigation(controller: ReturnType<typeof createPeerlessShellController>,
  binding: () => PeerlessEditSessionBinding | null, scroller: RefObject<HTMLDivElement | null>,
  host: RefObject<HTMLDivElement | null>, zoom: number, props: DocxEditorProps) {
  const [showOutline, setShowOutline] = useState(props.showOutline ?? false);
  useEffect(() => { setShowOutline(props.showOutline ?? false); }, [props.showOutline]);
  const [headings, setHeadings] = useState<DocxOutlineHeading[]>([]);
  const [pageInfo, setPageInfo] = useState({ currentPage: 1, totalPages: 0, visible: false, scrollLeft: 0 });
  const latest = useRef({ zoom }); latest.current = { zoom };
  const readCurrentPage = useCallback(() => {
    const element = scroller.current;
    const pages = controller.getSnapshot().view?.geometry?.pages ?? [];
    if (!element || !pages.length) return 1;
    const offset = (host.current?.getBoundingClientRect().top ?? 0) - element.getBoundingClientRect().top + element.scrollTop;
    return workerPageAtViewport(pages, element.scrollTop - offset, element.clientHeight, latest.current.zoom);
  }, [controller, host, scroller]);
  useEffect(() => {
    const element = scroller.current;
    if (!element) return;
    let fade: ReturnType<typeof setTimeout> | undefined;
    const update = (visible = false) => {
      const count = controller.getSnapshot().view?.geometry?.pages.length ?? 0;
      const currentPage = readCurrentPage();
      setPageInfo((previous) => previous.currentPage === currentPage && previous.totalPages === count &&
        previous.visible === visible && previous.scrollLeft === element.scrollLeft ? previous :
        { currentPage, totalPages: count, visible, scrollLeft: element.scrollLeft });
    };
    const scroll = () => { update(true); clearTimeout(fade); fade = setTimeout(() => update(false), 600); };
    update();
    const unsubscribe = controller.subscribe(() => update());
    element.addEventListener('scroll', scroll, { passive: true });
    window.addEventListener('resize', scroll);
    return () => { unsubscribe(); clearTimeout(fade); element.removeEventListener('scroll', scroll); window.removeEventListener('resize', scroll); };
  }, [controller, readCurrentPage, scroller, zoom]);
  const state = controller.getSnapshot();
  useEffect(() => {
    let current = true;
    setHeadings([]);
    if (!showOutline || state.status !== 'ready') return;
    void controller.ref.readParagraphs({ view: 'accepted', paraIds: [] }).then(async (version) => {
      if (!version.ok) return;
      const result = await controller.read({ kind: 'headings', expectVersion: version.version });
      if (current && result) setHeadings(result);
    }).catch((error) => { if (current) props.onError?.(error instanceof Error ? error : new Error(String(error))); });
    return () => { current = false; };
  }, [controller, showOutline, state.status, state.view?.identity.session, state.view?.version, props.onError]);
  const scrollToPage = useCallback((number: number) => {
    if (!Number.isInteger(number) || number < 1) return;
    const page = controller.getSnapshot().view?.geometry?.pages.find(({ page }) => page.pageIndex + 1 === number);
    const element = scroller.current;
    if (!page || !element) return;
    element.scrollTop += (host.current?.getBoundingClientRect().top ?? 0) - element.getBoundingClientRect().top + page.bounds.y * latest.current.zoom;
  }, [controller, host, scroller]);
  const navigate = useCallback(async (operation: (current: PeerlessEditSessionBinding) => Promise<boolean>) => {
    const current = binding();
    if (!current || controller.getSnapshot().status !== 'ready') return false;
    const identity = current.viewStore.getSnapshot().identity;
    const isCurrent = () => current === binding() && controller.getSnapshot().status === 'ready' &&
      identity.session === current.viewStore.getSnapshot().identity.session && identity.incarnation === current.viewStore.getSnapshot().identity.incarnation &&
      identity.epoch === current.viewStore.getSnapshot().identity.epoch;
    try { const revealed = await operation(current); return isCurrent() && revealed; }
    catch (error) {
      if (error instanceof WorkerReadSupersededError) return false;
      if (error && typeof error === 'object' && 'code' in error && error.code === 'disposed' &&
        !isCurrent()) return false;
      throw error;
    }
  }, [binding, controller]);
  const revealSelection = useCallback(async (current: PeerlessEditSessionBinding, selection: DocxCommandSelection) => {
    const identity = current.viewStore.getSnapshot().identity;
    const isCurrent = () => current === binding() && identity.session === current.viewStore.getSnapshot().identity.session &&
      identity.incarnation === current.viewStore.getSnapshot().identity.incarnation && identity.epoch === current.viewStore.getSnapshot().identity.epoch;
    const select = async (value: DocxCommandSelection) => {
      const throughSeq = current.viewStore.getSnapshot().throughSeq;
      const result = current.commands.submit({ kind: 'select', payload: { selection: value } }, { source: 'selection', fence: true });
      if (result.status !== 'submitted') return false;
      const [outcome, presented] = await Promise.all([result.ticket.ack, result.ticket.presented]);
      if (!isCurrent() || (outcome.status !== 'applied' && outcome.status !== 'noop')) return false;
      const seq = outcome.seq ?? result.ticket.seq ?? throughSeq;
      return presented === 'presented' && current.viewStore.getSnapshot().throughSeq <= seq;
    };
    if (!await select({ anchor: selection.anchor, focus: selection.anchor })) return false;
    const shown = current.viewStore.getSnapshot().geometry?.caret;
    if (!shown || !scroller.current || !host.current) return false;
    if (selection.anchor.story !== selection.focus.story || selection.anchor.paraId !== selection.focus.paraId || selection.anchor.offset !== selection.focus.offset) {
      if (!await select({ anchor: selection.focus, focus: selection.anchor })) return false;
    }
    if (!scroller.current || !host.current) return false;
    scroller.current.scrollTop += host.current.getBoundingClientRect().top - scroller.current.getBoundingClientRect().top + shown.y * latest.current.zoom - 24;
    return true;
  }, [binding, host, scroller]);
  const scrollToParagraph = useCallback((paraId: string, story = 'body') => navigate(async current => {
    const target = await controller.read({ kind: 'navigationTarget', story, paraId });
    if (typeof target === 'string' || current !== binding()) return false;
    return revealSelection(current, { anchor: target.loc, focus: target.loc });
  }), [binding, controller, navigate, revealSelection]);
  const scrollToBookmark = useCallback((name: string, story = 'body') => navigate(async current => {
    const target = await controller.read({ kind: 'bookmarkTarget', story, name });
    if (!target || typeof target === 'string' || current !== binding()) return false;
    return revealSelection(current, { anchor: target.loc, focus: target.loc });
  }), [binding, controller, navigate, revealSelection]);
  /** Selects the first range a comment or revision anchors and scrolls its start into view; false when it does not exist. */
  const scrollToReviewTarget = useCallback((target: ResidentReviewTarget) => navigate(async current => {
    const range = (await controller.read({ kind: 'reviewTarget', target }))?.ranges[0];
    if (!range || current !== binding()) return false;
    return revealSelection(current, { anchor: { story: range.story, ...range.start }, focus: { story: range.story, ...range.end } });
  }), [binding, controller, navigate, revealSelection]);
  const revealParagraph = useCallback(async (paraId: string, story: string, focus = false) => {
    if (focus) return scrollToParagraph(paraId, story);
    const current = binding();
    if (!current || !host.current || !scroller.current) return false;
    const target = await controller.read({ kind: 'navigationTarget', story, paraId });
    if (typeof target === 'string' || current !== binding()) return false;
    const native = await loadRustDisplayListQueryEngine();
    const deadline = Date.now() + 30000;
    while (current === binding() && Date.now() < deadline) {
      const frame = current.viewStore.getSnapshot().geometry;
      if (!frame || !host.current || !scroller.current) return false;
      const queries = createDisplayListQueries({ pages: frame.pages.map(page => page.displayPage!).filter(Boolean) }, native);
      try {
        const rect = queries.anchorRect(target.position);
        const point = rect && createCanvasHostProjector(host.current, queries, latest.current.zoom).projectRect(rect);
        if (point) {
          scroller.current.scrollTop += host.current.getBoundingClientRect().top - scroller.current.getBoundingClientRect().top + point.y * latest.current.zoom - 24;
          return true;
        }
      } finally { queries.dispose(); }
      const page = frame.pages.find(({ displayPage }) => displayPage?.positionSpan &&
        displayPage.positionSpan[0] <= target.position && target.position <= displayPage.positionSpan[1]);
      if (!page || !current.setDisplayWindow) return false;
      const presented = new Promise<boolean>(resolve => {
        let remove = () => {};
        const timer = setTimeout(() => { remove(); resolve(false); }, Math.max(0, deadline - Date.now()));
        const check = () => {
          if (current === binding() && current.viewStore.getSnapshot().geometry === frame) return;
          clearTimeout(timer); remove(); resolve(current === binding());
        };
        remove = controller.subscribe(check);
        check();
      });
      scrollToPage(page.page.pageIndex + 1);
      current.setDisplayWindow(page.page.pageIndex, page.page.pageIndex + 1);
      if (!await presented) return false;
    }
    return false;
  }, [binding, controller, host, scroller, scrollToPage, scrollToParagraph]);
  return { showOutline, setShowOutline, headings, pageInfo, readCurrentPage, scrollToPage, scrollToParagraph, scrollToBookmark, scrollToReviewTarget, revealParagraph };
}
