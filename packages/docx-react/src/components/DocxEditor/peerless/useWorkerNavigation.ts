import { useCallback, useEffect, useRef, useState, type RefObject } from 'react';
import type { DocxEditorProps } from '../../DocxEditor';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { createPeerlessShellController } from './shellController';
import { workerPageAtViewport } from './pageNavigation';
import type { DocxOutlineHeading } from '@betteroffice/docx/yrs';

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
  const scrollToParagraph = useCallback(async (paraId: string) => {
    const current = binding();
    if (!current) return false;
    const target = await controller.read({ kind: 'navigationTarget', story: 'body', paraId });
    if (typeof target === 'string' || current !== binding()) return false;
    const result = current.commands.submit({ kind: 'select', payload: { selection: { anchor: target.loc, focus: target.loc } } }, { source: 'selection', fence: true });
    if (result.status !== 'submitted') return false;
    const outcome = await result.ticket.ack;
    if (current !== binding() || (outcome.status !== 'applied' && outcome.status !== 'noop')) return false;
    if (await result.ticket.presented !== 'presented' || current !== binding()) return false;
    const caret = current.viewStore.getSnapshot().geometry?.caret;
    if (caret && scroller.current) scroller.current.scrollTop += (host.current?.getBoundingClientRect().top ?? 0) - scroller.current.getBoundingClientRect().top + caret.y * latest.current.zoom - 24;
    return true;
  }, [binding, controller, host, scroller]);
  return { showOutline, setShowOutline, headings, pageInfo, readCurrentPage, scrollToPage, scrollToParagraph };
}
