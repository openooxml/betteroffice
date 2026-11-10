import { memo, useLayoutEffect, useMemo, useRef, useState, type RefObject } from 'react';
import type { TFunction } from '@betteroffice/docx-i18n';
import { displayPageRevision } from '@betteroffice/docx/layout/render';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { geometryMatchesDocument, type PeerlessFrameGeometry } from './viewStore';
import { createUnbuiltPageMirror, scheduleUnbuiltMirrorIdle, type MirrorPageText, type MirrorStamp, type PageTexts, type UnbuiltMirrorSnapshot } from './unbuiltPageMirror';

type Page = PeerlessFrameGeometry['pages'][number];
export type ReadPageTexts = (pages: number[], stamp: MirrorStamp) => Promise<PageTexts | null>;
const BLOCKS = 'p';
const normalize = (text: string | null) => (text ?? '').replace(/\s+/g, ' ').trim();
type FocusedBlock = { pageIndex: number; ordinal: number; text: string; active: Element | null; selection: boolean };

function focusMirrorBlock(element: HTMLElement, cleanups: Set<() => void>) {
  if (element.ownerDocument.activeElement === element) return;
  const added = !element.hasAttribute('tabindex');
  const clear = () => {
    element.removeEventListener('blur', clear);
    if (added && element.getAttribute('tabindex') === '-1') element.removeAttribute('tabindex');
    cleanups.delete(clear);
  };
  cleanups.add(clear);
  element.addEventListener('blur', clear, { once: true });
  if (added) element.tabIndex = -1;
  element.focus({ preventScroll: true });
  if (element.ownerDocument.activeElement !== element) clear();
}

const LINE_CHARS = 120;

export function buildUnbuiltMirrorSection(page: MirrorPageText, t: TFunction, owner: Document = document) {
  const section = owner.createElement('section');
  section.setAttribute('aria-label', t('a11y.pageLabel', { number: page.pageIndex + 1 }));
  section.dataset.docxUnbuiltPage = String(page.pageIndex);
  let nodes = 1;
  for (const paragraph of page.paragraphs) {
    if (!paragraph.text.trim()) continue;
    const block = owner.createElement('p');
    block.dir = 'auto';
    block.style.margin = '0';
    block.tabIndex = -1;
    block.textContent = paragraph.text;
    section.appendChild(block);
    nodes += 2 + Math.ceil(paragraph.text.length / LINE_CHARS);
  }
  return { value: section, nodes };
}

export const UnbuiltMirrorSection = memo(function UnbuiltMirrorSection({ page, value, remember }: { page: Page; value: HTMLElement; remember(section: HTMLElement): void }) {
  const ref = useRef<HTMLElement>(null);
  const installed = useRef<HTMLElement | null>(null);
  useLayoutEffect(() => {
    const section = ref.current!;
    const active = section.ownerDocument.activeElement;
    const previous = active instanceof HTMLElement && section.contains(active) ? active : null;
    const ordinal = previous ? Array.from(section.querySelectorAll<HTMLElement>(BLOCKS)).indexOf(previous) : -1;
    if (installed.current !== value) {
      section.replaceChildren(...Array.from(value.childNodes));
      installed.current = value;
    }
    if (previous) {
      const replacement = section.querySelectorAll<HTMLElement>(BLOCKS)[ordinal] ?? section;
      replacement.tabIndex = -1;
      replacement.focus({ preventScroll: true });
    }
    return () => { remember(section); };
  }, [value, remember]);
  return <section ref={ref} data-docx-unbuilt-page={page.page.pageIndex} aria-label={value.getAttribute('aria-label') ?? undefined}
    style={{ position: 'absolute', left: page.bounds.x, top: page.bounds.y, width: page.bounds.width, height: page.bounds.height,
      opacity: 0, pointerEvents: 'none', fontSize: 12, lineHeight: 1.2 }} />;
}, (previous, next) => previous.value === next.value && previous.remember === next.remember &&
  previous.page.page.pageIndex === next.page.page.pageIndex &&
  previous.page.bounds.x === next.page.bounds.x && previous.page.bounds.y === next.page.bounds.y &&
  previous.page.bounds.width === next.page.bounds.width && previous.page.bounds.height === next.page.bounds.height);

export function useUnbuiltPageMirrors({ binding, geometry, center, enabled, read, blocked, navigate, column, t, mountedPages, fullMirror = true }: {
  binding: PeerlessEditSessionBinding;
  geometry: PeerlessFrameGeometry | null;
  mountedPages: readonly Page[];
  center: number;
  enabled: boolean;
  read?: ReadPageTexts;
  blocked?: () => boolean;
  navigate(number: number): void;
  column: RefObject<HTMLDivElement | null>;
  t: TFunction;
  /** Visible pages build the full mirror; false when they keep the text mirror. */
  fullMirror?: boolean;
}) {
  const [snapshot, setSnapshot] = useState<UnbuiltMirrorSnapshot<HTMLElement>>({ sections: [], ranges: [], nodeCount: 0 });
  const latest = useRef({ read, blocked, navigate });
  latest.current = { read, blocked, navigate };
  const focused = useRef<FocusedBlock | null>(null);
  const parked = useRef<HTMLElement | null>(null);
  const focusCleanups = useRef(new Set<() => void>());
  useLayoutEffect(() => () => { for (const clear of focusCleanups.current) clear(); }, []);
  const remember = useMemo(() => (section: HTMLElement, target: Element | null = section.ownerDocument.activeElement, selection = false) => {
    if (!target || !section.contains(target)) return;
    const block = target.closest<HTMLElement>(BLOCKS);
    if (!block) return;
    const paragraphs = Array.from(section.querySelectorAll<HTMLElement>(BLOCKS)).filter((paragraph) => normalize(paragraph.textContent));
    focused.current = { pageIndex: Number(section.dataset.docxUnbuiltPage), ordinal: paragraphs.indexOf(block), text: normalize(block.textContent),
      active: section.ownerDocument.activeElement, selection };
    parked.current = null;
  }, []);
  const scheduler = useMemo(() => createUnbuiltPageMirror({
    idle: scheduleUnbuiltMirrorIdle,
    query: (pages, stamp) => latest.current.read?.(pages, stamp) ?? Promise.resolve(null),
    blocked: () => {
      const admission = binding.admission.getSnapshot();
      const view = binding.viewStore.getSnapshot();
      return !!(admission.pendingInputs || admission.originalUnackedIntents || admission.heldIntents || admission.compositionSlot !== null ||
        !geometryMatchesDocument(view) || view.presentedThroughSeq < binding.session.state.throughSeq || latest.current.blocked?.());
    },
    build: (page) => buildUnbuiltMirrorSection(page, t, column.current?.ownerDocument),
    focusedPage: () => {
      const active = column.current?.ownerDocument.activeElement;
      const section = active?.closest<HTMLElement>('[data-docx-unbuilt-page]');
      return section && column.current?.contains(section) ? Number(section.dataset.docxUnbuiltPage) : null;
    },
    publish: setSnapshot,
  }), [binding, column, t]);
  useLayoutEffect(() => { scheduler.resume(); return () => scheduler.dispose(); }, [scheduler]);
  useLayoutEffect(() => {
    const identity = geometry?.identity;
    const mounted = new Set(mountedPages.map(({ page }) => page.pageId));
    scheduler.update({ identity: identity ? `${identity.session}:${identity.incarnation}:${identity.epoch}` : '',
      frameEpoch: geometry?.frameEpoch ?? 0, layoutRevision: geometry?.layoutRevision ?? 0, center,
      enabled: enabled && !!read,
      pages: enabled && read ? (geometry?.pages ?? []).map(({ page, displayPage }) => ({ id: page.pageId, pageIndex: page.pageIndex,
        source: displayPage, revision: displayPage ? displayPageRevision(displayPage) : 0, unbuilt: displayPage?.unbuilt === true && !mounted.has(page.pageId) })) : [] });
  }, [scheduler, geometry, center, enabled, read, mountedPages]);
  useLayoutEffect(() => {
    if (!enabled) return;
    const host = column.current?.querySelector<HTMLElement>('[data-docx-pages-host]');
    if (!host) return;
    const owner = host.ownerDocument;
    const selected = new WeakSet<HTMLElement>();
    const activate = (event: Event) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const section = target.closest<HTMLElement>('[data-docx-unbuilt-page]');
      if (!section || !host.contains(section)) return;
      if (event.type === 'focusin') remember(section);
      latest.current.navigate(Number(section.dataset.docxUnbuiltPage) + 1);
    };
    const sectionOf = (node: Node | null | undefined) =>
      (node instanceof Element ? node : node?.parentElement)?.closest<HTMLElement>('[data-docx-unbuilt-page]') ?? null;
    let settle: ReturnType<typeof setTimeout> | null = null;
    const check = () => {
      settle = null;
      const selection = owner.getSelection();
      if (!selection || selection.rangeCount !== 1) return;
      const section = sectionOf(selection.anchorNode);
      if (!section || sectionOf(selection.focusNode) !== section || !host.contains(section) || section.contains(owner.activeElement)) return;
      const anchor = selection.anchorNode;
      remember(section, anchor instanceof Element ? anchor : anchor?.parentElement ?? null, true);
      if (selected.has(section)) return;
      selected.add(section);
      latest.current.navigate(Number(section.dataset.docxUnbuiltPage) + 1);
    };
    const select = () => {
      if (settle !== null) clearTimeout(settle);
      settle = setTimeout(check, 50);
    };
    const leave = (event: FocusEvent) => {
      if (focused.current && event.target !== focused.current.active && event.target !== parked.current) {
        focused.current = null;
        parked.current = null;
      }
    };
    host.addEventListener('focusin', activate);
    host.addEventListener('click', activate);
    owner.addEventListener('focusin', leave);
    owner.addEventListener('selectionchange', select);
    return () => {
      host.removeEventListener('focusin', activate);
      host.removeEventListener('click', activate);
      owner.removeEventListener('focusin', leave);
      owner.removeEventListener('selectionchange', select);
      if (settle !== null) clearTimeout(settle);
    };
  }, [column, enabled, remember]);
  useLayoutEffect(() => {
    const wanted = focused.current;
    const host = column.current?.querySelector<HTMLElement>('[data-docx-pages-host]');
    if (!enabled) {
      if (wanted) column.current?.querySelector<HTMLTextAreaElement>('.paged-editor__input')?.focus({ preventScroll: true });
      focused.current = null;
      parked.current = null;
      for (const clear of focusCleanups.current) clear();
      return;
    }
    if (!wanted || !host || host.querySelector(`[data-docx-unbuilt-page="${wanted.pageIndex}"]`)) return;
    const handover = () => {
      if (focused.current !== wanted) return true;
      const active = host.ownerDocument.activeElement;
      if (active !== wanted.active && active !== parked.current && active !== host.ownerDocument.body) {
        focused.current = null;
        parked.current = null;
        return true;
      }
      const mirrorHost = host.querySelector<HTMLElement>(`[data-page-index="${wanted.pageIndex}"] .canvas-page-mirror`);
      const root = mirrorHost?.firstElementChild as HTMLElement | null;
      const ready = fullMirror ? mirrorHost?.dataset.pageChrome === 'chrome' : mirrorHost?.dataset.pageChrome !== undefined;
      const paragraphs = root && ready ? Array.from(root.querySelectorAll<HTMLElement>('[role="paragraph"], p'))
        .map((element) => ({ element, text: normalize(element.textContent) })).filter(({ text }) => text) : [];
      // A page shell shows the mirror root before its paragraphs fill it in place: wait for the paragraph itself.
      const block = paragraphs.slice(Math.max(0, wanted.ordinal)).find(({ text }) => text === wanted.text)
        ?? paragraphs.find(({ text }) => text === wanted.text);
      if (!block || geometry?.pages.find(({ page }) => page.pageIndex === wanted.pageIndex)?.displayPage?.unbuilt) {
        if (wanted.selection) return false;
        const target = mirrorHost ?? host;
        parked.current = target;
        focusMirrorBlock(target, focusCleanups.current);
        return false;
      }
      if (wanted.selection) {
        const selection = host.ownerDocument.getSelection();
        if (selection) {
          const range = host.ownerDocument.createRange();
          range.selectNodeContents(block.element);
          selection.removeAllRanges();
          selection.addRange(range);
        }
      } else focusMirrorBlock(block.element, focusCleanups.current);
      focused.current = null;
      parked.current = null;
      return true;
    };
    if (handover()) return;
    const observer = new MutationObserver(() => { if (handover()) observer.disconnect(); });
    observer.observe(host, { childList: true, subtree: true });
    return () => observer.disconnect();
  });
  return useMemo(() => ({ snapshot, remember, navigate }), [snapshot, remember, navigate]);
}
