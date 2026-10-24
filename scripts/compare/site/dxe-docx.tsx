import '@docx-editor.dev/react/styles.css';
import { createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor, type DocxEditorRef } from '@docx-editor.dev/react';
import { defaultFonts } from '@docx-editor.dev/fonts';
import { docName, expose, fetchBytes, firstPaint, rectOf, scrollByPx, scrollFraction, untilStable, visible } from '../common/probe';

const ref = createRef<DocxEditorRef>();
const sleep = (ms: number) => new Promise((done) => setTimeout(done, ms));
const pages = () => [...document.querySelectorAll('.docx-page:not(.docx-editor__loading-page)')];
const count = () => {
  return ref.current?.getEditor()?.getTotalPages() ?? 0;
};
const bytes = await fetchBytes(`./docs/${docName}.docx`);
const t0 = performance.now();
const extendedFonts = import.meta.env.VITE_CMP_EXTENDED_FONTS === '1';
const fonts = extendedFonts
  ? JSON.parse(await (await fetch('./preloaded-fonts.json')).text(), (_key, value) => value?.__fontBytes ? Uint8Array.from(atob(value.__fontBytes), (c) => c.charCodeAt(0)) : value)
  : await defaultFonts();
let onReady!: () => void;
const editorReady = new Promise<void>((done) => { onReady = done; });
const ready = (async () => {
  const firstPage = firstPaint(() => !ref.current?.snapshot().isOpening && pages().some((page) => { const r = page.getBoundingClientRect(); return r.width > 0 && r.bottom > 180 && r.top < innerHeight; }), t0).then((ms) => (cmp.firstPageMs = ms));
  await editorReady;
  for (const deadline = performance.now() + 600_000; ;) {
    const snapshot = ref.current?.snapshot();
    if (snapshot?.parseError || snapshot?.openFailure) throw new Error('document open refused: ' + JSON.stringify(snapshot.openFailure ?? snapshot.parseError));
    if (snapshot && !snapshot.isLoading && !snapshot.isOpening && snapshot.page.total > 0) {
      cmp.info.editable = snapshot.editable;
      cmp.info.fontMeasurement = (ref.current?.getEditor() as any)?.fontMeasurement?.();
      cmp.info.readiness = 'snapshot has pages and is neither loading nor opening';
      break;
    }
    if (performance.now() > deadline) throw new Error('document never became mounted');
    await sleep(50);
  }
  cmp.openMs = Math.round(performance.now() - t0);
  await firstPage;
  await document.fonts.ready;
})();
const layoutDone = (async () => {
  await ready;
  for (const deadline = performance.now() + 600_000; ;) {
    const state = (ref.current?.getEditor() as any)?.fontMeasurement?.();
    cmp.info.finalFontMeasurement = state;
    if (!state?.resolving) break;
    if (performance.now() > deadline) throw new Error('font resolution did not finish');
    await sleep(50);
  }
  const pending = () => document.querySelectorAll('[data-docx-layout-pending]').length;
  // Complete: no layout-pending element and a stable page count for 3 s. Stalled: still pending with no page-count change for 120 s.
  // Timeout: neither within 300 s of the open start; the window then ends there.
  let pages = count(), flag = pending(), countAt = performance.now(), quietAt = countAt, state = 'timeout';
  const timeline: [number, number, number][] = [[Math.round(countAt - t0), pages, flag]];
  for (const deadline = t0 + 300_000; performance.now() < deadline;) {
    await sleep(100);
    const c = count(), p = pending(), now = performance.now();
    if (c !== pages || p !== flag) {
      if (timeline.length < 400) timeline.push([Math.round(now - t0), c, p]);
      if (c !== pages) countAt = now;
      quietAt = now; pages = c; flag = p;
    }
    if (!p && c > 0 && now - quietAt >= 3000) { state = 'complete'; break; }
    if (p && now - countAt >= 120_000) { state = 'stalled'; break; }
  }
  Object.assign(cmp.info, { layoutState: state, layoutTimeline: timeline, lastPageCountChangeMs: Math.round(countAt - t0), layoutPendingAtEnd: pending() });
  cmp.layoutDoneMs = Math.round((state === 'complete' ? quietAt : performance.now()) - t0);
  cmp.layoutStable = state === 'complete';
  return pages;
})();
const cmp = expose('dxe', t0, {
  pageCount: count,
  async hasText(text) { return (ref.current?.getEditor()?.findMatches(text, { matchCase: true }).length ?? 0) > 0; },
  pageRects: () => visible(pages().map(rectOf)),
  async toMiddle() {
    const result = await scrollFraction(0.5);
    await sleep(500);
    const center = innerHeight / 2;
    const page = pages().find((p) => { const r = p.getBoundingClientRect(); return r.top <= center && r.bottom >= center; }) as HTMLElement | undefined;
    return { ...result, page: page ? Number(page.dataset.pageIndex) + 1 : null };
  },
  scrollBy: scrollByPx,
  focus: () => ref.current?.focus(),
}, ready, layoutDone, { bytes: bytes.byteLength, fontFailures: fonts.failures, fontPreset: extendedFonts ? 'all defaults plus prefetched pinned catalog' : 'eager 20 default faces' });
createRoot(document.getElementById('root')!).render(
  <DocxEditor ref={ref} fonts={fonts} document={bytes.buffer as ArrayBuffer} mode="edit" onReady={() => onReady()} onFontError={(error) => cmp.errors.push(`font: ${String((error as any)?.message ?? error).slice(0, 200)}`)} />,
);
