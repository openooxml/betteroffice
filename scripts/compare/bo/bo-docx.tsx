import '@betteroffice/docx-react/styles.css';
import { createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor, type DocxEditorRef } from '@betteroffice/docx-react';
import * as docxLayout from '@betteroffice/docx/layout';
import * as docxUtils from '@betteroffice/docx/utils';
import { docName, expose, fetchBytes, firstPaint, rectOf, scrollByPx, scrollFraction, visible } from '../common/probe';
(docxLayout as any).configureDefaultFonts?.({ load: () => import('@betteroffice/fonts') });
(docxUtils as any).setGoogleFontsEnabled?.(false);

const ref = createRef<DocxEditorRef>();
const sleep = (ms: number) => new Promise((done) => setTimeout(done, ms));
let refused: string | undefined;
async function paragraphs() {
  const result = await ref.current!.readParagraphs({ view: 'accepted' });
  if (!result.ok) { refused = JSON.stringify(result).slice(0, 200); throw new Error('readParagraphs refused'); }
  return result.paragraphs;
}
const pages = () => [...document.querySelectorAll('[data-page-id]')];
const bytes = await fetchBytes(`./docs/${docName}.docx`);
const t0 = performance.now();
let mounted!: () => void;
const rendered = new Promise<void>((done) => { mounted = done; });
const ready = (async () => {
  await rendered;
  const firstPage = firstPaint(() => pages().length > 0, t0).then((ms) => (cmp.firstPageMs = ms));
  for (const deadline = performance.now() + 300_000; ;) {
    try { if ((await paragraphs()).length > 0) break; } catch {}
    // A refused read (e.g. a view the doc cannot serve) still counts once pages are laid out; the refusal is recorded.
    if (refused && (ref.current?.getTotalPages() ?? 0) > 0) { cmp.info.readRefused = refused; cmp.info.readRefusedAtMs = Math.round(performance.now() - t0); break; }
    if (performance.now() > deadline) throw new Error('document never became readable');
    await sleep(50);
  }
  cmp.openMs = Math.round(performance.now() - t0);
  await firstPage;
  await document.fonts.ready;
})();
const layoutDone = (async () => {
  await ready;
  const count = await ref.current!.whenLayoutComplete({ timeoutMs: 600_000 });
  cmp.layoutDoneMs = Math.round(performance.now() - t0);
  cmp.layoutStable = true;
  return count;
})();
const cmp = expose('bo', t0, {
  pageCount: () => ref.current?.getTotalPages() ?? 0,
  async hasText(text) {
    try { if ((await paragraphs()).some((p) => p.text.includes(text))) return true; } catch {}
    return (await ref.current!.search(text)).total > 0;
  },
  pageRects: () => visible(pages().map(rectOf)),
  async toMiddle() {
    const result = await scrollFraction(0.5);
    await sleep(500);
    return { ...result, page: ref.current?.getCurrentPage() };
  },
  scrollBy: scrollByPx,
  // Read-path probe: each read raced against a timeout, so a refusal, an error and no reply at all are told apart.
  async diagRead(ms = 8000) {
    const race = (p: Promise<any>) => Promise.race([
      p.then((r) => ({ done: true, ok: Boolean(r?.ok), n: r?.ok ? r.paragraphs.length : null, refusal: r?.ok ? null : JSON.stringify(r).slice(0, 300) }),
        (e) => ({ done: true, error: String(e).slice(0, 200) })),
      sleep(ms).then(() => ({ done: false })),
    ]);
    const editor = ref.current as any;
    const t0 = performance.now();
    const accepted = editor ? await race(editor.readParagraphs({ view: 'accepted' })) : { noRef: true };
    const t1 = performance.now();
    const plain = editor ? await race(editor.readParagraphs({})) : { noRef: true };
    const t2 = performance.now();
    return { alertText: [...document.querySelectorAll('[role=alert]')].map((e) => (e.textContent ?? '').slice(0, 300)), onErrors: cmp.info.onErrors ?? null, ref: Boolean(editor), pages: editor?.getTotalPages?.() ?? null, dom: document.querySelectorAll('[data-page-id]').length,
      busy: document.querySelector('[data-peerless-docx]')?.getAttribute('aria-busy') ?? null, alerts: document.querySelectorAll('[role=alert]').length,
      accepted: { ...accepted, ms: Math.round(t1 - t0) }, plain: { ...plain, ms: Math.round(t2 - t1) }, errors: cmp.errors.slice(0, 5) };
  },
  pageRectAt(index: number) {
    const element = document.querySelector(`[data-page-id][data-page-index="${index}"]`);
    return element ? rectOf(element) : null;
  },
  async pageOfText(text: string) {
    ref.current!.scrollToPage(1);
    await sleep(300);
    const found = await ref.current!.search(text);
    await sleep(1200);
    return { total: found.total, page: ref.current!.getCurrentPage() };
  },
  async gotoText(text: string) {
    const found = await ref.current!.search(text);
    await sleep(1500);
    const page = ref.current!.getCurrentPage();
    ref.current!.scrollToPage(page);
    await sleep(1500);
    return { total: found.total, page };
  },
  focus: () => ref.current?.focus(),
  editor: () => ref.current,
} as any, ready, layoutDone, { bytes: bytes.byteLength });
const onError = (error: Error, info?: unknown) => {
  const list = ((cmp.info.onErrors as unknown[]) ??= []);
  if (list.length < 10) list.push({ message: String(error?.message ?? error).slice(0, 400), name: error?.name, info, stack: String(error?.stack ?? '').split('\n').slice(0, 6).join(' | ').slice(0, 600) });
};
createRoot(document.getElementById('root')!).render(<DocxEditor ref={ref} documentBuffer={bytes.buffer} onError={onError} />);
mounted();
