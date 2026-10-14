import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import { CANVAS_PAGE_GAP_PX } from '@betteroffice/docx/layout/render';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { DocxCommandFrame, ResidentCommandConnection } from '@betteroffice/docx/peerless';
import type { UsePeerlessEditSessionOptions } from '../hooks/usePeerlessEditSession';
import type { PeerlessFrameGeometry } from './viewStore';
import { PeerlessDocxEditor } from './PeerlessDocxEditor';
import * as paged from './PeerlessPagedEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const identity = { session: 'scroll-container', incarnation: 1, epoch: 1 };
const stride = 100 + CANVAS_PAGE_GAP_PX;

function frame(): DocxCommandFrame {
  const bytes = new ArrayBuffer(88);
  const data = new DataView(bytes);
  new Uint8Array(bytes).set([0x46, 0x44, 0x56, 0x31]);
  data.setUint16(4, 1, true);
  data.setUint16(6, 80, true);
  data.setUint32(8, 88, true);
  data.setUint32(12, 1, true);
  for (const offset of [16, 24, 32]) data.setBigUint64(offset, 1n, true);
  for (const offset of [56, 60]) data.setUint32(offset, 80, true);
  data.setUint32(64, 4, true);
  data.setUint32(68, 88, true);
  return { revision: 1, coversThroughSeq: 0, payload: {
    frame: bytes, layoutRevision: 1, caret: { frameEpoch: 1, caretRect: null },
    stagedPresentation: { frameEpoch: 1, pages: [0, 1, 2].map((pageIndex) => ({
      page: { pageId: String(pageIndex + 1), pageIndex }, frameEpoch: 1,
      bitmap: { width: 200, height: 100, close() {} } as ImageBitmap,
    })) },
  } };
}

async function mount() {
  let scroller: HTMLElement | undefined;
  const rect = spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
    const top = this === scroller ? 0 : -(scroller?.scrollTop ?? 0);
    return { x: 0, y: top, top, bottom: top + 60, left: 0, right: 200, width: 200, height: 60, toJSON() {} };
  });
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation((() => ({
    save() {}, restore() {}, setTransform() {}, clearRect() {}, drawImage() {}, putImageData() {},
    getImageData: () => ({ width: 200, height: 100, data: new Uint8ClampedArray(200 * 100 * 4) }),
  }) as unknown as CanvasRenderingContext2D) as unknown as typeof HTMLCanvasElement.prototype.getContext);
  restorers.push(() => rect.mockRestore(), () => context.mockRestore());
  const requests: { start: number; end: number; frameEpoch: number }[] = [];
  const initial: ResidentCommandConnection['initial'] = { ...identity, version: 0, throughSeq: 0,
    outcomes: [], selection: null,
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: null },
    frame: frame(),
  };
  const connection: ResidentCommandConnection = {
    initial, requestFullFrame: async () => initial,
    async requestViewport(request) { requests.push(request); },
    client: { failure: undefined, onFailure: () => () => {}, dispose: async () => {},
      onFrame: () => () => {},
      call: { checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }),
        edit: async () => initial },
    },
  };
  const options: UsePeerlessEditSessionOptions = { source: Uint8Array.of(1), openConnection: async () => connection,
    canvasForPage: () => null,
    prepareGeometry(candidate) {
      return { identity, coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision, layoutRevision: 1,
        pages: Array.from({ length: 20 }, (_, pageIndex): PeerlessFrameGeometry['pages'][number] => ({
          page: { pageId: String(pageIndex + 1), pageIndex },
          bounds: { x: 0, y: pageIndex * stride, width: 200, height: 100 }, transform: [1, 0, 0, 1, 0, 0],
        })), caret: null, selectionRects: [],
      };
    },
  };
  const PagedEditor = paged.PeerlessPagedEditor;
  const surface = spyOn(paged, 'PeerlessPagedEditor').mockImplementation((props) =>
    <PagedEditor {...props} sessionOptions={options} />);
  restorers.push(() => surface.mockRestore());
  const view = render(<PeerlessDocxEditor documentBuffer={Uint8Array.of(1)} className="host-editor"
    style={{ height: 60, width: 200 }} />);
  scroller = view.container.querySelector<HTMLElement>('.docx-editor__scroll-container') ?? (view.container.firstElementChild as HTMLElement);
  Object.defineProperties(scroller, { clientHeight: { value: 60 }, scrollHeight: { value: stride * 20 } });
  await waitFor(() => expect(view.container.querySelector('[data-page-id="1"] canvas')).not.toBeNull());
  await waitFor(() => expect(requests.length).toBeGreaterThan(0));
  const canvas = view.container.querySelector<HTMLCanvasElement>('[data-page-id="1"] canvas')!;
  return { ...view, scroller, column: canvas.parentElement!.parentElement!.parentElement!, requests };
}

test('peerless_scroll_container_exposes_computed_overflow_before_document_open', () => {
  const view = render(<PeerlessDocxEditor />);
  const scroller = view.container.querySelector<HTMLElement>('.docx-editor__scroll-container');
  expect(scroller).not.toBeNull();
  expect(getComputedStyle(scroller!).overflowY).toBe('auto');
  expect(getComputedStyle(scroller!).overflowX).toBe('auto');
});

test('peerless_editor_owns_scroll_container', async () => {
  const editor = await mount();
  const scrollParent = findVerticalScrollParentOrRoot(editor.column);
  expect(scrollParent).toBe(editor.scroller);
  expect(scrollParent).not.toBe(document.documentElement);
  expect(editor.scroller.classList.contains('docx-editor__scroll-container')).toBe(true);
  expect(editor.getByTestId('docx-editor').classList.contains('host-editor')).toBe(true);
  expect(editor.container.querySelectorAll('.docx-editor__scroll-container')).toHaveLength(1);
  expect(getComputedStyle(editor.scroller).overflowY).toBe('auto');
  expect(getComputedStyle(editor.scroller).overflowX).toBe('auto');
  expect(editor.scroller.tabIndex).toBe(0);
});

test('page_keys_scroll_peerless_container_without_blurring_document_input', async () => {
  const editor = await mount();
  const input = editor.getByRole('textbox', { name: 'Document input' });
  act(() => { input.focus(); });
  expect(document.activeElement).toBe(input);
  const before = editor.scroller.scrollTop;
  const down = new KeyboardEvent('keydown', { key: 'PageDown', bubbles: true, cancelable: true });
  act(() => { fireEvent(input, down); });
  const distance = editor.scroller.scrollTop - before;
  expect(distance).toBeGreaterThanOrEqual(editor.scroller.clientHeight * 0.8);
  expect(distance).toBeLessThanOrEqual(editor.scroller.clientHeight);
  expect(down.defaultPrevented).toBe(true);
  expect(document.activeElement).toBe(input);
  const up = new KeyboardEvent('keydown', { key: 'PageUp', bubbles: true, cancelable: true });
  act(() => { fireEvent(input, up); });
  expect(editor.scroller.scrollTop).toBe(before);
  expect(up.defaultPrevented).toBe(true);
  expect(document.activeElement).toBe(input);
});

test('scrolling_peerless_container_updates_display_window', async () => {
  const editor = await mount();
  const before = editor.requests.length;
  act(() => { editor.scroller.scrollTop = 15 * stride + 1; fireEvent.scroll(editor.scroller); });
  await waitFor(() => expect(editor.requests.slice(before).some(({ start, end }) => start <= 15 && end > 15)).toBe(true));
  expect(editor.requests.at(-1)?.start).toBeGreaterThan(0);
  const after = editor.requests.length;
  act(() => { editor.scroller.scrollTop = 0; fireEvent.scroll(editor.scroller); });
  await waitFor(() => expect(editor.requests.slice(after).some(({ start }) => start === 0)).toBe(true));
});
