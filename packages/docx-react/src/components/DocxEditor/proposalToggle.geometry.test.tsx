import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef, useLayoutEffect } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { takePreloadedResidentEngineWorker, type YrsSession } from '@betteroffice/docx/yrs';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import {
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorRef,
  type DocxPluginGeometry,
  type DocxAnchorGeometryResult,
} from '../../index';
import { awaitWorkerOpenReplica, requestWorkerOpenReplica } from './internals/workerOpenReplica';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restores: Array<() => void> = [];
const workers: InProcessResidentWorker[] = [];
const sessions: YrsSession[] = [];
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
const font = readFileSync(resolve(
  import.meta.dir, '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
));

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});

beforeEach(() => {
  const originalWorker = globalThis.Worker;
  restores.push(() => { globalThis.Worker = originalWorker; });
  const fonts = Object.getOwnPropertyDescriptor(document, 'fonts');
  restores.push(() => {
    if (fonts) Object.defineProperty(document, 'fonts', fonts);
    else Reflect.deleteProperty(document, 'fonts');
  });
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: { addEventListener() {}, removeEventListener() {}, ready: Promise.resolve() },
  });
});

afterEach(() => {
  try {
    cleanup();
  } finally {
    takePreloadedResidentEngineWorker()?.destroy();
    for (const worker of workers.splice(0)) {
      worker.terminate();
      for (const session of worker.sessions) session.destroy();
    }
    for (const session of sessions.splice(0)) session.destroy();
    for (const restore of restores.splice(0).reverse()) restore();
  }
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function installWorker() {
  const compile = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  restores.push(() => compile.mockRestore());
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
}

function documentBuffer(): ArrayBuffer {
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>'
  ));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>'
  ));
  const body = Array.from({ length: 20 }, (_, index) =>
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}">` +
    `<w:pPr>${index ? '<w:pageBreakBefore/>' : ''}</w:pPr>` +
    `<w:r><w:t>Target ${index}</w:t></w:r></w:p>`
  ).join('');
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    '<w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>' +
    '</w:sectPr></w:body></w:document>'
  ));
  return rezipPartsToArrayBuffer(parts);
}

type Anchor = Extract<DocxAnchorGeometryResult, { ok: true }>;
type Sample = {
  layoutId: string;
  previewVersion: number;
  anchor: Anchor['anchor'];
  rects: Anchor['rects'];
  pageRect: Anchor['pageRect'];
  clientY: number;
};

test.each([0, 19])('Accept, Undo, Reject, Undo show only settled proposal geometry with a distant caret on page %s', async (caretPage) => {
  installWorker();
  const ref = createRef<DocxEditorRef>();
  let geometry: DocxPluginGeometry | null = null;
  let recording = false;
  let scroller: HTMLElement | null = null;
  let scrollTop = 0;
  const samples: Sample[] = [];
  const errors: unknown[] = [];
  const sample = () => {
    if (!recording || !geometry) return;
    const result = geometry.getAnchorGeometry({ kind: 'proposal', id: 'toggle-0' });
    if (!result.ok) return;
    const layer = document.querySelector<HTMLElement>('.docx-plugin-overlays');
    if (!layer) return;
    samples.push({
      layoutId: result.layoutId, previewVersion: result.previewVersion,
      anchor: { ...result.anchor }, clientY: layer.getBoundingClientRect().top + result.anchor.y,
      rects: result.rects.map((rect) => ({ ...rect })), pageRect: { ...result.pageRect },
    });
  };
  const bounds = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    if (this === scroller) return new DOMRect(0, 0, 816, 400);
    const page = this.dataset.pageIndex === undefined ? null : this.closest<HTMLElement>('.canvas-page');
    if (page) {
      const height = Number.parseFloat(page.style.height) || 1056;
      const width = Number.parseFloat(page.style.width) || 816;
      return new DOMRect(0, 24 + Number(page.dataset.pageIndex) * (height + 16) - scrollTop, width, height);
    }
    return new DOMRect(0, -scrollTop, 816, 20 * 1072 + 48);
  });
  restores.push(() => bounds.mockRestore());
  const plugin = defineDocxPlugin({
    id: 'test.proposal-toggle-geometry',
    createState: () => null,
    onEvent(context, event) {
      if (event.type === 'layout-change' && context.geometry) {
        geometry = context.geometry;
        sample();
      }
    },
    overlay: function Overlay(props) {
      geometry = props.geometry;
      useLayoutEffect(sample);
      const anchor = props.geometry.getAnchorGeometry({ kind: 'proposal', id: 'toggle-0' });
      return anchor.ok ? <div data-testid="toggle-action-bar" style={{
        position: 'absolute', left: anchor.anchor.x, top: anchor.anchor.y,
      }} /> : null;
    },
  });
  const host = document.createElement('div');
  document.body.append(host);
  restores.push(() => host.remove());
  const provider = { resolve: () => () => Promise.resolve(font.buffer.slice(
    font.byteOffset, font.byteOffset + font.byteLength
  ) as ArrayBuffer) };
  render(<DocxEditor
    ref={ref} documentBuffer={documentBuffer()} allowHostProposals experimentalWorkerOpen
    measurementFontProvider={provider} plugins={[plugin]} onPluginError={(error) => errors.push(error)}
  />, { container: host });
  const scrollElement = await waitFor(() => {
    const element = host.querySelector<HTMLElement>('.docx-editor__scroll-container');
    expect(element).not.toBeNull();
    return element!;
  }, { timeout: 10_000 });
  scroller = scrollElement;
  scrollElement.style.overflowY = 'auto';
  scrollElement.style.height = '400px';
  Object.defineProperties(scrollElement, {
    clientHeight: { value: 400 }, scrollHeight: { value: 20 * 1072 + 48 },
    scrollTop: {
      configurable: true,
      get: () => scrollTop,
      set: (value: number) => { sample(); scrollTop = value; sample(); },
    },
  });
  await waitFor(() => expect(ref.current?.getEditorRef()?.getYrsSession()).toBeTruthy());
  await act(async () => { await ref.current!.whenLayoutComplete(); });
  const worker = workers.find((worker) => worker.requests.includes('open'))!;
  expect(worker).toBeDefined();
  const assertWorkerAvailable = () => expect(ref.current!.getMemoryStats().worker).not.toBeNull();
  assertWorkerAvailable();
  const editor = ref.current!.getEditorRef()!;
  const session = editor.getYrsSession()!;
  sessions.push(session);
  await act(async () => {
    requestWorkerOpenReplica(session);
    await awaitWorkerOpenReplica(session);
  });
  assertWorkerAvailable();
  await act(async () => {
    const caret = session.paragraphs('body')[caretPage]!;
    session.setSelection({ story: 'body', paraId: caret.paraId, offset: 0 });
    editor.syncYrsInputState(false);
  });
  const identities = await ref.current!.getParagraphIdentities();
  const targets = identities.paragraphs.filter(({ session }) => session?.story === 'body').slice(8, 18);
  expect(targets).toHaveLength(10);
  await act(async () => {
    const before = await ref.current!.getProposals();
    expect(await ref.current!.proposeChanges({
      expectVersion: before.version,
      proposals: targets.map((entry, index) => ({
        id: `toggle-${index}`, paragraph: entry.session!,
        suggest: { author: 'Host', date: '2026-10-05T00:00:00Z' },
        op: 'replaceText' as const, search: 'Target', replaceWith: 'Change',
      })),
    })).toMatchObject({ ok: true });
    await ref.current!.whenLayoutComplete();
  });
  await waitFor(() => expect(geometry?.getAnchorGeometry({ kind: 'proposal', id: 'toggle-0' }).ok).toBe(true));
  const frames = async () => {
    await act(async () => {
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });
  };
  await frames();
  scrollTop = 8 * 1072 + 24;
  fireEvent.scroll(scrollElement);
  await frames();
  const selection = session.selection();
  const failures: unknown[] = [];
  for (const state of ['accepted', 'proposed', 'rejected', 'proposed'] as const) {
    assertWorkerAvailable();
    const initialScroll = scrollTop;
    samples.length = 0;
    recording = true;
    sample();
    const before = await ref.current!.getProposals();
    const postedBefore = worker.requests.length;
    await act(async () => {
      expect(await ref.current!.setProposalStates({
        expectVersion: before.version, expectPreviewVersion: before.previewVersion,
        changes: [{ id: 'toggle-0', state }],
      })).toMatchObject({ ok: true });
    });
    await act(async () => { await ref.current!.whenLayoutComplete(); });
    await frames();
    sample();
    recording = false;
    assertWorkerAvailable();
    expect(worker.requests.length).toBeGreaterThan(postedBefore);
    const settled = samples.at(-1)!;
    expect(settled).toBeDefined();
    expect(settled.previewVersion).toBe(before.previewVersion + 1);
    const shown = samples.filter((entry) => entry.previewVersion === settled.previewVersion);
    expect(shown.length).toBeGreaterThan(0);
    const intermediate = shown.filter((entry) =>
      JSON.stringify([entry.anchor, entry.rects, entry.pageRect]) !==
        JSON.stringify([settled.anchor, settled.rects, settled.pageRect]) || entry.clientY !== settled.clientY
    );
    if (intermediate.length || scrollTop !== initialScroll) failures.push({
      state, initialScroll, finalScroll: scrollTop, settled, intermediate,
      layouts: [...new Set(shown.map(({ layoutId }) => layoutId))],
    });
    expect(session.selection()).toEqual(selection);
    scrollTop = 8 * 1072 + 24;
    fireEvent.scroll(scrollElement);
    await frames();
  }
  assertWorkerAvailable();
  expect(worker.requests).toContain('open');
  expect(worker.requests).toContain('sync');
  expect(errors).toEqual([]);
  expect(failures).toEqual([]);
}, 30_000);
