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
import { hasEditorWorkerProposalRounds, registeredWorkerProposalAuthority } from './internals/workerProposalAuthority';
import { resolveAnchorTarget } from '../../plugins/anchorGeometry';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './internals/engineChoice';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await step(
  'load testing library', () => import('@testing-library/react')
);
const restores: Array<() => void> = [];
const workers: InProcessResidentWorker[] = [];
const sessions: YrsSession[] = [];
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
const font = readFileSync(resolve(
  import.meta.dir, '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
));

async function step<T>(label: string, run: () => T | PromiseLike<T>, flushReact = false): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  let expired = false;
  let done = false;
  const pending = Promise.resolve().then(run);
  void pending.then(() => { done = true; }, () => { done = true; });
  const flush = async () => {
    while (!done && !expired) {
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 10));
      });
    }
    return pending;
  };
  try {
    return await Promise.race([
      flushReact ? flush() : pending,
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => {
          expired = true;
          reject(new Error(`stalled at ${label}`));
        }, 5000);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

beforeAll(async () => {
  await step('preload edit wasm', () => preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  )))));
  startWorker = await step('create worker factory', () => residentWorkerFactory());
});

beforeEach(() => {
  setMissingWorkerCapabilitiesForTests([]);
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
    resetEngineChoiceForTests();
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
  if (ownsDom) await step('unregister DOM', () => GlobalRegistrator.unregister());
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

function documentBuffer(firstParagraphText = 'Target 0'): ArrayBuffer {
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
    `<w:r><w:t>${index ? `Target ${index}` : firstParagraphText}</w:t></w:r></w:p>`
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

test('search inside a rejected peer-local insertion anchors after Hello when worker rounds activate', async () => {
  installWorker();
  const ref = createRef<DocxEditorRef>();
  let geometry: DocxPluginGeometry | null = null;
  const errors: unknown[] = [];
  const bounds = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    const page = this.closest<HTMLElement>('.canvas-page');
    if (page) {
      const height = Number.parseFloat(page.style.height) || 1056;
      const width = Number.parseFloat(page.style.width) || 816;
      return new DOMRect(0, 24 + Number(page.dataset.pageIndex) * (height + 16), width, height);
    }
    return new DOMRect(0, 0, 816, 20 * 1072 + 48);
  });
  restores.push(() => bounds.mockRestore());
  const plugin = defineDocxPlugin({
    id: 'test.peer-local-hidden-search',
    createState: () => null,
    onEvent(context, event) {
      if (event.type === 'layout-change' && context.geometry) geometry = context.geometry;
    },
    overlay: function Overlay(props) {
      geometry = props.geometry;
      return null;
    },
  });
  const host = document.createElement('div');
  document.body.append(host);
  restores.push(() => host.remove());
  const provider = { resolve: () => () => Promise.resolve(font.buffer.slice(
    font.byteOffset, font.byteOffset + font.byteLength
  ) as ArrayBuffer) };
  render(<DocxEditor
    ref={ref} documentBuffer={documentBuffer('Hello world')} allowHostProposals experimentalWorkerOpen
    measurementFontProvider={provider} plugins={[plugin]} onPluginError={(error) => errors.push(error)}
  />, { container: host });
  await step('wait for session', () => waitFor(() => expect(ref.current?.getEditorRef()?.getYrsSession()).toBeTruthy()));
  await step('initial layout', () => ref.current!.whenLayoutComplete(), true);
  const editor = ref.current!.getEditorRef()!;
  const session = editor.getYrsSession()!;
  sessions.push(session);
  await step('hydrate worker replica', () => {
    requestWorkerOpenReplica(session);
    return awaitWorkerOpenReplica(session);
  }, true);
  expect(hasEditorWorkerProposalRounds(session)).toBe(false);
  const paragraph = {
    kind: 'persisted' as const,
    story: { kind: 'body' as const, partUri: '/word/document.xml' }, paraId: '00000001',
  };
  act(() => {
    const inserted = session.proposeChanges({
      expectVersion: session.version(),
      proposals: [{
        id: 'peer-local', paragraph,
        suggest: { author: 'Peer', date: '2026-10-05T00:00:00Z' },
        op: 'insertText', at: { offset: 6 }, text: 'ABCDEFGHIJ',
      }],
    });
    expect(inserted.ok).toBe(true);
    if (!inserted.ok) throw new Error(inserted.failure.message);
    expect(session.setProposalStates({
      expectVersion: inserted.snapshot.version, expectPreviewVersion: inserted.snapshot.previewVersion,
      changes: [{ id: 'peer-local', state: 'rejected' }],
    }).ok).toBe(true);
    editor.syncYrsInputState(true, ['body'], { inWorker: true });
  });
  expect(await step('activate empty worker round', () => ref.current!.proposeChanges({
    expectVersion: session.version(), proposals: [],
  }), true)).toMatchObject({ ok: true, snapshot: { proposals: [] } });
  expect(hasEditorWorkerProposalRounds(session)).toBe(true);
  const local = session.getProposals();
  expect(local.proposals.map(({ id, state }) => ({ id, state }))).toEqual([{ id: 'peer-local', state: 'rejected' }]);
  expect(registeredWorkerProposalAuthority(session)!.revisionPreview()).toEqual({
    [local.proposals[0]!.revisionIds[0]!]: 'rejected',
  });
  const target = { kind: 'search' as const, paragraph, text: 'DEF' };
  expect(resolveAnchorTarget(session, target, session.version())).toEqual({
    ok: true,
    ranges: [{
      start: { story: 'body', paraId: '00000001', offset: 9 },
      end: { story: 'body', paraId: '00000001', offset: 12 },
    }],
    paragraph: { story: 'body', paraId: '00000001', offset: 0 },
  });
  await step('combined preview layout', () => ref.current!.whenLayoutComplete(), true);
  const { geometry: settledGeometry, hidden } = await step('settled search geometry', () => waitFor(() => {
    const current = geometry;
    const result = current?.getAnchorGeometry(target);
    expect(result?.ok).toBe(true);
    if (!current || !result?.ok) throw new Error('No settled search geometry');
    return { geometry: current, hidden: result };
  }));
  const prefix = settledGeometry.getAnchorGeometry({ kind: 'search', paragraph, text: 'Hello ' });
  const start = settledGeometry.getAnchorGeometry({
    kind: 'range', version: session.version(),
    range: { story: 'body', view: 'accepted', start: { paraId: '00000001', offset: 0 }, end: { paraId: '00000001', offset: 0 } },
  });
  expect(prefix.ok).toBe(true);
  expect(start.ok).toBe(true);
  if (!prefix.ok || !start.ok) throw new Error('No settled comparison geometry');
  expect(hidden.rects).toEqual([]);
  expect(hidden.anchor).toEqual(prefix.anchor);
  expect(hidden.anchor).toEqual({ ...prefix.anchor, pageIndex: 0, width: 0 });
  expect(prefix.anchor.x).not.toBe(start.anchor.x);
  expect(hidden.layoutId).toBe(prefix.layoutId);
  expect(errors).toEqual([]);
  expect(ref.current!.getMemoryStats().worker).not.toBeNull();
}, 30_000);

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
  const scrollElement = await step('find scroll container', () => waitFor(() => {
    const element = host.querySelector<HTMLElement>('.docx-editor__scroll-container');
    expect(element).not.toBeNull();
    return element!;
  }, { timeout: 10_000 }));
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
  await step('wait for session', () => waitFor(() => expect(ref.current?.getEditorRef()?.getYrsSession()).toBeTruthy()));
  await step('initial layout', () => ref.current!.whenLayoutComplete(), true);
  const worker = workers.find((worker) => worker.requests.includes('open'))!;
  expect(worker).toBeDefined();
  const assertWorkerAvailable = () => expect(ref.current!.getMemoryStats().worker).not.toBeNull();
  assertWorkerAvailable();
  const editor = ref.current!.getEditorRef()!;
  const session = editor.getYrsSession()!;
  sessions.push(session);
  await step('hydrate worker replica', () => {
    requestWorkerOpenReplica(session);
    return awaitWorkerOpenReplica(session);
  }, true);
  assertWorkerAvailable();
  act(() => {
    const caret = session.paragraphs('body')[caretPage]!;
    session.setSelection({ story: 'body', paraId: caret.paraId, offset: 0 });
    editor.syncYrsInputState(false);
  });
  const identities = await step('get paragraph identities', () => ref.current!.getParagraphIdentities(), true);
  const targets = identities.paragraphs.filter(({ session }) => session?.story === 'body').slice(8, 18);
  expect(targets).toHaveLength(10);
  const before = await step('get proposals before creation', () => ref.current!.getProposals(), true);
  expect(await step('create proposals', () => ref.current!.proposeChanges({
    expectVersion: before.version,
    proposals: targets.map((entry, index) => ({
      id: `toggle-${index}`, paragraph: entry.session!,
      suggest: { author: 'Host', date: '2026-10-05T00:00:00Z' },
      op: 'replaceText' as const, search: 'Target', replaceWith: 'Change',
    })),
  }), true)).toMatchObject({ ok: true });
  await step('created proposal layout', () => ref.current!.whenLayoutComplete(), true);
  await step('wait for proposal geometry', () => waitFor(() => expect(geometry?.getAnchorGeometry({ kind: 'proposal', id: 'toggle-0' }).ok).toBe(true)));
  const frames = async (label: string) => {
    await step(`${label} frame 1`, () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())), true);
    await step(`${label} frame 2`, () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())), true);
  };
  await frames('created proposals');
  scrollTop = 8 * 1072 + 24;
  fireEvent.scroll(scrollElement);
  await frames('initial scroll');
  const selection = session.selection();
  const failures: unknown[] = [];
  for (const state of ['accepted', 'proposed', 'rejected', 'proposed'] as const) {
    assertWorkerAvailable();
    const initialScroll = scrollTop;
    samples.length = 0;
    recording = true;
    sample();
    const before = await step(`${state} get proposals`, () => ref.current!.getProposals(), true);
    const postedBefore = worker.requests.length;
    expect(await step(`${state} set proposal states`, () => ref.current!.setProposalStates({
      expectVersion: before.version, expectPreviewVersion: before.previewVersion,
      changes: [{ id: 'toggle-0', state }],
    }), true)).toMatchObject({ ok: true });
    await step(`${state} layout`, () => ref.current!.whenLayoutComplete(), true);
    await frames(`${state} settled`);
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
    await frames(`${state} scroll`);
  }
  assertWorkerAvailable();
  expect(worker.requests).toContain('open');
  expect(worker.requests).toContain('sync');
  expect(errors).toEqual([]);
  expect(failures).toEqual([]);
}, 30_000);
