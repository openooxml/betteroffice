import { LEGACY_SYNC_API_PROP } from './peerless/contracts';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import JSZip from 'jszip';
import type { DisplayList } from '@betteroffice/docx/layout/render';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import { takePreloadedResidentEngineWorker } from '@betteroffice/docx/yrs';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import type {
  DocxEngineWorkerRequest,
  DocxEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { pagedDocx } from './__fixtures__/pagedDocx';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './internals/engineChoice';
import { setupWorkerEngine } from './__fixtures__/workerEngine';
import { onPresented, presentedWorkerFrame } from './internals/layoutProvenance';

const { act, cleanup, render, waitFor } = await import('@testing-library/react');

// With a preview, a load's second session is its full one.
const real = await import('@betteroffice/docx/yrs');
const { createYrsSession } = real;
let created = 0;
let holdFullOpen: Promise<void> | null = null;
mock.module('@betteroffice/docx/yrs', () => ({
  ...real,
  createYrsSession: async (options: Parameters<typeof createYrsSession>[0]) => {
    created += 1;
    if (created === 2 && holdFullOpen) await holdFullOpen;
    return createYrsSession(options);
  },
}));
const canvasReplay = await import('./canvasReplay');
const { presentCanvasReplay } = canvasReplay;
let holdReplays: Promise<void> | null = null;
mock.module('./canvasReplay', () => ({
  ...canvasReplay,
  presentCanvasReplay: async (...args: Parameters<typeof presentCanvasReplay>) => {
    if (holdReplays) await holdReplays;
    return presentCanvasReplay(...args);
  },
}));
const { DocxEditor } = await import('../../index');
type Editor = import('../../index').DocxEditorRef;

const ROOT = resolve(import.meta.dir, '../../../../..');
const WASM = resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const fixture = (path: string) => {
  const bytes = readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures', path));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};
const quiet = { error: console.error, warn: console.warn };
const originalWorker = globalThis.Worker;
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>> | null = null;

async function installWorker() {
  const startWorker = await residentWorkerFactory();
  const workers: InProcessResidentWorker[] = [];
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  setMissingWorkerCapabilitiesForTests([]);
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
  return { workers };
}

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: {
        addEventListener: () => {},
        removeEventListener: () => {},
        ready: Promise.resolve(),
      },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  console.error = () => {};
  console.warn = () => {};
});
beforeEach(() => {
  globalThis.Worker = undefined as unknown as typeof Worker;
});
afterEach(async () => {
  cleanup();
  takePreloadedResidentEngineWorker()?.destroy();
  await act(async () => {});
  compileModule?.mockRestore();
  compileModule = null;
  resetEngineChoiceForTests();
  globalThis.Worker = originalWorker;
  created = 0;
  holdFullOpen = null;
  holdReplays = null;
});
afterAll(async () => {
  await new Promise((done) => setTimeout(done, 100));
  console.error = quiet.error;
  console.warn = quiet.warn;
  globalThis.Worker = originalWorker;
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a preview reports its first page painted before the full document opens', async () => {
  let releaseFullOpen = () => {};
  holdFullOpen = new Promise((done) => (releaseFullOpen = done));
  const ref = createRef<Editor>();
  let painted = 0;
  render(
    <DocxEditor
      ref={ref}
      previewFirstPage
      experimentalWorkerOpen={false}
      documentBuffer={fixture('page-fragments/pages.docx')}
      onFirstPagePainted={() => (painted += 1)}
    />
  );
  await waitFor(() => expect(created).toBe(2), { timeout: 20_000 });
  expect(painted).toBe(1);

  await act(async () => releaseFullOpen());
  await waitFor(() => expect(ref.current!.getDocument()).not.toBeNull(), {
    timeout: 20_000,
  });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 200));
  });
  expect(painted).toBe(1);
}, 40_000);

test('a document reports its first page painted once its canvas presents, not at layout', async () => {
  let releaseReplays = () => {};
  holdReplays = new Promise((done) => (releaseReplays = done));
  let painted = 0;
  const view = render(
    <DocxEditor
      experimentalWorkerOpen={false}
      documentBuffer={fixture('page-fragments/pages.docx')}
      onFirstPagePainted={() => (painted += 1)}
    />
  );
  await waitFor(() => expect(view.container.querySelector('.canvas-page')).not.toBeNull(), {
    timeout: 20_000,
  });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 200));
  });
  expect(painted).toBe(0);

  await act(async () => releaseReplays());
  await waitFor(() => expect(painted).toBe(1), { timeout: 20_000 });
}, 40_000);

test('a legacy editor measures and paints an app URL family through the no-provider path', async () => {
  const zip = await JSZip.loadAsync(await pagedDocx(1));
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:r>',
    '<w:r><w:rPr><w:rFonts w:ascii="Brand Sans" w:hAnsi="Brand Sans"/></w:rPr>'));
  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  const noProvider = {
    resolve: () => undefined,
    resolveFamily: () => undefined,
    resolveScriptFallback: () => undefined,
    resolveLastResort: () => undefined,
  };
  const open = async (measurementFontProvider?: typeof noProvider) => {
    const ref = createRef<Editor>();
    let painted = false;
    let display: DisplayList | undefined;
    const unsubscribe = onPresented((list) => { display = list as DisplayList; });
    const view = render(<DocxEditor ref={ref} experimentalWorkerOpen={false}
      documentBuffer={buffer} fonts={[{ family: 'Brand Sans', src: '/brand.woff2' }]}
      measurementFontProvider={measurementFontProvider} onFirstPagePainted={() => { painted = true; }} />);
    try {
      await waitFor(() => expect(ref.current).not.toBeNull(), { timeout: 20_000 });
      expect(await ref.current!.whenLayoutComplete({ timeoutMs: 20_000 })).toBe(1);
      await waitFor(() => expect(painted).toBe(true), { timeout: 20_000 });
      expect(document.head.textContent).toContain('/brand.woff2');
      return display!.pages.flatMap((page) => page.primitives).filter(
        (primitive) => primitive.kind === 'text' || primitive.kind === 'glyphRun',
      );
    } finally { unsubscribe(); view.unmount(); }
  };
  const browser = await open(noProvider);
  expect(browser.some((primitive) => primitive.kind === 'text' && primitive.font.includes('Brand Sans'))).toBe(true);
  expect(await open()).toEqual(browser);
}, 60_000);

test('a legacy editor leaves hosted Calibri to the browser while unhosted Cambria uses the manager', async () => {
  const zip = await JSZip.loadAsync(await pagedDocx(1));
  const xml = await zip.file('word/document.xml')!.async('string');
  const paragraph = (family: string, text: string) =>
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  zip.file('word/document.xml', xml.replace(/<w:p>.*?<\/w:p>/,
    paragraph('Calibri', 'Hosted Calibri') + paragraph('Cambria', 'Unhosted Cambria')));
  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  const noProvider = { resolve: () => undefined };
  const open = async (measurementFontProvider?: typeof noProvider) => {
    const ref = createRef<Editor>();
    let painted = false;
    let display: DisplayList | undefined;
    const unsubscribe = onPresented((list) => { display = list as DisplayList; });
    const view = render(<DocxEditor ref={ref} experimentalWorkerOpen={false}
      documentBuffer={buffer} fonts={[{ family: 'Calibri', src: '/calibri.woff2' }]}
      measurementFontProvider={measurementFontProvider} onFirstPagePainted={() => { painted = true; }} />);
    try {
      await waitFor(() => expect(ref.current).not.toBeNull(), { timeout: 20_000 });
      expect(await ref.current!.whenLayoutComplete({ timeoutMs: 20_000 })).toBe(1);
      await waitFor(() => expect(painted).toBe(true), { timeout: 20_000 });
      expect(document.head.textContent).toContain('/calibri.woff2');
      return display!.pages.flatMap((page) => page.primitives).filter(
        (primitive) => primitive.kind === 'text' || primitive.kind === 'glyphRun',
      );
    } finally { unsubscribe(); view.unmount(); }
  };
  const browser = await open(noProvider);
  const implicit = await open();
  const hosted = (primitives: typeof browser) => primitives.filter(
    (primitive) => primitive.kind === 'text' && primitive.font.includes('Calibri'),
  );
  expect(hosted(browser)).not.toHaveLength(0);
  expect(hosted(implicit)).toEqual(hosted(browser));
  expect(implicit.some((primitive) => primitive.kind === 'glyphRun' && primitive.text === 'Hosted Calibri')).toBe(false);
  expect(implicit.some((primitive) => primitive.kind === 'glyphRun' && primitive.text === 'Unhosted Cambria')).toBe(true);
}, 60_000);

test('each document reports its first page painted once, however often its pages repaint', async () => {
  const ref = createRef<Editor>();
  const painted: string[] = [];
  const editor = (name: string, buffer: ArrayBuffer) => (
    <DocxEditor ref={ref} experimentalWorkerOpen={false} documentBuffer={buffer} onFirstPagePainted={() => painted.push(name)} />
  );
  const view = render(editor('pages', fixture('page-fragments/pages.docx')));
  await waitFor(() => expect(painted).toEqual(['pages']), { timeout: 20_000 });

  await act(async () => ref.current!.setZoom(1.5));
  await act(async () => {
    await new Promise((done) => setTimeout(done, 300));
  });
  expect(painted).toEqual(['pages']);

  view.rerender(editor('principal', fixture('structured-export/principal.docx')));
  await waitFor(() => expect(painted).toEqual(['pages', 'principal']), { timeout: 20_000 });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 300));
  });
  expect(painted).toEqual(['pages', 'principal']);
}, 60_000);

test('each worker document reports its first page painted once, however often its pages repaint', async () => {
  const { workers } = await installWorker();
  const ref = createRef<Editor>();
  const painted: string[] = [];
  const editor = (name: string, buffer: ArrayBuffer) => (
    <DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }} ref={ref} documentBuffer={buffer} onFirstPagePainted={() => painted.push(name)} />
  );
  const view = render(editor('pages', await pagedDocx(2)));
  await waitFor(() => expect(painted).toEqual(['pages']), { timeout: 20_000 });
  expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);

  await act(async () => ref.current!.setZoom(1.5));
  await act(async () => {
    await new Promise((done) => setTimeout(done, 300));
  });
  expect(painted).toEqual(['pages']);

  view.rerender(editor('principal', await pagedDocx(1)));
  await waitFor(() => expect(painted).toEqual(['pages', 'principal']), { timeout: 20_000 });
  expect(workers.reduce((count, worker) => count + worker.requests.filter((request) => request === 'open').length, 0)).toBe(2);
  await act(async () => {
    await new Promise((done) => setTimeout(done, 300));
  });
  expect(painted).toEqual(['pages', 'principal']);
}, 60_000);

describe('DocxEditor first page painted (worker engine)', () => {
  let attachGate: Promise<void>;
  let attachRequested = false;
  const workers = setupWorkerEngine((worker) => {
    const post = worker.postMessage.bind(worker);
    worker.postMessage = (request, transfer) => {
      if (request.type === 'attachCanvases') {
        attachRequested = true;
        void attachGate.then(() => post(request, transfer));
        return;
      }
      post(request, transfer);
    };
  });

  test('a document on the worker engine reports its first page painted once its canvas presents, not at layout', async () => {
    let releaseAttach = () => {};
    attachGate = new Promise<void>((done) => (releaseAttach = done));
    const ref = createRef<Editor>();
    let painted = 0;
    try {
      const view = render(
        <DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }}
          ref={ref}
          experimentalWorkerOpen
          documentBuffer={await pagedDocx(2)}
          onFirstPagePainted={() => (painted += 1)}
        />
      );
      await waitFor(() => expect(view.container.querySelector('.canvas-page')).not.toBeNull(), {
        timeout: 20_000,
      });
      await waitFor(() => expect(attachRequested).toBe(true), { timeout: 20_000 });
      expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);
      expect(await ref.current!.whenLayoutComplete({ timeoutMs: 20_000 })).toBe(2);
      await act(async () => {
        await new Promise((done) => setTimeout(done, 200));
      });
      expect(painted).toBe(0);

      await act(async () => releaseAttach());
      await waitFor(() => expect(painted).toBe(1), { timeout: 20_000 });
      await act(async () => {
        await new Promise((done) => setTimeout(done, 200));
      });
      expect(painted).toBe(1);
    } finally {
      releaseAttach();
    }
  }, 40_000);
});

describe('DocxEditor first page preview (worker engine)', () => {
  let fullOpenGate: Promise<void>;
  const posted: DocxEngineWorkerRequest[] = [];
  const fullOpens: DocxEngineWorkerRequest[] = [];
  const fullOpenReplies: DocxEngineWorkerResponse[] = [];
  const workers = setupWorkerEngine((worker) => {
    const post = worker.postMessage.bind(worker);
    const fullOpenIds = new Set<number>();
    let listener: typeof worker.onmessage = null;
    Object.defineProperty(worker, 'onmessage', {
      get: () => (event: MessageEvent<DocxEngineWorkerResponse>) => {
        if (fullOpenIds.has(event.data.id)) {
          fullOpenReplies.push(event.data);
          void fullOpenGate.then(() => listener?.(event));
        } else listener?.(event);
      },
      set: (next: typeof worker.onmessage) => { listener = next; },
    });
    worker.postMessage = (request, transfer) => {
      if (request.type !== 'editModule') posted.push(request);
      if (request.type === 'open' && request.previewBlocks === undefined) {
        fullOpens.push(request);
        fullOpenIds.add(request.id);
      }
      post(request, transfer);
    };
  });

  test('a preview on the worker engine reports its first page painted before the full document opens', async () => {
    let releaseFullOpen = () => {};
    fullOpenGate = new Promise<void>((done) => (releaseFullOpen = done));
    const ref = createRef<Editor>();
    let painted = 0;
    let fullPresented = false;
    const offPresented = onPresented((displayList, options) => {
      const frame = presentedWorkerFrame({ displayList });
      if (options?.worker && frame?.preview === false) fullPresented = true;
    });
    try {
      render(
        <DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }}
          ref={ref}
          previewFirstPage
          experimentalWorkerOpen
          documentBuffer={await pagedDocx(2)}
          onFirstPagePainted={() => (painted += 1)}
        />
      );
      await waitFor(() => expect(fullOpens).toHaveLength(1), { timeout: 20_000 });
      const previewOpen = posted.findIndex((request) => request.type === 'open' && request.previewBlocks !== undefined);
      const previewBootstrap = posted.findIndex((request) => request.type === 'bootstrap' && request.opened === true);
      expect(previewOpen).toBeGreaterThanOrEqual(0);
      expect(previewBootstrap).toBeGreaterThan(previewOpen);
      expect(posted.indexOf(fullOpens[0]!)).toBeGreaterThan(previewBootstrap);
      await waitFor(() => expect(fullOpenReplies).toHaveLength(1), { timeout: 20_000 });
      expect(fullOpenReplies[0]!.ok).toBe(true);
      await waitFor(() => expect(painted).toBe(1), { timeout: 20_000 });
      expect(ref.current!.getDocument()).toBeNull();
      expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);

      await act(async () => releaseFullOpen());
      await waitFor(() => expect(ref.current!.getDocument()).not.toBeNull(), { timeout: 20_000 });
      await waitFor(() => expect(fullPresented).toBe(true), { timeout: 20_000 });
      await act(async () => {
        await new Promise<void>((done) => requestAnimationFrame(() => requestAnimationFrame(() => done())));
      });
      expect(painted).toBe(1);
    } finally {
      offPresented();
      releaseFullOpen();
    }
  }, 40_000);
});
