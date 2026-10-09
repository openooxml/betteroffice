import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import { takePreloadedResidentEngineWorker } from '@betteroffice/docx/yrs';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './internals/engineChoice';

const { act, cleanup, render, waitFor } = await import('@testing-library/react');

// Records every call a session gets after it was destroyed.
const real = await import('@betteroffice/docx/yrs');
const { createYrsSession } = real;
const afterDestroy: string[] = [];
const live = new Set<unknown>();
mock.module('@betteroffice/docx/yrs', () => ({
  ...real,
  createYrsSession: async (options: Parameters<typeof createYrsSession>[0]) => {
    const session = await createYrsSession(options);
    let destroyed = false;
    const proxy: typeof session = new Proxy(session, {
      get(target, key, receiver) {
        const value = Reflect.get(target, key, receiver);
        if (typeof value !== 'function') return value;
        return (...args: unknown[]) => {
          if (key === 'destroy') {
            destroyed = true;
            live.delete(proxy);
          } else if (destroyed) {
            afterDestroy.push(String(key));
          }
          return value.apply(target, args);
        };
      },
    });
    live.add(proxy);
    return proxy;
  },
}));
const { DocxEditor } = await import('../../index');
type Editor = import('../../index').DocxEditorRef;

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
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
  return workers;
}

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve() },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  console.error = () => {};
  console.warn = () => {};
});
afterEach(async () => {
  cleanup();
  takePreloadedResidentEngineWorker()?.destroy();
  await act(async () => {});
  compileModule?.mockRestore();
  compileModule = null;
  resetEngineChoiceForTests();
  globalThis.Worker = originalWorker;
});
afterAll(async () => {
  cleanup();
  await new Promise((done) => setTimeout(done, 100));
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** A document of `pages` pages, each with a commented paragraph. */
async function commentedDocx(pages: number): Promise<ArrayBuffer> {
  const body = Array.from(
    { length: pages },
    (_, i) =>
      `<w:p>${i ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : ''}<w:commentRangeStart w:id="${i}"/>` +
      `<w:r><w:t>Page ${i + 1}</w:t></w:r><w:commentRangeEnd w:id="${i}"/>` +
      `<w:r><w:commentReference w:id="${i}"/></w:r></w:p>`
  ).join('');
  const comments = Array.from(
    { length: pages },
    (_, i) =>
      `<w:comment w:id="${i}" w:author="A" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Note ${i}</w:t></w:r></w:p></w:comment>`
  ).join('');
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
      '</Relationships>'
  );
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  zip.file('word/comments.xml', `<w:comments xmlns:w="${W}">${comments}</w:comments>`);
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function settle(ms: number) {
  for (let waited = 0; waited < ms; waited += 50) {
    await act(async () => {
      await new Promise((done) => setTimeout(done, 50));
    });
  }
}

test('loading another document never calls into the session it replaced', async () => {
  const view = render(<DocxEditor experimentalWorkerOpen={false} documentBuffer={await commentedDocx(3)} />);
  await settle(2500);
  view.rerender(<DocxEditor experimentalWorkerOpen={false} documentBuffer={await commentedDocx(4)} />);
  await settle(3000);
  expect(afterDestroy).toEqual([]);
  view.unmount();
}, 30_000);

test('a load that fails frees the document it replaced', async () => {
  afterDestroy.length = 0;
  const view = render(<DocxEditor experimentalWorkerOpen={false} documentBuffer={await commentedDocx(2)} />);
  await settle(2500);
  const detached = await commentedDocx(1);
  structuredClone(detached, { transfer: [detached] });
  view.rerender(<DocxEditor experimentalWorkerOpen={false} documentBuffer={detached} />);
  await settle(1500);
  expect(live.size).toBe(0);
  expect(afterDestroy).toEqual([]);
  view.unmount();
}, 30_000);

test('loading another document frees the worker document it replaced', async () => {
  afterDestroy.length = 0;
  const workers = await installWorker();
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const editor = (buffer: ArrayBuffer) => (
    <DocxEditor ref={ref} documentBuffer={buffer} onError={(error) => errors.push(error)} />
  );
  const view = render(editor(await commentedDocx(3)));
  await waitFor(() => expect(ref.current?.getTotalPages()).toBe(3), { timeout: 20_000 });
  await waitFor(() => expect(ref.current!.getComments()).toHaveLength(3), { timeout: 20_000 });
  const previous = workers.filter((worker) => worker.requests.includes('open') && worker.sessions.length > 0);
  expect(previous.length).toBeGreaterThan(0);
  const terminated = previous.map((worker) => spyOn(worker, 'terminate'));
  const destroyed = previous.flatMap((worker) => worker.sessions.map((session) => spyOn(session, 'destroy')));
  try {
    view.rerender(editor(await commentedDocx(4)));
    await waitFor(() => expect(ref.current!.getTotalPages()).toBe(4), { timeout: 20_000 });
    await waitFor(() => expect(ref.current!.getComments()).toHaveLength(4), { timeout: 20_000 });
    expect(await ref.current!.whenLayoutComplete({ timeoutMs: 20_000 })).toBe(4);
    for (const release of terminated) expect(release).toHaveBeenCalledTimes(1);
    for (const release of destroyed) expect(release).toHaveBeenCalledTimes(1);
    expect(workers.some((worker) => !previous.includes(worker) && worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);
    expect(errors).toEqual([]);
    expect(afterDestroy).toEqual([]);
    view.unmount();
  } finally {
    for (const release of [...terminated, ...destroyed]) release.mockRestore();
  }
}, 60_000);

test('a load that fails frees the worker document it replaced', async () => {
  afterDestroy.length = 0;
  const workers = await installWorker();
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const editor = (buffer: ArrayBuffer) => (
    <DocxEditor ref={ref} documentBuffer={buffer} onError={(error) => errors.push(error)} />
  );
  const view = render(editor(await commentedDocx(2)));
  await waitFor(() => expect(ref.current?.getTotalPages()).toBe(2), { timeout: 20_000 });
  const previous = workers.filter((worker) => worker.requests.includes('open') && worker.sessions.length > 0);
  expect(previous.length).toBeGreaterThan(0);
  const terminated = previous.map((worker) => spyOn(worker, 'terminate'));
  const destroyed = previous.flatMap((worker) => worker.sessions.map((session) => spyOn(session, 'destroy')));
  try {
    const detached = await commentedDocx(1);
    structuredClone(detached, { transfer: [detached] });
    view.rerender(editor(detached));
    await waitFor(() => expect(errors).toHaveLength(1), { timeout: 20_000 });
    await waitFor(() => expect(live.size).toBe(0), { timeout: 20_000 });
    for (const release of terminated) expect(release).toHaveBeenCalledTimes(1);
    for (const release of destroyed) expect(release).toHaveBeenCalledTimes(1);
    expect(view.container.querySelector('.docx-editor-error')).not.toBeNull();
    expect(afterDestroy).toEqual([]);
    view.unmount();
  } finally {
    for (const release of [...terminated, ...destroyed]) release.mockRestore();
  }
}, 60_000);
