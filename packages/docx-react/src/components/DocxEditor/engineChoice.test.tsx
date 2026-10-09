import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { StrictMode } from 'react';
import { parseDocx } from '@betteroffice/docx/docx';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { Document } from '@betteroffice/docx/types/document';
import * as yrs from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import type { DocxEditorProps } from '../DocxEditor';
import * as displayList from './hooks/useDisplayList';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './internals/engineChoice';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
const { DocxEditor } = await import('../DocxEditor');
const originalWorker = globalThis.Worker;
const originalFonts = Object.getOwnPropertyDescriptor(document, 'fonts');
const createYrsSession = yrs.createYrsSession;
const documentToYrs = yrs.documentToYrs;
const useCanvasRenderer = displayList.useCanvasRenderer;
const parts = new Map<string, Uint8Array>([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
  ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Editor content</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')],
]);
const buffer = rezipPartsToArrayBuffer(parts);
let parsed: Document;
let initialUpdate: Uint8Array;
let mainThreadLoads = 0;
const workerModes: boolean[] = [];
const errors: Error[] = [];
let warn: ReturnType<typeof spyOn<typeof console, 'warn'>>;

class FakeWorker {
  static requests: ResidentEngineWorkerRequest[] = [];
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  postMessage(request: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule): void {
    if ('id' in request) FakeWorker.requests.push(request);
  }
  terminate(): void {}
}

beforeAll(async () => {
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: { addEventListener() {}, removeEventListener() {}, ready: Promise.resolve() },
  });
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  parsed = await parseDocx(buffer, { preloadFonts: false });
  const session = await createYrsSession();
  try {
    session.openDocx(new Uint8Array(buffer), true);
    initialUpdate = session.encodeState();
  } finally {
    session.destroy();
  }
});

beforeEach(() => {
  setMissingWorkerCapabilitiesForTests([]);
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  mainThreadLoads = 0;
  workerModes.length = 0;
  errors.length = 0;
  FakeWorker.requests = [];
  warn = spyOn(console, 'warn').mockImplementation(() => {});
  spyOn(displayList, 'useCanvasRenderer').mockImplementation((...args) => {
    workerModes.push(args[5] === true);
    return useCanvasRenderer(...args);
  });
  spyOn(yrs, 'createYrsSession').mockImplementation(async (...args) => {
    const session = await createYrsSession(...args);
    const openDocx = session.openDocx.bind(session);
    session.openDocx = (...openArgs) => {
      mainThreadLoads += 1;
      return openDocx(...openArgs);
    };
    return session;
  });
  spyOn(yrs, 'documentToYrs').mockImplementation((...args) => {
    mainThreadLoads += 1;
    return documentToYrs(...args);
  });
});

afterEach(async () => {
  cleanup();
  yrs.takePreloadedResidentEngineWorker()?.destroy();
  await act(async () => {});
  mock.restore();
  resetEngineChoiceForTests();
  globalThis.Worker = originalWorker;
});

afterAll(async () => {
  if (originalFonts) Object.defineProperty(document, 'fonts', originalFonts);
  else Reflect.deleteProperty(document, 'fonts');
  if (ownsDom) await GlobalRegistrator.unregister();
});

const openRequests = () => FakeWorker.requests.filter((request) => request.type === 'open');
const warnings = () => warn.mock.calls
  .map(([message]) => String(message))
  .filter((message) => message.startsWith('[DocxEditor]'));
const editor = (props: DocxEditorProps = {}) => (
  <DocxEditor showToolbar={false} documentBuffer={buffer} onError={(error) => errors.push(error)} {...props} />
);

async function expectInThread(): Promise<void> {
  await waitFor(() => expect(mainThreadLoads).toBeGreaterThan(0));
  expect(workerModes.every((worker) => !worker)).toBe(true);
  expect(openRequests()).toHaveLength(0);
  expect(errors).toEqual([]);
}

test('the default opens in the worker and keeps its engine when the prop changes', async () => {
  const view = render(editor());
  await waitFor(() => expect(openRequests()).toHaveLength(1));
  view.rerender(editor({ experimentalWorkerOpen: false }));
  await act(async () => {});
  expect(workerModes.every(Boolean)).toBe(true);
  expect(mainThreadLoads).toBe(0);
  expect(warnings()).toEqual([]);
  expect(errors).toEqual([]);
});

const lateOpeningOptions: Array<[string, () => DocxEditorProps]> = [
  ['mediaTokens', () => ({ mediaTokens: true })],
  ['collaboration.initialUpdate', () => ({ collaboration: { initialUpdate } })],
];
test.each(lateOpeningOptions)('late %s keeps worker opens and warns once until remount', async (prop, props) => {
  const view = render(editor());
  await waitFor(() => expect(openRequests()).toHaveLength(1));
  view.rerender(editor(props()));
  await act(async () => {});
  expect(openRequests()).toHaveLength(1);
  for (let load = 2; load <= 3; load += 1) {
    view.rerender(editor({ ...props(), documentBuffer: buffer.slice(0) }));
    await waitFor(() => expect(openRequests()).toHaveLength(load));
  }
  expect(workerModes.every(Boolean)).toBe(true);
  expect(mainThreadLoads).toBe(0);
  expect(warnings()).toHaveLength(1);
  expect(warnings()[0]).toContain(prop);
  expect(warnings()[0]).toContain('take effect only on remount');
  expect(errors).toEqual([]);
});

test('bytes take precedence over an editable parsed document', async () => {
  render(editor({ document: parsed }));
  await waitFor(() => expect(openRequests()).toHaveLength(1));
  expect(mainThreadLoads).toBe(0);
  expect(warnings()).toEqual([]);
  expect(errors).toEqual([]);
});

test('a later parsed source keeps the mounted worker engine', async () => {
  const view = render(editor());
  await waitFor(() => expect(openRequests()).toHaveLength(1));
  view.rerender(editor({ document: parsed, documentBuffer: undefined }));
  await waitFor(() => expect(openRequests()).toHaveLength(2));
  expect(workerModes.every(Boolean)).toBe(true);
  expect(mainThreadLoads).toBe(0);
  expect(warnings()).toEqual([]);
  expect(errors).toEqual([]);
});

test('false selects in-thread and warns once across StrictMode and multiple editors', async () => {
  const view = render(<StrictMode>{editor({ experimentalWorkerOpen: false })}</StrictMode>);
  render(editor({ experimentalWorkerOpen: false }));
  await expectInThread();
  view.rerender(<StrictMode>{editor({ experimentalWorkerOpen: true })}</StrictMode>);
  expect(workerModes.every((worker) => !worker)).toBe(true);
  expect(warnings()).toHaveLength(1);
  expect(warnings()[0]).toContain('experimentalWorkerOpen={false}');
  expect(warnings()[0]).toContain('deprecated');
});

test('missing capabilities fall back even with an explicit worker request', async () => {
  setMissingWorkerCapabilitiesForTests(['OffscreenCanvas', 'createImageBitmap']);
  render(editor({ experimentalWorkerOpen: true }));
  await expectInThread();
  expect(warnings()).toHaveLength(1);
  expect(warnings()[0]).toContain('OffscreenCanvas');
  expect(warnings()[0]).toContain('createImageBitmap');
});

const unsupportedSources: Array<[string, () => DocxEditorProps]> = [
  ['mediaTokens', () => ({ mediaTokens: true })],
  ['collaboration.initialUpdate', () => ({ collaboration: { initialUpdate }, experimentalWorkerOpen: true })],
  ['document', () => ({ document: parsed, documentBuffer: undefined, experimentalWorkerOpen: true })],
];
test.each(unsupportedSources)('%s falls back with one named warning', async (prop, props) => {
  render(editor(props()));
  await expectInThread();
  expect(warnings()).toHaveLength(1);
  expect(warnings()[0]).toContain(prop);
});

const viewerSources: Array<[string, DocxEditorProps]> = [
  ['readOnly', { readOnly: true }],
  ['viewing mode', { mode: 'viewing' }],
  ['readOnly with collaboration', { readOnly: true, collaboration: { clientId: 7 } }],
];
test.each(viewerSources)('%s with a parsed document uses the worker', async (_name, props) => {
  render(editor({
    document: parsed,
    documentBuffer: undefined,
    ...props,
  }));
  await waitFor(() => expect(openRequests()).toHaveLength(1));
  expect(workerModes.every(Boolean)).toBe(true);
  expect(mainThreadLoads).toBe(0);
  expect(warnings()).toEqual([]);
  expect(errors).toEqual([]);
});
