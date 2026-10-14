import { expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { runInNewContext } from 'node:vm';
import JSZip from 'jszip';
import { isClientMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { mainThreadWasmSpies } from '../../../../shared/office-session/testing/wasmWorker';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import type {
  PresentationCommandClient, PresentationCommandEnvelope,
  PresentationCommandOpenResult, PresentationCommandSelection,
} from './commandProtocol';

async function isolated(name: string): Promise<boolean> {
  if (process.env.PPTX_COMMAND_WASM_TEST === name) return false;
  const child = Bun.spawn([process.execPath, 'test', import.meta.path, '-t', name], {
    env: { ...process.env, PPTX_COMMAND_WASM_TEST: name }, stdout: 'pipe', stderr: 'pipe',
  });
  const [status, stdout, stderr] = await Promise.all([
    child.exited, new Response(child.stdout).text(), new Response(child.stderr).text(),
  ]);
  expect({ status, output: status === 0 ? '' : stdout + stderr }).toEqual({ status: 0, output: '' });
  return true;
}

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>alpha</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>');
  return zip.generateAsync({ type: 'uint8array' });
}

function workerTransport() {
  const pair = createInProcessPair();
  const requests: string[] = [];
  const errors = new Set<(error: unknown) => void>();
  const initializations: Array<ArrayBuffer | WebAssembly.Module | undefined> = [];
  const workerWasm = {
    Module: WebAssembly.Module, Instance: WebAssembly.Instance,
    Memory: WebAssembly.Memory, Table: WebAssembly.Table, Global: WebAssembly.Global,
    CompileError: WebAssembly.CompileError, LinkError: WebAssembly.LinkError,
    RuntimeError: WebAssembly.RuntimeError,
    compile: WebAssembly.compile.bind(WebAssembly),
    compileStreaming: WebAssembly.compileStreaming.bind(WebAssembly),
    instantiate: WebAssembly.instantiate.bind(WebAssembly),
    instantiateStreaming: WebAssembly.instantiateStreaming.bind(WebAssembly),
    validate: WebAssembly.validate.bind(WebAssembly),
  };
  let startup: Promise<void> | undefined;
  let closed = false;

  async function start(): Promise<void> {
    const workerPath = resolve(import.meta.dir, 'worker.ts');
    const result = await Bun.build({
      entrypoints: [workerPath], target: 'bun', format: 'esm',
      plugins: [{ name: 'pptx-command-worker', setup(build) {
        build.onLoad({ filter: /\/session\/worker\.ts$/ }, () => ({
          loader: 'ts', contents: `
            import { createSessionHost } from ${JSON.stringify(resolve(import.meta.dir, '../../../../shared/office-session/host.ts'))};
            import { registerPresentationCommandHandlers } from ${JSON.stringify(resolve(import.meta.dir, 'commandHost.ts'))};
            import { initWasm } from ${JSON.stringify(resolve(import.meta.dir, '../wasm/loader.ts'))};
            const registration = registerPresentationCommandHandlers({
              initWasm: async (source) => {
                testHarness.initializations.push(source);
                await initWasm(source ?? testHarness.wasm);
              }
            });
            createSessionHost(testHarness.transport, {
              handlers: registration.handlers, policies: registration.policies,
              context: null, onDispose: registration.dispose
            });
          `,
        }));
      } }],
    });
    expect(result.success).toBe(true);
    if (!result.success) throw new Error(result.logs.join('\n'));
    const wasm = await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url));
    if (closed) return;
    const code = (await result.outputs[0]!.text()).split('import.meta.url')
      .join(JSON.stringify(pathToFileURL(workerPath).href));
    runInNewContext(code, {
      testHarness: { transport: pair.host, wasm: new Uint8Array(wasm), initializations },
      WebAssembly: workerWasm, Uint8Array, ArrayBuffer, URL, TextEncoder, TextDecoder,
      setTimeout, clearTimeout, queueMicrotask, performance, crypto,
    });
  }

  const transport: SessionTransport = {
    post(message, transfer = []) {
      if (isClientMessage(message)) requests.push(message.kind === 'call' ? message.method : message.kind);
      const cloned: unknown = structuredClone(message, { transfer });
      startup ??= start();
      void startup.then(() => { if (!closed) pair.client.post(cloned); }).catch((error: unknown) => {
        for (const listener of errors) listener(error);
      });
    },
    listen: pair.client.listen,
    onError(listener) {
      errors.add(listener);
      const unsubscribe = pair.client.onError(listener);
      return () => { errors.delete(listener); unsubscribe(); };
    },
    close() { closed = true; errors.clear(); pair.client.close(); },
  };
  return { transport, requests, initializations };
}

test('p1_command_session_never_initializes_main_wasm', async () => {
  if (await isolated('p1_command_session_never_initializes_main_wasm')) return;
  const worker = workerTransport();
  const source = await deck();
  const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const spies = mainThreadWasmSpies();
  let client: PresentationCommandClient | undefined;
  try {
    const { createPresentationCommandClient } = await import('../index');
    client = await createPresentationCommandClient(source, {
      session: 'p1-cold', incarnation: 'p1-worker', clientId: 7812,
      fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }],
    }, worker.transport);
    expect(client.state.frame!.displayList.byteLength).toBeGreaterThan(0);
    const read = await client.read();
    expect(read.result.ok).toBe(true);
    if (!read.result.ok) throw new Error(read.result.failure.message);
    const story = read.result.stories[0]!;
    const selection: PresentationCommandSelection = { slideId: story.slideId,
      shapeId: story.shapeId, storyId: story.storyId, anchor: 0, focus: 0, affinity: 'downstream' };
    const ack = await client.dispatch({ kind: 'insertText', payload: {
      inputTime: 100, undoGroup: 'p1-text', at: { source: 'confirmed', selection }, text: 'worker ',
    } });
    expect(ack.frame).toMatchObject({ coversThroughSeq: ack.throughSeq, slideId: story.slideId });
    expect((await client.flush()).throughSeq).toBe(ack.throughSeq);
    expect(await client.thumbnail(story.slideId, ack.frame!.slideRevision)).toMatchObject({ status: 'unchanged' });
    const saved = await JSZip.loadAsync(await client.save());
    expect(await saved.file('ppt/slides/slide1.xml')!.async('string')).toContain('worker ');
    await client.dispose();
    expect(worker.initializations).toHaveLength(1);
    expect(worker.requests.filter((method) => method === 'openCommandEditor')).toHaveLength(1);
    expect(worker.requests).not.toContain('beginEditor');
    expect(worker.requests).not.toContain('attachPeer');
    expect(worker.requests).not.toContain('editorFrame');
    expect(worker.requests).not.toContain('frame');
    for (const spy of spies) expect(spy).not.toHaveBeenCalled();
  } finally {
    try { await client?.dispose(); } finally {
      worker.transport.close();
      for (const spy of spies) spy.mockRestore();
    }
  }
}, 30_000);

test('p1_public_command_client_import_and_dispatch_do_not_load_main_wasm', async () => {
  if (await isolated('p1_public_command_client_import_and_dispatch_do_not_load_main_wasm')) return;
  const spies = mainThreadWasmSpies();
  const initialize = mock(() => { throw new Error('Main wasm initialization'); });
  const load = mock(() => { throw new Error('Main wasm load'); });
  const forbiddenFetch = (() => { throw new Error('Main wasm fetch'); }) as unknown as typeof globalThis.fetch;
  const fetch = spyOn(globalThis, 'fetch').mockImplementation(forbiddenFetch);
  mock.module('../wasm/generated/pptx_wasm.js', () => ({
    default: initialize, initSync: initialize,
    PptxDocument: class { static openCollaborative = load; static open = load; },
    PptxRenderer: class { static open = load; },
    decodeTiffPng: load, exportPptxMarkdownJson: load, exportPptxStructuredJson: load,
    parsePptxJson: load, parsePptxJsonWithoutMedia: load, renderPptxMarkdownJson: load, rendererVersion: load,
  }));
  const pair = createInProcessPair();
  const requests: string[] = [];
  let text = 'alpha';
  const selection: PresentationCommandSelection = { slideId: 'slide', shapeId: 'shape',
    storyId: 'story', anchor: 0, focus: 0, affinity: 'downstream' };
  let state: PresentationCommandOpenResult = {
    session: 'p1-public', incarnation: 'p1-cold', epoch: 1, throughSeq: 0, version: 'v0',
    selection, outcomes: [], dirtySlideIds: [],
    slides: [{ id: 'slide', index: 0, name: null, layoutPartPath: null }],
    size: { width: 960, height: 720 }, activeSlideId: 'slide',
    editState: { readOnly: false, canUndo: false, canRedo: false,
      toolbar: { target: selection, storedMarks: {}, formatting: null, alignment: null,
        hyperlink: null, table: null, canEditText: true } },
    frame: { slideId: 'slide', slideIndex: 0, slideRevision: 0, coversThroughSeq: 0, documentEpoch: 1,
      incarnation: 'p1-cold', version: 'v0', epoch: 1,
      displayList: new TextEncoder().encode(text).buffer, media: [] },
  };
  const unlisten = pair.host.listen((message) => {
    if (!isClientMessage(message) || message.kind !== 'call') return;
    requests.push(message.method);
    if (message.method === 'command') {
      const command = message.args[0] as PresentationCommandEnvelope;
      expect(command.kind).toBe('insertText');
      if (command.kind !== 'insertText') throw new Error('Unexpected command');
      text = command.payload.text + text;
      state = { ...state, throughSeq: command.seq, version: 'v1',
        outcomes: [{ seq: command.seq, status: 'applied', committed: true }], dirtySlideIds: ['slide'],
        selection: { ...selection, anchor: command.payload.text.length, focus: command.payload.text.length },
        frame: { ...state.frame!, slideRevision: 1, coversThroughSeq: command.seq,
          version: 'v1', epoch: 2, displayList: new TextEncoder().encode(text).buffer } };
    } else expect(message.method).toBe('openCommandEditor');
    pair.host.post({ protocol: 1, kind: 'reply', id: message.id, ok: true, value: state });
  });
  let client: PresentationCommandClient | undefined;
  try {
    const { createPresentationCommandClient } = await import('../index');
    client = await createPresentationCommandClient(new Uint8Array([1, 2, 3]), {
      session: state.session, incarnation: state.incarnation,
    }, pair.client);
    expect(new TextDecoder().decode(client.state.frame!.displayList)).toBe('alpha');
    const ack = await client.dispatch({ kind: 'insertText', payload: { inputTime: 100,
      undoGroup: 'p1-public', at: { source: 'confirmed', selection }, text: 'worker ' } });
    expect(ack.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    expect(new TextDecoder().decode(ack.frame!.displayList)).toBe('worker alpha');
    expect(text).toBe('worker alpha');
    expect(requests).toEqual(['openCommandEditor', 'command']);
    expect(initialize).not.toHaveBeenCalled();
    expect(load).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
    for (const spy of spies) expect(spy).not.toHaveBeenCalled();
  } finally {
    try { await client?.dispose(); } finally {
      unlisten(); pair.client.close(); fetch.mockRestore();
      for (const spy of spies) spy.mockRestore();
      mock.restore();
    }
  }
}, 30_000);

test('p1_main_wasm_guard_detects_transitive_import_initialization', async () => {
  if (await isolated('p1_main_wasm_guard_detects_transitive_import_initialization')) return;
  const spies = mainThreadWasmSpies();
  const initialize = mock(() => WebAssembly.compile(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0])));
  let initialization: Promise<WebAssembly.Module> | undefined;
  mock.module('../wasm/asset', () => {
    initialization = initialize();
    return { wasmAssetUrl: () => new URL('https://example.test/pptx.wasm') };
  });
  try {
    expect(initialize).not.toHaveBeenCalled();
    for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    await import('../index');
    await initialization;
    expect(initialize).toHaveBeenCalledTimes(1);
    expect(spies[4]).toHaveBeenCalledTimes(1);
    expect(() => {
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    }).toThrow();
  } finally {
    for (const spy of spies) spy.mockRestore();
    mock.restore();
  }
}, 30_000);
