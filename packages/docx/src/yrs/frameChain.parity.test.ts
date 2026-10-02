import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { applyFrameChain, applyFrameDelta, decodeFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import { encodeDisplayListFrameExtras, type DisplayListBuildInputs } from '../layout/render/rustDisplayList';
import { preloadEditWasm } from '../wasm/edit';
import { decodeDocxHostJson, type YrsResidentWorkerSnapshot } from './index';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerRequestWithoutId,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';

type WorkerExports = { setFrameChainLimits: (bytes?: number, frames?: number) => void };
let startWorker: (scope: unknown, harness: unknown) => void;
const FONT = new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
)));

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  const modules: Record<string, string> = {
    './residentEngineSession':
      'export const createResidentEngineSession = async () => testHarness.session;',
    './wasm/index':
      'export const preloadEditWasm = async () => {}; export const preloadEditWasmFrom = async () => {};',
    '../wasm/loadWasmAsset': 'export const wasmModuleMemories = () => [];',
  };
  const result = await Bun.build({
    entrypoints: [resolve(import.meta.dir, 'residentEngineWorker.ts')],
    target: 'bun',
    format: 'iife',
    plugins: [{
      name: 'resident-worker-frame-chain',
      setup(build) {
        build.onLoad({ filter: /residentEngineWorker\.ts$/ }, ({ path }) => ({
          contents: readFileSync(path, 'utf8') + '\ntestHarness.setFrameChainLimits = setFrameChainLimits;',
          loader: 'ts',
        }));
        build.onResolve({ filter: /.*/ }, ({ path, importer }) =>
          importer.endsWith('/residentEngineWorker.ts') && path in modules
            ? { path, namespace: 'worker-test' }
            : undefined
        );
        build.onLoad({ filter: /.*/, namespace: 'worker-test' }, ({ path }) => ({
          contents: modules[path], loader: 'js',
        }));
      },
    }],
  });
  if (!result.success) throw new AggregateError(result.logs, 'Worker test bundle failed');
  startWorker = new Function(
    'self', 'testHarness', await result.outputs[0].text()
  ) as typeof startWorker;
});

function syntheticDocx(): Uint8Array {
  const parts: PartsMap = new Map();
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
  const body = Array.from({ length: 24 }, (_, index) =>
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}">` +
    '<w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>' +
    `Paragraph ${index + 1}. ${'alpha beta gamma delta '.repeat(8)}` +
    '</w:t></w:r></w:p>'
  ).join('');
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
    `<w:body>${body}<w:sectPr/></w:body></w:document>`
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function worker(session: ResidentEngineSession) {
  let nextId = 0;
  const replies = new Map<number, (reply: ResidentEngineWorkerResponse) => void>();
  const scope = {
    onmessage: (_event: { data: ResidentEngineWorkerRequest }) => {},
    postMessage(reply: ResidentEngineWorkerResponse, transfer: Transferable[] = []) {
      replies.get(reply.id)?.(structuredClone(reply, { transfer }));
      replies.delete(reply.id);
    },
  };
  const harness = { session, setFrameChainLimits: undefined as WorkerExports['setFrameChainLimits'] | undefined };
  startWorker(scope, harness);
  return {
    setFrameChainLimits: harness.setFrameChainLimits!,
    send(request: ResidentEngineWorkerRequestWithoutId): Promise<ResidentEngineWorkerResponse> {
      const id = ++nextId;
      return new Promise((resolve) => {
        replies.set(id, resolve);
        scope.onmessage({ data: { ...request, id } as ResidentEngineWorkerRequest });
      });
    },
  };
}

function frameReply(reply: ResidentEngineWorkerResponse) {
  if (!reply.ok || !reply.frame) throw new Error('Expected a worker frame: ' + JSON.stringify(reply));
  const frame = new Uint8Array(reply.frame);
  return {
    frame,
    frames: reply.pageFrames?.map((bytes) => new Uint8Array(bytes)) ?? [frame],
    delta: decodeFrameDelta(frame),
    pageFrames: reply.pageFrames,
  };
}

async function fixture() {
  const main = await createResidentEngineSession();
  const engine = await createResidentEngineSession();
  const w = worker(engine);
  try {
    const bytes = syntheticDocx();
    const host = decodeDocxHostJson(main.openDocx(bytes), bytes);
    const request = buildResidentRegionLayoutRequest(host.document, 24, {});
    const requirements = JSON.parse(main.layoutFontRequirementsJson(JSON.stringify(request))) as Array<{ key: string }>;
    main.clearFonts();
    const font = main.registerFont(FONT);
    const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [font]]));
    request.measurement = {
      fontChains,
      defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
      compat: { noLeading: false, doNotExpandShiftReturn: false },
      authoritativeShaping: true,
    };
    const layoutInput = JSON.stringify(request);
    const extras = encodeDisplayListFrameExtras({ fontChains } as DisplayListBuildInputs);
    const layoutExtras = JSON.stringify({ fontChains });
    main.layoutDocumentWithRegionsRetainedJson(layoutInput);
    main.buildDisplayListFrame(extras, 0);
    const caret = { story: 'body', paraId: main.geometryReader.paragraphs('body')[0]!.paraId, offset: 0 };
    main.setSelection(caret, caret);
    let revision = 0;
    const snapshot = (): YrsResidentWorkerSnapshot => ({
      clientId: 1, state: main.encodeState(), selection: main.selection(),
      fontsRevision: 1, fonts: [FONT], renderInputs: [], measureInputs: [],
      layoutInput, layoutWithRegions: true, layoutRevision: ++revision,
    });
    const base = applyFrameDelta(null, frameReply(await w.send({
      type: 'bootstrap', snapshot: snapshot(), extras, layoutExtras,
      displayWindow: [0, 2], retainBuiltPages: true, expectedFrameEpoch: 0,
    })).delta);
    expect(base.displayList.pages.length).toBeGreaterThan(20);
    expect(base.displayList.pages[12]!.unbuilt).toBe(true);
    return {
      engine, w, base,
      edit(text: string) { main.applyInput(text, 0); },
      async sync(expected: number, frameChain?: boolean) {
        return frameReply(await w.send({
          type: 'sync', snapshot: snapshot(), extras, layoutExtras,
          displayWindow: [0, 2], retainBuiltPages: true,
          expectedFrameEpoch: expected, paintCaret: false,
          ...(frameChain === undefined ? {} : { frameChain }),
        }));
      },
      async build(pages: number[], expected: number, frameChain?: boolean) {
        return frameReply(await w.send({
          type: 'buildPages', pages, expectedFrameEpoch: expected, paintCaret: false,
          ...(frameChain === undefined ? {} : { frameChain }),
        }));
      },
      async full() {
        return applyFrameDelta(null, frameReply(await w.send({
          type: 'buildPages', pages: [], expectedFrameEpoch: 0, paintCaret: false,
        })).delta);
      },
      destroy() {
        w.setFrameChainLimits();
        engine.destroy();
        main.destroy();
      },
    };
  } catch (error) {
    engine.destroy();
    main.destroy();
    throw error;
  }
}

function expectPages(frame: RetainedFrame, full: RetainedFrame): void {
  expect(frame.displayList.pages).toEqual(full.displayList.pages);
}

test('a page build the client never applied reaches it as a chain ahead of the next sync frame', async () => {
  for (const chained of [true, undefined]) {
    const f = await fixture();
    try {
      const built = await f.build([6, 7], f.base.frameEpoch);
      expect(built.delta.full).toBe(false);
      f.edit('First ');
      const reply = await f.sync(f.base.frameEpoch, chained);
      expect(reply.delta.full).toBe(!chained);
      if (chained) {
        expect(reply.frames).toHaveLength(2);
        expect(reply.frames[0]).toEqual(built.frame);
      } else {
        expect(reply.pageFrames).toBeUndefined();
      }
      const applied = applyFrameChain(f.base, reply.frames, false).frame;
      expect(applied.displayList.pages[6]!.unbuilt).not.toBe(true);
      expectPages(applied, await f.full());
    } finally {
      f.destroy();
    }
  }
});

test('superseded background slices reach the client in the next chained reply', async () => {
  const f = await fixture();
  const buildPages = f.engine.buildDisplayPagesFrame;
  try {
    const built = await f.build([5], f.base.frameEpoch);
    let sync: ReturnType<typeof f.sync> | undefined;
    let slice: Uint8Array | undefined;
    f.engine.buildDisplayPagesFrame = (pages, expected) => {
      const bytes = buildPages(pages, expected);
      if (!sync) {
        slice = bytes.slice();
        f.edit('Background ');
        sync = f.sync(f.base.frameEpoch, true);
      }
      return bytes;
    };
    const background = await f.w.send({
      type: 'buildPages', pages: [6, 7, 8, 9, 10, 11, 12, 13], background: true,
      expectedFrameEpoch: f.base.frameEpoch, frameChain: true, paintCaret: false,
    });
    expect(background.ok && background.pageBuildSuperseded).toBe(true);
    expect(sync).toBeDefined();
    const reply = await sync!;
    expect(reply.delta.full).toBe(false);
    expect(reply.frames).toHaveLength(3);
    expect(reply.frames[0]).toEqual(built.frame);
    expect(reply.frames[1]).toEqual(slice!);
    expect(decodeFrameDelta(slice!).baseFrameEpoch).toBe(built.delta.frameEpoch);
    f.engine.buildDisplayPagesFrame = buildPages;
    const applied = applyFrameChain(f.base, reply.frames, false).frame;
    expect(applied.displayList.pages[6]!.unbuilt).not.toBe(true);
    expect(applied.displayList.pages[10]!.unbuilt).toBe(true);
    expectPages(applied, await f.full());
  } finally {
    f.engine.buildDisplayPagesFrame = buildPages;
    f.destroy();
  }
});

test('a chain that does not start at the expected epoch falls back to a full frame', async () => {
  const f = await fixture();
  try {
    const first = await f.build([6], f.base.frameEpoch);
    await f.build([7], first.delta.frameEpoch);
    const reply = await f.sync(f.base.frameEpoch, true);
    expect(reply.delta.full).toBe(true);
    expect(reply.pageFrames).toBeUndefined();
    expectPages(applyFrameChain(f.base, reply.frames, false).frame, await f.full());
  } finally {
    f.destroy();
  }
});

test('a chain over its byte limit falls back to a full frame and releases its frames', async () => {
  const f = await fixture();
  try {
    f.w.setFrameChainLimits(1);
    const built = await f.build([6], f.base.frameEpoch);
    expect(built.frame.byteLength).toBeGreaterThan(1);
    f.w.setFrameChainLimits();
    const reply = await f.sync(f.base.frameEpoch, true);
    expect(reply.delta.full).toBe(true);
    expect(reply.pageFrames).toBeUndefined();
    const base = applyFrameChain(f.base, reply.frames, false).frame;
    const fresh = await f.build([7], base.frameEpoch);
    const next = await f.sync(base.frameEpoch, true);
    expect(next.delta.full).toBe(false);
    expect(next.frames).toHaveLength(2);
    expect(next.frames[0]).toEqual(fresh.frame);
    expectPages(applyFrameChain(base, next.frames, false).frame, await f.full());
  } finally {
    f.destroy();
  }
});

test('a dropped scroll sweep and edit reach the client across the next edit', async () => {
  const f = await fixture();
  try {
    const first = await f.build([14, 15], f.base.frameEpoch, true);
    const second = await f.build([16, 17], f.base.frameEpoch, true);
    expect(second.delta.full).toBe(false);
    expect(second.frames[0]).toEqual(first.frame);
    f.edit('First ');
    const overtaken = await f.sync(f.base.frameEpoch, true);
    expect(overtaken.delta.full).toBe(false);
    f.edit('Second ');
    const reply = await f.sync(f.base.frameEpoch, true);
    expect(reply.delta.full).toBe(false);
    expect(reply.frames).toHaveLength(4);
    expect(reply.frames.slice(0, 3)).toEqual([first.frame, second.frame, overtaken.frame]);
    const originalPages = structuredClone(f.base.displayList.pages);
    const applied = applyFrameChain(f.base, reply.frames, false);
    expect(f.base.displayList.pages).toEqual(originalPages);
    expect(applied.delta).toEqual(reply.delta);
    for (const index of [14, 15, 16, 17]) {
      expect(applied.frame.displayList.pages[index]!.unbuilt).not.toBe(true);
    }
    const full = await f.full();
    expectPages(applied.frame, full);
    const advanced = applyFrameDelta(f.base, decodeFrameDelta(first.frame));
    expectPages(applyFrameChain(advanced, reply.frames, true).frame, full);
    expect(() => applyFrameChain(applied.frame, reply.frames, false)).toThrow('not newer');
  } finally {
    f.destroy();
  }
});
