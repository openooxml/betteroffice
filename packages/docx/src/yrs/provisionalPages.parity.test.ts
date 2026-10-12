import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { applyFrameDelta, decodeFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import {
  encodeDisplayListFrameExtras,
  loadRustDisplayListQueryEngine,
  loadedRustDisplayListQueryEngine,
  type DisplayListBuildInputs,
} from '../layout/render/rustDisplayList';
import { preloadEditWasm } from '../wasm/edit';
import { decodeDocxHostJson, type YrsResidentWorkerSnapshot } from './index';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandFrame, ResidentCommandInitRequest } from './residentCommandProtocol';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerRequestWithoutId,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
)));
let startWorker: (scope: unknown, harness: unknown, turns: unknown) => void;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  const modules: Record<string, string> = {
    './yrsSessionFacade': 'export const decodeDocxHostJson = () => { throw new Error("Unused source opening"); };',
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
      name: 'resident-worker-real-engine',
      setup(build) {
        build.onLoad({ filter: /\/residentCommandPresentation\.ts$/ }, () => ({
          contents: 'export const stageCommandPresentation = async ({ layout }) => ({ value: { frameEpoch: layout.frame.frameEpoch, pages: [] }, transfer: [] });',
          loader: 'js',
        }));
        build.onLoad({ filter: /\/rustDisplayList\.ts$/ }, () => ({
          contents: `
            export const encodeDisplayListFrameExtras = (...args) => testHarness.encodeDisplayListFrameExtras(...args);
            export const loadRustDisplayListQueryEngine = () => testHarness.loadRustDisplayListQueryEngine();
            export const loadedRustDisplayListQueryEngine = () => testHarness.loadedRustDisplayListQueryEngine();
          `,
          loader: 'js',
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
  startWorker = new Function('self', 'testHarness', 'MessageChannel', await result.outputs[0].text()) as typeof startWorker;
});

function syntheticDocx(feature = 'plain'): Uint8Array {
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
  const body = Array.from({ length: 80 }, (_, index) =>
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}">` +
    '<w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>' +
    `Paragraph ${index + 1}. ${'alpha beta gamma delta epsilon zeta eta theta '.repeat(8 + index % 5)}` +
    '</w:t></w:r></w:p>'
  ).join('');
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
    `<w:body>${featureBody(feature, body)}${feature === 'columns' ? '<w:sectPr><w:cols w:num="2"/></w:sectPr>' : '<w:sectPr/>'}</w:body></w:document>`
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function worker(session: ResidentEngineSession, controlled = false) {
  const turns: (() => void)[] = [];
  const frames: DocxCommandFrame[] = [];
  let nextId = 0;
  const replies = new Map<number, (reply: ResidentEngineWorkerResponse) => void>();
  const scope = {
    onmessage: (_event: { data: ResidentEngineWorkerRequest }) => {},
    postMessage(reply: ResidentEngineWorkerResponse) {
      if (reply.ok && reply.commandFrame) frames.push(reply.commandFrame.frame);
      replies.get(reply.id)?.(reply);
      replies.delete(reply.id);
    },
  };
  startWorker(scope, { session, encodeDisplayListFrameExtras, loadRustDisplayListQueryEngine, loadedRustDisplayListQueryEngine }, controlled ? class {
    port1 = { onmessage: () => {} };
    port2 = { postMessage: () => turns.push(() => this.port1.onmessage()) };
  } : MessageChannel);
  return {
    frames,
    async turn() { turns.shift()?.(); await new Promise<void>((resolve) => setTimeout(resolve, 0)); },
    async drain() { for (let i = 0; i < 2000 && turns.length; i += 1) { turns.shift()!(); await new Promise<void>((resolve) => setTimeout(resolve, 0)); } },
    send(request: ResidentEngineWorkerRequestWithoutId): Promise<ResidentEngineWorkerResponse> {
      const id = ++nextId;
      return new Promise((resolve) => {
        replies.set(id, resolve);
        scope.onmessage({ data: { ...request, id } as ResidentEngineWorkerRequest });
      });
    },
  };
}

function adopt(reply: ResidentEngineWorkerResponse, previous: RetainedFrame | null): RetainedFrame {
  if (!reply.ok || !reply.frame) throw new Error('Expected a worker frame: ' + JSON.stringify(reply));
  return applyFrameDelta(previous, decodeFrameDelta(reply.frame));
}

function expectExactPrefix(frame: RetainedFrame, full: RetainedFrame): void {
  expect(frame.displayList.pages.length).toBeGreaterThan(3);
  const built = frame.displayList.pages.filter((page) => !page.unbuilt);
  for (const page of built) {
    expect(page.pageIndex).toBeLessThan(3);
    expect(page).toEqual(full.displayList.pages[page.pageIndex]);
  }
  for (const page of frame.displayList.pages.slice(3)) expect(page.unbuilt).toBe(true);
}

for (const path of ['bootstrap', 'sync', 'no-window'] as const) {
  test(`whole-document provisional ${path} sends only pages equal to a cold full layout`, async () => {
    const bytes = syntheticDocx();
    const cold = await createResidentEngineSession();
    const engine = await createResidentEngineSession();
    const w = worker(engine);
    try {
      const host = decodeDocxHostJson(cold.openDocx(bytes), bytes);
      const request = buildResidentRegionLayoutRequest(host.document, 24, {});
      const requirements = JSON.parse(cold.layoutFontRequirementsJson(JSON.stringify(request))) as Array<{ key: string }>;
      cold.clearFonts();
      const font = cold.registerFont(FONT);
      const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [font]]));
      request.measurement = {
        fontChains,
        defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
        compat: { noLeading: false, doNotExpandShiftReturn: false },
        authoritativeShaping: true,
      };
      const layoutInput = JSON.stringify(request);
      cold.layoutDocumentWithRegionsRetainedJson(layoutInput);
      cold.setDisplayWindow(0, path === 'no-window' ? 2 ** 32 - 1 : 6);
      cold.setDisplayRetainBuiltPages(false);
      cold.setWindowedIncrementalBuilds(true);
      const extras = encodeDisplayListFrameExtras({ fontChains } as DisplayListBuildInputs);
      const full = applyFrameDelta(null, decodeFrameDelta(cold.buildDisplayListFrame(extras, 0)));
      expect(full.displayList.pages.length).toBeGreaterThan(6);

      expect((await w.send({ type: 'open', bytes: bytes.slice().buffer })).ok).toBe(true);
      const caret = { story: 'body', paraId: engine.geometryReader.paragraphs('body')[4]!.paraId, offset: 0 };
      const snapshot: YrsResidentWorkerSnapshot = {
        clientId: 1,
        state: cold.encodeState(),
        fontsRevision: 1,
        fonts: [FONT],
        renderInputs: [],
        measureInputs: [],
        layoutInput,
        layoutWithRegions: true,
        layoutRevision: 1,
        selection: path === 'no-window' ? null : { anchor: caret, head: caret },
      };
      let frame: RetainedFrame | null = null;
      if (path === 'sync') {
        frame = adopt(await w.send({
          type: 'bootstrap', opened: true, snapshot, displayWindow: [0, 6],
          retainBuiltPages: true, extras, layoutExtras: JSON.stringify({ fontChains }), expectedFrameEpoch: 0,
        }), null);
        expect(frame.displayList.pages[4]?.unbuilt).not.toBe(true);
        snapshot.workerAuthoritative = true;
        snapshot.fontsRevision += 1;
      }
      const options = {
        snapshot, provisionalPages: 3,
        ...(path === 'no-window' ? {} : { displayWindow: [0, 6] as [number, number] }),
        retainBuiltPages: true, extras, layoutExtras: JSON.stringify({ fontChains }),
        expectedFrameEpoch: frame?.frameEpoch ?? 0,
      };
      const prefix = path === 'sync'
        ? await w.send({ type: 'sync', paintCaret: false, ...options })
        : await w.send({ type: 'bootstrap', opened: true, ...options });
      expect(prefix.ok && prefix.layoutProvisional).toBe(true);
      expect(prefix.ok && JSON.parse(prefix.layoutJson!).layout.pages.length).toBeGreaterThan(3);
      frame = adopt(prefix, frame);
      expect(frame.displayList.pages.filter((page) => !page.unbuilt)).toHaveLength(3);
      expectExactPrefix(frame, full);
      for (const background of [false, true]) {
        frame = adopt(await w.send({
          type: 'buildPages', pages: [2, 3, 4, 5], background,
          expectedFrameEpoch: frame.frameEpoch, paintCaret: false,
        }), frame);
        expectExactPrefix(frame, full);
      }
      if (path === 'sync') {
        frame = adopt(await w.send({
          type: 'buildFrame', extras, displayWindow: [3, 6], retainBuiltPages: true,
          expectedFrameEpoch: frame.frameEpoch, paintCaret: false,
        }), frame);
        expectExactPrefix(frame, full);
        frame = adopt(await w.send({
          type: 'buildFrame', extras, expectedFrameEpoch: frame.frameEpoch, paintCaret: false,
        }), frame);
        expectExactPrefix(frame, full);
      }
      const completed = await w.send({
        type: 'completeLayout', expectedFrameEpoch: frame.frameEpoch, paintCaret: false,
        ...(path === 'sync' ? { sliceBlocks: 8 } : {}),
      });
      expect(completed.ok && completed.layoutProvisional).toBeUndefined();
      frame = adopt(completed, frame);
      expect(frame.displayList.pages.length).toBe(full.displayList.pages.length);
      for (const index of [3, 4, 5]) {
        expect(frame.displayList.pages[index]?.unbuilt).not.toBe(true);
        expect(frame.displayList.pages[index]).toEqual(full.displayList.pages[index]);
      }
      if (path === 'no-window') expect(frame.displayList.pages).toEqual(full.displayList.pages);
      frame = adopt(await w.send({
        type: 'buildPages', pages: [4], expectedFrameEpoch: frame.frameEpoch, paintCaret: false,
      }), frame);
      expect(frame.displayList.pages[4]).toEqual(full.displayList.pages[4]);
    } finally {
      engine.destroy();
      cold.destroy();
    }
  });
}


function featureBody(feature: string, body: string): string {
  const field = (type: string) => `<w:p><w:fldSimple w:instr=" ${type} "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>`;
  if (['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'DATE', 'TIME'].includes(feature)) return field(feature) + body;
  if (feature === 'table') return '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    Array.from({ length: 8 }, (_, i) => `<w:tr><w:tc><w:p><w:r><w:t>Row ${i}</w:t></w:r></w:p></w:tc></w:tr>`).join('') + '</w:tbl>' + body;
  if (feature === 'keep') return body.replace('<w:pageBreakBefore/>', '<w:keepNext/><w:keepLines/>');
  if (feature === 'section') return body.replace(
    '<w:p w14:paraId="00000029">', '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p><w:p w14:paraId="00000029">');
  if (feature === 'float') return '<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="margin" w:horzAnchor="margin" w:tblpX="0" w:tblpY="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Float</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' + body;
  if (feature === 'footnote' || feature === 'endnote') return `<w:p><w:r><w:${feature}Reference w:id="1"/></w:r></w:p>` + body;
  return body;
}

async function commandFixture(feature = 'plain') {
  const bytes = syntheticDocx(feature);
  const cold = await createResidentEngineSession();
  const engine = await createResidentEngineSession();
  const w = worker(engine, true);
  const host = decodeDocxHostJson(cold.openDocx(bytes), bytes);
  const request = buildResidentRegionLayoutRequest(host.document, 24, {});
  if (feature === 'footnote' || feature === 'endnote') {
    cold.proposalEngine.readParagraphs({ view: 'accepted' });
    request.notes = { contents: [{ id: 1, noteKind: feature, height: 0 }] } as typeof request.notes;
  }
  const requirements = JSON.parse(cold.layoutFontRequirementsJson(JSON.stringify(request))) as Array<{ key: string }>;
  cold.clearFonts();
  const font = cold.registerFont(FONT);
  const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [font]]));
  request.measurement = { fontChains, defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, authoritativeShaping: true };
  const layoutInput = JSON.stringify(request);
  cold.layoutDocumentWithRegionsRetainedJson(layoutInput);
  cold.setDisplayWindow(0, 3);
  cold.setDisplayRetainBuiltPages(false);
  cold.setWindowedIncrementalBuilds(true);
  const extras = JSON.stringify({ fontChains });
  const full = applyFrameDelta(null, decodeFrameDelta(cold.buildDisplayListFrame(extras, 0)));
  await w.send({ type: 'open', bytes: bytes.slice().buffer });
  const methods = ['layoutDocumentWithRegionsRetained', 'layoutDocumentWithRegionsRetainedJson', 'beginRegionLayout', 'resumeRegionLayout'] as const;
  const calls: string[] = [];
  for (const method of methods) {
    const original = engine[method].bind(engine) as (...args: never[]) => unknown;
    (engine as unknown as Record<string, unknown>)[method] = (...args: never[]) => { calls.push(method); return original(...args); };
  }
  const init: ResidentCommandInitRequest = {
    session: 'viewport-first', incarnation: 1, epoch: 1, readOnly: false,
    fonts: [FONT], layoutInput, layoutWithRegions: true, layoutRevision: 1, extras,
    displayWindow: [0, 3], ...{ provisionalPages: 3 },
  };
  const reply = await w.send({ type: 'commandInit', request: init });
  if (!reply.ok || !reply.commandInitial) throw new Error(JSON.stringify(reply));
  const initial = reply.commandInitial;
  const frame = applyFrameDelta(null, decodeFrameDelta(initial.frame.payload.frame));
  return { cold, engine, w, init, initial, frame, full, calls,
    destroy() { engine.destroy(); cold.destroy(); },
    async present() {
      return w.send({ type: 'commandViewport', request: { start: 0, end: 3, frameEpoch: frame.frameEpoch } });
    },
  };
}

for (const feature of ['plain', 'table', 'keep', 'section', 'PAGE']) {
  test(`command prefix ${feature} is exact and bounded before a full pass`, async () => {
    const f = await commandFixture(feature);
    try {
      expect(f.initial.frame.payload.layoutProvisional).toBe(true);
      expect(f.calls).toEqual([]);
      expect(f.frame.pages).toHaveLength(3);
      expect(f.initial.frame.payload.parts ?? []).toEqual([]);
      for (const { page, pageIndex } of f.frame.pages) {
        expect(page.unbuilt).not.toBe(true);
        expect(page).toEqual(f.full.displayList.pages[pageIndex]);
      }
      expect(f.initial.frame.payload.frame.byteLength).toBeLessThan(500_000);
      await f.present();
      await f.w.turn();
      expect(f.calls).toContain('beginRegionLayout');
      await f.w.drain();
      expect(f.calls.filter(call => call === 'resumeRegionLayout').length).toBeGreaterThan(1);
      const completed = f.w.frames.at(-1)!;
      expect(completed.payload.layoutProvisional).not.toBe(true);
      const final = commandFramesAfter(f.frame, f.w.frames);
      expect(final.pages.length).toBe(f.full.pages.length);
      for (let i = 0; i < 3; i += 1) expect(final.displayList.pages[i]).toEqual(f.full.displayList.pages[i]);
    } finally { f.destroy(); }
  });
}

for (const feature of ['NUMPAGES', 'SECTIONPAGES', 'footnote', 'endnote', 'float', 'columns', 'DATE', 'TIME']) {
  test(`command prefix excludes ${feature} and returns the cold full layout`, async () => {
    const f = await commandFixture(feature);
    try {
      expect(f.initial.frame.payload.layoutProvisional).not.toBe(true);
      expect(f.frame.displayList.pages).toEqual(f.full.displayList.pages);
      await f.present();
      await f.w.drain();
      expect(f.calls).not.toContain('beginRegionLayout');
    } finally { f.destroy(); }
  });
}

test('command prefix keeps geometry uncertified until a completed frame', async () => {
  const f = await commandFixture();
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    const version = f.engine.proposalEngine.version();
    const parts = await f.w.send({ type: 'documentRead', read: { kind: 'partSnapshot' } });
    expect(parts.ok && parts.read?.value).toMatchObject({ geometry: [], frameEpoch: null });
    const layout = await f.w.send({ type: 'layoutJson', layoutRevision: 1 });
    expect(layout.ok && layout.layoutJsonStatus).toBe('stale');
    const highlights = await f.w.send({ type: 'documentRead', read: { kind: 'rangeHighlights', ranges: [], expectVersion: version } });
    expect(highlights.ok && highlights.read?.value).toBeNull();
    await f.present();
    await f.w.drain();
    const completed = await f.w.send({ type: 'layoutJson', layoutRevision: 1 });
    expect(completed.ok && completed.layoutJsonStatus).toBe('ok');
    const certified = await f.w.send({ type: 'documentRead', read: { kind: 'partSnapshot' } });
    expect(certified.ok && (certified.read?.value as { frameEpoch: number }).frameEpoch).toBe(f.w.frames.at(-1)!.revision);
  } finally { f.destroy(); }
});

test('command viewport beyond the prefix completes before painting and cancels stale slices', async () => {
  const f = await commandFixture();
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    await f.present();
    await f.w.turn();
    await f.w.send({ type: 'commandViewport', request: { start: 40, end: 43, frameEpoch: f.w.frames.at(-1)!.revision } });
    const away = f.w.frames.at(-1)!;
    const frame = commandFramesAfter(f.frame, f.w.frames);
    expect(away.payload.layoutProvisional).not.toBe(true);
    f.cold.setDisplayWindow(40, 43);
    const cold = applyFrameDelta(f.full, decodeFrameDelta(f.cold.buildDisplayListFrame(f.init.extras, f.full.frameEpoch)));
    for (const index of [40, 41, 42]) expect(frame.displayList.pages[index]).toEqual(cold.displayList.pages[index]);
    const count = f.w.frames.length;
    await f.w.drain();
    expect(f.w.frames).toHaveLength(count);
  } finally { f.destroy(); }
});

test('an edit overtakes sliced completion and finishes it before changing the document', async () => {
  const f = await commandFixture();
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    await f.present();
    await f.w.turn();
    const point = { story: 'body', paraId: f.engine.geometryReader.paragraphs('body')[0]!.paraId, offset: 0 };
    const edited = await f.w.send({ type: 'edit', command: {
      session: f.init.session, incarnation: 1, seq: 1, basis: { epoch: 1, version: 0 },
      kind: 'insertText', payload: { text: 'Edit ', selection: { anchor: point, focus: point } },
    } });
    expect(edited.ok && edited.commandAck?.outcomes[0]?.status).toBe('applied');
    expect(edited.ok && edited.commandAck?.frame?.payload.layoutProvisional).not.toBe(true);
    const count = f.w.frames.length;
    await f.w.drain();
    expect(f.w.frames).toHaveLength(count);
  } finally { f.destroy(); }
});

test('collaboration and replacement cannot publish an abandoned command completion', async () => {
  for (const replace of [false, true]) {
    const f = await commandFixture();
    try {
      expect(f.initial.frame.payload.layoutProvisional).toBe(true);
      await f.present();
      await f.w.turn();
      if (replace) await f.w.send({ type: 'open', bytes: syntheticDocx('PAGE').buffer });
      else await f.w.send({ type: 'commandUpdate', update: f.cold.encodeState() });
      const count = f.w.frames.length;
      await f.w.drain();
      expect(f.w.frames).toHaveLength(count);
    } finally { f.destroy(); }
  }
});


function commandFramesAfter(initial: RetainedFrame, frames: DocxCommandFrame[]): RetainedFrame {
  let result = initial;
  for (const frame of frames) {
    for (const bytes of [...frame.payload.pageFrames ?? [], frame.payload.frame]) result = applyFrameDelta(result, decodeFrameDelta(bytes));
  }
  return result;
}
