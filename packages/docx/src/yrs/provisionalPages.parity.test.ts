import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { applyFrameDelta, decodeFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import {
  encodeDisplayListFrameExtras,
  type DisplayListBuildInputs,
} from '../layout/render/rustDisplayList';
import { preloadEditWasm } from '../wasm/edit';
import { decodeDocxHostJson, type YrsResidentWorkerSnapshot } from './index';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerRequestWithoutId,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
)));
let startWorker: (scope: unknown, harness: unknown) => void;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
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
      name: 'resident-worker-real-engine',
      setup(build) {
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
  startWorker = new Function('self', 'testHarness', await result.outputs[0].text()) as typeof startWorker;
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
  const body = Array.from({ length: 80 }, (_, index) =>
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}">` +
    '<w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>' +
    `Paragraph ${index + 1}. ${'alpha beta gamma delta epsilon zeta eta theta '.repeat(8 + index % 5)}` +
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
    postMessage(reply: ResidentEngineWorkerResponse) {
      replies.get(reply.id)?.(reply);
      replies.delete(reply.id);
    },
  };
  startWorker(scope, { session });
  return {
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
