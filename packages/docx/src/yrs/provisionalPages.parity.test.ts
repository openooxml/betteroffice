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
import { echoBackgroundPage } from './residentCommandPresentation';
import type {
  DocxCommandAck, DocxCommandEnvelope, DocxCommandFrame, DocxCommandSelection,
  DocxSelectionPointIntent, ResidentCommandInitRequest,
} from './residentCommandProtocol';
import type { DocxPageExportOptions, DocxPageFragment } from './pagedExport';
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
          contents: `
            export const echoBackgroundPage = (...args) => testHarness.echoBackgroundPage(...args);
            export const stageCommandPresentation = async ({ layout }) => ({ value: { frameEpoch: layout.frame.frameEpoch, pages: [] }, transfer: [] });
          `,
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
  startWorker = new Function('self', 'testHarness', 'MessageChannel', `
    const performance = testHarness.clock ?? globalThis.performance;
    const setTimeout = testHarness.setTimer ?? globalThis.setTimeout;
    const clearTimeout = testHarness.clearTimer ?? globalThis.clearTimeout;
    ${await result.outputs[0].text()}
  `) as typeof startWorker;
});

function syntheticDocx(feature = 'plain'): Uint8Array {
  if (feature.startsWith('flow-')) return flowDocx(feature);
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
    '<w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr>' +
    (feature === 'echo' ? '<w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/>' : '') +
    '<w:sz w:val="22"/></w:rPr><w:t>' +
    `Paragraph ${index + 1}. ${feature === 'echo' ? 'alpha beta gamma' : 'alpha beta gamma delta epsilon zeta eta theta '.repeat(8 + index % 5)}` +
    '</w:t></w:r></w:p>'
  ).join('');
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<w:body>${featureBody(feature, body)}${featureSection(feature)}</w:body></w:document>`
  ));
  if (feature.startsWith('header-') || feature.startsWith('footer-') || feature === 'footnote' || feature === 'endnote') {
    const kind = feature.split('-')[0]!;
    const part = kind === 'footnote' || kind === 'endnote' ? `${kind}s` : kind;
    const content = kind === 'footnote' || kind === 'endnote'
      ? `<w:${part}><w:${kind} w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:${kind}></w:${part}>`
      : `<w:${kind === 'header' ? 'hdr' : 'ftr'}><w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:fldSimple w:instr=" ${feature.split('-')[1]} "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl></w:${kind === 'header' ? 'hdr' : 'ftr'}>`;
    parts.set(`word/${part}.xml`, toBytes(content.replace(/^<w:([^>]+)>/, '<w:$1 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">')));
    parts.set('word/_rels/document.xml.rels', toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${part}" Target="${part}.xml"/></Relationships>`));
    parts.set('[Content_Types].xml', toBytes(new TextDecoder().decode(parts.get('[Content_Types].xml')!).replace('</Types>',
      `<Override PartName="/word/${part}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${part}+xml"/></Types>`)));
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function worker(session: ResidentEngineSession, controlled = false, virtualTime = false) {
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
  let now = 0;
  let timerId = 0;
  const timers = new Map<number, { at: number; callback: () => void }>();
  const harness = { session, echoBackgroundPage, encodeDisplayListFrameExtras, loadRustDisplayListQueryEngine,
    loadedRustDisplayListQueryEngine,
    ...(virtualTime ? {
      clock: { now: () => now },
      setTimer(callback: () => void, ms: number) {
        const id = ++timerId;
        timers.set(id, { at: now + ms, callback });
        return id;
      },
      clearTimer(id: number) { timers.delete(id); },
    } : {}),
  };
  startWorker(scope, harness, controlled ? class {
    port1 = { onmessage: () => {} };
    port2 = { postMessage: () => turns.push(() => this.port1.onmessage()) };
  } : MessageChannel);
  return {
    frames,
    replaceSession(next: ResidentEngineSession) { harness.session = next; },
    async turn() { turns.shift()?.(); await new Promise<void>((resolve) => setTimeout(resolve, 0)); },
    async drain() { for (let i = 0; i < 2000 && turns.length; i += 1) { turns.shift()!(); await new Promise<void>((resolve) => setTimeout(resolve, 0)); } },
    async advance(ms: number) {
      await this.drain();
      now += ms;
      for (const [id, timer] of timers) {
        if (timer.at > now) continue;
        timers.delete(id);
        timer.callback();
      }
      await this.drain();
    },
    send(request: ResidentEngineWorkerRequestWithoutId): Promise<ResidentEngineWorkerResponse> {
      const id = ++nextId;
      return new Promise((resolve) => {
        replies.set(id, resolve);
        scope.onmessage({ data: { ...request, id } as ResidentEngineWorkerRequest });
      });
    },
    /** Sends `request` while a completion step's turn is queued; that turn runs the request first. */
    async interleave(request: ResidentEngineWorkerRequestWithoutId): Promise<ResidentEngineWorkerResponse> {
      const reply = this.send(request);
      await this.turn();
      return reply;
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


function featureSection(feature: string): string {
  if (feature === 'columns') return '<w:sectPr><w:cols w:num="2"/></w:sectPr>';
  if (feature.startsWith('header-') || feature.startsWith('footer-')) return `<w:sectPr><w:${feature.split('-')[0]}Reference w:type="default" r:id="rIdPart"/></w:sectPr>`;
  return '<w:sectPr/>';
}

function featureBody(feature: string, body: string): string {
  const field = (type: string) => `<w:p><w:fldSimple w:instr=" ${type} "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>`;
  if (['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'DATE', 'TIME'].includes(feature)) return field(feature) + body;
  if (feature === 'late-NUMPAGES') return body + field('NUMPAGES');
  if (feature === 'complex-NUMPAGES') return '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' + body;
  if (feature === 'nested-NUMPAGES') return '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' + field('NUMPAGES') + '</w:tc></w:tr></w:tbl>' + body;
  if (feature === 'table') return '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    Array.from({ length: 8 }, (_, i) => `<w:tr><w:tc><w:p><w:r><w:t>Row ${i}</w:t></w:r></w:p></w:tc></w:tr>`).join('') + '</w:tbl>' + body;
  if (feature === 'keep') {
    let kept = body;
    for (const index of [2, 3, 31, 32, 33]) {
      const id = (index + 1).toString(16).padStart(8, '0');
      kept = kept.replace(`<w:p w14:paraId="${id}"><w:pPr><w:pageBreakBefore/>`,
        `<w:p w14:paraId="${id}"><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/><w:keepNext/><w:keepLines/>`);
    }
    return kept;
  }
  if (feature === 'continuous-section') return featureBody('section', body).replace('<w:pgSz w:w="9000"', '<w:type w:val="continuous"/><w:pgSz w:w="9000"');
  if (feature === 'section') return body.replace(
    '<w:p w14:paraId="00000021">', '<w:p><w:pPr><w:sectPr><w:pgSz w:w="9000" w:h="15840"/></w:sectPr></w:pPr></w:p><w:p w14:paraId="00000021">');
  if (feature === 'float') return '<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="margin" w:horzAnchor="margin" w:tblpX="0" w:tblpY="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Float</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' + body;
  if (feature === 'footnote' || feature === 'endnote') return `<w:p><w:r><w:${feature}Reference w:id="1"/></w:r></w:p>` + body;
  return body;
}

const FLOW_LINES = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>';
const PNG = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='),
  (char) => char.charCodeAt(0)
);

/** A paragraph of `lines` 16px lines, each `label` and its line number. */
function flowParagraph(lines: number, label: string, properties = '', section = ''): string {
  const text = Array.from({ length: lines }, (_, line) => `<w:t xml:space="preserve">${label} ${line + 1}</w:t>`).join('<w:br/>');
  return `<w:p><w:pPr>${properties}${FLOW_LINES}${section}</w:pPr><w:r>${text}</w:r></w:p>`;
}

/** A Letter section with 53 body lines a page, or 38 in landscape. */
function flowSection(options: { type?: string; references?: string; titlePage?: boolean; landscape?: boolean } = {}): string {
  const size = options.landscape
    ? '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'
    : '<w:pgSz w:w="12240" w:h="15840"/>';
  return `<w:sectPr>${options.references ?? ''}${options.type ? `<w:type w:val="${options.type}"/>` : ''}${size}` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1500" w:left="1440" w:header="200" w:footer="720" w:gutter="0"/>' +
    `${options.titlePage ? '<w:titlePg/>' : ''}</w:sectPr>`;
}

/** Documents whose page three ends inside the construct the feature names. */
function flowDocx(feature: string): Uint8Array {
  const third = '<w:pageBreakBefore/>';
  const opening = flowParagraph(1, 'Page one') + flowParagraph(1, 'Page two', third);
  const tail = (label: string, count = 300) =>
    Array.from({ length: count }, (_, index) => flowParagraph(1, `${label} ${index}`)).join('');
  const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  const parts: PartsMap = new Map();
  const relationships: string[] = [];
  const types: string[] = [];
  let body: string;
  switch (feature) {
    case 'flow-selection':
      body = flowParagraph(4, 'alpha beta gamma delta epsilon') + tail('Tail') + flowSection();
      break;
    case 'flow-paragraph':
      body = opening + flowParagraph(120, 'Long', third) + tail('Tail') + flowSection();
      break;
    case 'flow-table':
      body = opening + flowParagraph(1, 'Table follows', third) +
        '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders><w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
        `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(flowParagraph(1, 'Head A'))}${cell(flowParagraph(1, 'Head B'))}</w:tr>` +
        `<w:tr>${cell(flowParagraph(80, 'Split'))}${cell(flowParagraph(2, 'Beside'))}</w:tr>` +
        Array.from({ length: 10 }, (_, row) => `<w:tr>${cell(flowParagraph(1, `Row ${row}`))}${cell(flowParagraph(1, 'Cell'))}</w:tr>`).join('') +
        '</w:tbl>' + tail('Tail') + flowSection();
      break;
    case 'flow-widow':
      body = opening + flowParagraph(49, 'Filler', third) + flowParagraph(5, 'Widow', '<w:widowControl/>') + tail('Tail') + flowSection();
      break;
    case 'flow-keep':
      body = opening + flowParagraph(46, 'Filler', third) +
        Array.from({ length: 10 }, (_, index) => flowParagraph(1, `Chain ${index}`, '<w:keepNext/>')).join('') +
        flowParagraph(1, 'Chain end') + tail('Tail') + flowSection();
      break;
    case 'flow-inline-image':
    case 'flow-image':
      types.push('<Default Extension="png" ContentType="image/png"/>');
      relationships.push('<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>');
      parts.set('word/media/image1.png', PNG);
      body = (feature === 'flow-inline-image' ? flowParagraph(1, 'Before image') : opening + flowParagraph(48, 'Filler', third)) +
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="1905000" cy="1428750"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
        '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1428750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' + tail('Tail') + flowSection();
      break;
    case 'flow-oddPage':
      body = opening + flowParagraph(20, 'Third', third, flowSection()) +
        tail('Odd') + flowParagraph(1, 'Odd end', '', flowSection({ type: 'oddPage' })) + tail('Tail') + flowSection();
      break;
    case 'flow-evenPage':
    case 'flow-evenPage-bare':
      body = flowParagraph(1, 'Page one') + flowParagraph(1, 'Page two', third, flowSection()) +
        tail('Even') + flowParagraph(1, 'Even end', '', flowSection({ type: 'evenPage' })) + tail('Tail') + flowSection();
      break;
    case 'flow-landscape-bare':
      body = opening + tail('Wide', 400) +
        flowParagraph(1, 'Wide end', '', flowSection({ landscape: true })) + tail('Tail') + flowSection();
      break;
    case 'flow-titlePg': {
      const header = (lines: number, label: string) =>
        `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${flowParagraph(lines, label)}</w:hdr>`;
      parts.set('word/header1.xml', toBytes(header(6, 'First header')));
      parts.set('word/header2.xml', toBytes(header(1, 'Header')));
      for (const index of [1, 2]) {
        relationships.push(`<Relationship Id="rIdHeader${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header${index}.xml"/>`);
        types.push(`<Override PartName="/word/header${index}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`);
      }
      body = Array.from({ length: 60 }, (_, index) => flowParagraph(10, `Flow ${index}`)).join('') + flowSection({
        titlePage: true,
        references: '<w:headerReference w:type="first" r:id="rIdHeader1"/><w:headerReference w:type="default" r:id="rIdHeader2"/>',
      });
      break;
    }
    default:
      throw new Error(`Unknown flow feature ${feature}`);
  }
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    `<Default Extension="xml" ContentType="application/xml"/>${types.join('')}` +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>'
  ));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>'
  ));
  if (relationships.length) {
    parts.set('word/_rels/document.xml.rels', toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join('')}</Relationships>`
    ));
  }
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<w:body>${body}</w:body></w:document>`
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function pageText(page: { primitives: unknown[] }): string {
  return page.primitives.map((primitive) => (primitive as { text?: string }).text ?? '').join('\n');
}

/** Asserts that page three of `pages` ends inside the construct the feature names. */
const FLOW_STRADDLES: Record<string, (pages: RetainedFrame['displayList']['pages']) => void> = {
  'flow-paragraph': (pages) => {
    expect(pageText(pages[2]!)).toContain('Long');
    expect(pageText(pages[3]!)).toContain('Long');
  },
  'flow-table': (pages) => {
    for (const page of [pages[2]!, pages[3]!]) {
      expect(pageText(page)).toContain('Head A');
      expect(pageText(page)).toContain('Split');
    }
  },
  'flow-widow': (pages) => {
    expect(pageText(pages[2]!)).toContain('Widow');
    expect(pageText(pages[3]!)).toContain('Widow');
  },
  'flow-keep': (pages) => {
    expect(pageText(pages[2]!)).toContain('Filler 46');
    expect(pageText(pages[3]!)).toContain('Chain');
  },
  'flow-image': (pages) => {
    expect(pageText(pages[2]!)).toContain('Filler 48');
    expect(pages[3]!.primitives.some((primitive) => primitive.kind === 'image')).toBe(true);
  },
  'flow-oddPage': (pages) => {
    expect(pageText(pages[2]!)).toContain('Third');
    expect(pageText(pages[3]!).trim()).toBe('');
    expect(pageText(pages[4]!)).toContain('Odd');
  },
  'flow-evenPage': (pages) => {
    expect(pageText(pages[2]!).trim()).toBe('');
    expect(pageText(pages[3]!)).toContain('Even');
  },
  'flow-titlePg': (pages) => {
    const last = pageText(pages[2]!).match(/Flow \d+/g)!.at(-1);
    expect(pageText(pages[3]!).match(/Flow \d+/g)![0]).toBe(last!);
    expect(pageText(pages[0]!.header!)).toContain('First header');
    expect(pageText(pages[2]!.header!)).not.toContain('First header');
  },
};
FLOW_STRADDLES['flow-evenPage-bare'] = FLOW_STRADDLES['flow-evenPage']!;
FLOW_STRADDLES['flow-landscape-bare'] = (pages) => {
  expect(pages[2]!.width).toBeGreaterThan(pages[2]!.height);
  expect(pages.at(-1)!.width).toBeLessThan(pages[2]!.width);
};

/** The pass a plain text edit relayouts through, on engines that keep it apart from the full retained pass. */
const EDIT_LAYOUT = 'layoutDocumentWithRegionsAfterEdit';

const MEASUREMENT = {
  defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
  compat: { noLeading: false, doNotExpandShiftReturn: false }, authoritativeShaping: true as const,
};

/** `exportable` registers the font for the default family too, as an export requires. */
async function commandFixture(feature = 'plain', exportable = false, options: {
  full?: boolean; virtualTime?: boolean; prepare?: (engine: ResidentEngineSession) => void;
} = {}) {
  const bytes = syntheticDocx(feature);
  const cold = await createResidentEngineSession();
  const engine = await createResidentEngineSession();
  const w = worker(engine, true, options.virtualTime);
  const host = decodeDocxHostJson(cold.openDocx(bytes), bytes);
  options.prepare?.(cold);
  const request = buildResidentRegionLayoutRequest(host.document, 24, {});
  if (feature === 'footnote' || feature === 'endnote') {
    request.notes = { contents: [{ id: 1, noteKind: feature, height: 0 }] } as typeof request.notes;
  }
  if (feature.endsWith('-bare')) request.regions.sections = [];
  if (exportable) request.measurement = { ...MEASUREMENT, fontChains: {} };
  const requirements = JSON.parse(cold.layoutFontRequirementsJson(JSON.stringify(request))) as Array<{ key: string }>;
  cold.clearFonts();
  const font = cold.registerFont(FONT);
  const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [font]]));
  request.measurement = { ...MEASUREMENT, fontChains };
  const layoutInput = JSON.stringify(request);
  cold.layoutDocumentWithRegionsRetainedJson(layoutInput);
  cold.setDisplayWindow(0, 3);
  cold.setDisplayRetainBuiltPages(false);
  cold.setWindowedIncrementalBuilds(true);
  const extras = JSON.stringify({ fontChains, headersFooters: JSON.parse(cold.retainedHeadersFootersJson() ?? 'null') });
  const full = applyFrameDelta(null, decodeFrameDelta(cold.buildDisplayListFrame(extras, 0)));
  await w.send({ type: 'open', bytes: bytes.slice().buffer });
  options.prepare?.(engine);
  const editLayout = EDIT_LAYOUT in engine ? EDIT_LAYOUT : 'layoutDocumentWithRegionsRetained';
  const methods = new Set([
    'layoutDocumentWithRegionsRetained', 'layoutDocumentWithRegionsRetainedJson', 'beginRegionLayout', 'resumeRegionLayout', editLayout,
  ]);
  const calls: string[] = [];
  const spied = engine as unknown as Record<string, (...args: never[]) => unknown>;
  for (const method of methods) {
    const original = spied[method].bind(engine);
    spied[method] = (...args: never[]) => { calls.push(method); return original(...args); };
  }
  const init: ResidentCommandInitRequest = {
    session: 'viewport-first', incarnation: 1, epoch: 1, readOnly: false,
    fonts: [FONT], layoutInput, layoutWithRegions: true, layoutRevision: 1, extras,
    displayWindow: [0, 3], ...(options.full ? {} : { provisionalPages: 3 }),
  };
  const reply = await w.send({ type: 'commandInit', request: init });
  if (!reply.ok || !reply.commandInitial) throw new Error(JSON.stringify(reply));
  const initial = reply.commandInitial;
  const frame = applyFrameDelta(null, decodeFrameDelta(initial.frame.payload.frame));
  return { cold, engine, w, init, initial, frame, full, calls, editLayout, acks: [] as DocxCommandFrame[],
    destroy() { engine.destroy(); cold.destroy(); },
    async present() {
      return w.send({ type: 'commandViewport', request: { start: 0, end: 3, frameEpoch: frame.frameEpoch } });
    },
  };
}

for (const feature of ['plain', 'table', 'keep', 'section', 'PAGE', 'header-PAGE', 'footer-PAGE']) {
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
      expect(final.displayList.pages).toEqual(f.full.displayList.pages);
    } finally { f.destroy(); }
  });
}

for (const feature of ['NUMPAGES', 'SECTIONPAGES', 'late-NUMPAGES', 'complex-NUMPAGES', 'nested-NUMPAGES', 'header-NUMPAGES', 'footer-SECTIONPAGES', 'footnote', 'endnote', 'float', 'columns', 'continuous-section', 'DATE', 'TIME']) {
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

for (const feature of Object.keys(FLOW_STRADDLES)) {
  test(`command prefix is exact where page three ends inside ${feature}`, async () => {
    const f = await commandFixture(feature);
    try {
      f.cold.setDisplayWindow(0, 2 ** 32 - 1);
      const whole = applyFrameDelta(f.full, decodeFrameDelta(f.cold.buildDisplayListFrame(f.init.extras, f.full.frameEpoch)));
      FLOW_STRADDLES[feature]!(whole.displayList.pages);
      expect(f.initial.frame.payload.layoutProvisional).toBe(true);
      expect(f.frame.pages).toHaveLength(3);
      for (const { page, pageIndex } of f.frame.pages) expect(page).toEqual(f.full.displayList.pages[pageIndex]);
      expect(f.initial.frame.payload.documentWidth).toBe(Math.max(...whole.displayList.pages.map(({ width }) => width)));
      await f.present();
      await f.w.drain();
      expect(f.w.frames.at(-1)!.payload.layoutProvisional).not.toBe(true);
      expect(commandFramesAfter(f.frame, f.w.frames).displayList.pages).toEqual(f.full.displayList.pages);
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

test('a review read answers for the presented prefix frame while its completion is pending', async () => {
  const f = await commandFixture();
  const review = () => f.w.send({ type: 'documentRead', read: { kind: 'reviewSnapshot', expectPreviewVersion: 0 } });
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    const version = f.engine.proposalEngine.version();
    const prefix = await review();
    expect(prefix.ok && prefix.read).toMatchObject({ version, value: { version, frameEpoch: f.frame.frameEpoch, layoutRevision: 1,
      proposals: { version, proposals: [] } } });
    expect(f.calls).toEqual([]);
    await f.present();
    await f.w.turn();
    expect(f.calls).toContain('beginRegionLayout');
    const presented = f.w.frames.at(-1)!;
    expect(presented.payload.layoutProvisional).toBe(true);
    const pending = await f.w.interleave({ type: 'documentRead', read: { kind: 'reviewSnapshot', expectPreviewVersion: 0 } });
    expect(pending.ok && pending.read?.value).toMatchObject({ frameEpoch: presented.revision });
    await f.w.drain();
    const completed = f.w.frames.at(-1)!;
    expect(completed.payload.layoutProvisional).not.toBe(true);
    const whole = await review();
    expect(whole.ok && whole.read?.value).toMatchObject({ frameEpoch: completed.revision });
  } finally { f.destroy(); }
});

test('command viewport beyond the prefix completes before painting and cancels stale slices', async () => {
  const f = await commandFixture();
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    await f.present();
    await f.w.turn();
    await f.w.interleave({ type: 'commandViewport', request: { start: 40, end: 43, frameEpoch: f.w.frames.at(-1)!.revision } });
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

type CommandFixture = Awaited<ReturnType<typeof commandFixture>>;

function commandPoint(f: CommandFixture, pageIndex = 0, offset = 8): DocxSelectionPointIntent {
  const page = f.full.pages[pageIndex]!;
  const run = page.page.primitives.find((primitive) => primitive.kind === 'glyphRun');
  if (!run || run.kind !== 'glyphRun') throw new Error('Missing text');
  const glyph = run.glyphs[Math.min(offset, run.glyphs.length - 1)]!;
  return {
    page: { pageId: page.pageId.toString(), pageIndex }, x: glyph.x, y: glyph.y - run.size / 2,
    frameEpoch: f.frame.frameEpoch, layoutRevision: 1, clickCount: 1,
    modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  };
}

async function sendCommand(f: CommandFixture, seq: number, kind: string, payload: DocxCommandEnvelope['payload']): Promise<DocxCommandAck> {
  const reply = await f.w.interleave({ type: 'edit', command: {
    session: f.init.session, incarnation: 1, seq, basis: { epoch: 1, version: 0 }, kind, payload,
    ...(seq > 1 ? { afterSeq: seq - 1 } : {}),
  } });
  if (!reply.ok || !reply.commandAck?.frame) throw new Error(JSON.stringify(reply));
  expect(reply.commandAck.outcomes[0]?.status).toBe(kind === 'insertText' ? 'applied' : 'noop');
  f.acks.push(reply.commandAck.frame);
  return reply.commandAck;
}

function selectionPositions(selection: DocxCommandSelection | null | undefined) {
  const position = (at: DocxCommandSelection['focus']) => ({ story: at.story, paraId: at.paraId, offset: at.offset });
  return selection && { anchor: position(selection.anchor), focus: position(selection.focus) };
}

function expectPrefixSelection(f: CommandFixture, ack: DocxCommandAck, calls: string[]) {
  expect(ack.frame!.payload.layoutProvisional).toBe(true);
  expect(ack.frame!.payload.finalizedPages).toBe(3);
  expect(f.calls).toEqual(calls);
  const frame = commandFramesAfter(f.frame, [...f.w.frames, ...f.acks].sort((a, b) => a.revision - b.revision));
  expect(frame.pages).toHaveLength(3);
  for (const { page, pageIndex } of frame.pages) expect(page).toEqual(f.full.displayList.pages[pageIndex]);
}

function expectSelectionParity(ack: DocxCommandAck, full: DocxCommandAck, requiresEcho = true) {
  expect(selectionPositions(ack.selection)).toEqual(selectionPositions(full.selection));
  const echo = ack.frame!.payload.echoContext;
  const completedEcho = full.frame!.payload.echoContext;
  if (!requiresEcho) {
    expect(echo).toEqual(completedEcho);
    return;
  }
  expect(echo).toBeDefined();
  expect(completedEcho).toBeDefined();
  expect(echo!.page).toEqual(completedEcho!.page);
  expect(echo!.text).toBe(completedEcho!.text);
  expect(echo!.selection).toEqual(completedEcho!.selection);
  expect(echo!.boundaries).toEqual(completedEcho!.boundaries);
  expect(echo!.styleRuns).toEqual(completedEcho!.styleRuns);
}

function expectAnsweredSelectionParity(ack: DocxCommandAck, full: DocxCommandAck, completed = true) {
  expect(selectionPositions(ack.selection)).toEqual(selectionPositions(full.selection));
  const payload = ack.frame!.payload;
  const reference = full.frame!.payload;
  expect(payload.caret.caretRect).toEqual(reference.caret.caretRect);
  expect(payload.caret.selection).toEqual(reference.caret.selection);
  expect(payload.selectionRects).toEqual(reference.selectionRects);
  expect(payload.parts).toEqual(reference.parts);
  if (completed) {
    const geometry = (value: typeof payload.geometry) => value && {
      selection: value.selection,
      objects: value.objects.map(({ kind, story, rects }) => ({ kind, story, rects })),
      unbuiltPages: value.unbuiltPages,
    };
    expect(payload.geometry).toBeDefined();
    expect(geometry(payload.geometry)).toEqual(geometry(reference.geometry));
  } else {
    expect(payload.caret.caretRect).not.toBeNull();
    expect(payload.selectionRects).toEqual([]);
    expect(payload.geometry?.selection ?? []).toEqual(reference.geometry!.selection);
  }
}

for (const [clickCount, label] of [[2, 'word'], [3, 'paragraph']] as const) {
  test(`prefix ${label} click answers the full selection frame`, async () => {
    const f = await commandFixture('flow-selection', false, { virtualTime: true });
    const full = await commandFixture('flow-selection', false, { full: true, virtualTime: true });
    try {
      expect(f.initial.frame.payload.layoutProvisional).toBe(true);
      const ack = await sendCommand(f, 1, 'selectPoint', { ...commandPoint(f), clickCount });
      const reference = await sendCommand(full, 1, 'selectPoint', { ...commandPoint(full), clickCount });
      expect(reference.frame!.payload.selectionRects!.length).toBeGreaterThan(0);
      expectAnsweredSelectionParity(ack, reference);
      expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
      expect(f.calls).toContain('layoutDocumentWithRegionsRetained');
    } finally { f.destroy(); full.destroy(); }
  });
}

for (const gesture of ['shift click', 'shift right', 'shift down'] as const) {
  test(`prefix ${gesture} answers the full selection frame`, async () => {
    const f = await commandFixture('flow-selection', false, { virtualTime: true });
    const full = await commandFixture('flow-selection', false, { full: true, virtualTime: true });
    try {
      await sendCommand(f, 1, 'selectPoint', commandPoint(f));
      await sendCommand(full, 1, 'selectPoint', commandPoint(full));
      const kind = gesture === 'shift click' ? 'extendSelection' : 'navigate';
      const point = commandPoint(f, 0, 24);
      const payload = gesture === 'shift click' ? { ...point, modifiers: { ...point.modifiers, shift: true } }
        : gesture === 'shift right' ? { unit: 'character' as const, direction: 'forward' as const, extend: true }
          : { unit: 'line' as const, direction: 'down' as const, extend: true, goalX: point.x };
      const ack = await sendCommand(f, 2, kind, payload);
      const reference = await sendCommand(full, 2, kind, payload);
      expect(reference.frame!.payload.selectionRects!.length).toBeGreaterThan(0);
      expectAnsweredSelectionParity(ack, reference);
      expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
      expect(f.calls).toContain('layoutDocumentWithRegionsRetained');
    } finally { f.destroy(); full.destroy(); }
  });
}

test('prefix mid-line shift click then shift down preserves the full focus column', async () => {
  const f = await commandFixture('flow-selection', false, { virtualTime: true });
  const full = await commandFixture('flow-selection', false, { full: true, virtualTime: true });
  try {
    await sendCommand(f, 1, 'selectPoint', commandPoint(f, 0, 4));
    await sendCommand(full, 1, 'selectPoint', commandPoint(full, 0, 4));
    const shifted = await sendCommand(f, 2, 'extendSelection', commandPoint(f, 0, 24));
    const shiftedFull = await sendCommand(full, 2, 'extendSelection', commandPoint(full, 0, 24));
    const goal = (ack: DocxCommandAck, fixture: CommandFixture) => {
      const payload = ack.frame!.payload;
      const rect = payload.selectionRects?.at(-1);
      return payload.caret.caretRect?.x ?? (rect ? rect.x + rect.width : fixture.full.pages[0]!.page.contentBounds!.x);
    };
    const ack = await sendCommand(f, 3, 'navigate', {
      unit: 'line', direction: 'down', extend: true, goalX: goal(shifted, f),
    });
    const reference = await sendCommand(full, 3, 'navigate', {
      unit: 'line', direction: 'down', extend: true, goalX: goal(shiftedFull, full),
    });
    expect(reference.selection!.focus.paraId).toBe(shiftedFull.selection!.focus.paraId);
    expect(reference.selection!.focus.offset).toBeGreaterThan(shiftedFull.selection!.focus.offset);
    expectAnsweredSelectionParity(ack, reference);
    expectAnsweredSelectionParity(shifted, shiftedFull);
  } finally { f.destroy(); full.destroy(); }
});

test('prefix inline image click answers with full object geometry', async () => {
  const f = await commandFixture('flow-inline-image', false, { virtualTime: true });
  const full = await commandFixture('flow-inline-image', false, { full: true, virtualTime: true });
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    const image = f.full.pages[0]!.page.primitives.find(primitive => primitive.kind === 'image');
    if (!image || image.kind !== 'image') throw new Error('Missing inline image');
    const point = { ...commandPoint(f), x: image.x + image.w / 2, y: image.y + image.h / 2 };
    const ack = await sendCommand(f, 1, 'selectPoint', point);
    const reference = await sendCommand(full, 1, 'selectPoint', point);
    expect(reference.frame!.payload.geometry!.objects.some(object => object.kind === 'image')).toBe(true);
    expectAnsweredSelectionParity(ack, reference);
    expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
    expect(f.calls).toContain('layoutDocumentWithRegionsRetained');
  } finally { f.destroy(); full.destroy(); }
});

test('prefix part click answers with full part geometry', async () => {
  const f = await commandFixture('header-PAGE', false, { virtualTime: true });
  const full = await commandFixture('header-PAGE', false, { full: true, virtualTime: true });
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    const header = f.full.pages[0]!.page.header!;
    const run = header.primitives.find(primitive => primitive.kind === 'glyphRun');
    if (!run || run.kind !== 'glyphRun') throw new Error('Missing header text');
    const point = { ...commandPoint(f), x: run.glyphs[0]!.x, y: run.glyphs[0]!.y - run.size / 2 };
    const ack = await sendCommand(f, 1, 'selectPoint', point);
    const reference = await sendCommand(full, 1, 'selectPoint', point);
    expect(reference.frame!.payload.parts!.length).toBeGreaterThan(0);
    expectAnsweredSelectionParity(ack, reference);
    expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
  } finally { f.destroy(); full.destroy(); }
});

test('prefix queued click waits for the layout after an identity assigning save', async () => {
  const prepare = (engine: ResidentEngineSession) => {
    const paragraph = engine.geometryReader.paragraphs('body').at(-1)!;
    engine.splitParagraph({ story: 'body', paraId: paragraph.paraId, offset: 1 });
  };
  const f = await commandFixture('echo', false, { virtualTime: true, prepare });
  const full = await commandFixture('echo', false, { full: true, virtualTime: true, prepare });
  const native = await loadRustDisplayListQueryEngine();
  const hit = native.hitTestRegionsByHandle;
  if (!hit) throw new Error('Missing retained hit testing');
  const completedBeforeHit: boolean[] = [];
  native.hitTestRegionsByHandle = (...args) => {
    completedBeforeHit.push(f.calls.includes('layoutDocumentWithRegionsRetained'));
    return hit.apply(native, args);
  };
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    const before = f.engine.proposalEngine.version();
    const saved = f.w.send({ type: 'save', comments: [] });
    const clicked = sendCommand(f, 1, 'selectPoint', commandPoint(f));
    const [save, ack] = await Promise.all([saved, clicked]);
    if (!save.ok || !save.saved) throw new Error(JSON.stringify(save));
    expect(save.version).not.toBe(before);
    expect(completedBeforeHit.length).toBeGreaterThan(0);
    expect(completedBeforeHit.every(Boolean)).toBe(true);
    native.hitTestRegionsByHandle = hit;
    const fullSave = full.w.send({ type: 'save', comments: [] });
    const referenceClick = sendCommand(full, 1, 'selectPoint', commandPoint(full));
    const [savedFull, reference] = await Promise.all([fullSave, referenceClick]);
    expect(savedFull.ok).toBe(true);
    expectAnsweredSelectionParity(ack, reference);
    expect(ack.frame!.payload.documentVersion).toBe(save.version);
    expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
  } finally { native.hitTestRegionsByHandle = hit; f.destroy(); full.destroy(); }
});

test('prefix text caret click and character and word moves retain complete caret frames without completion', async () => {
  const f = await commandFixture('echo', false, { virtualTime: true });
  const full = await commandFixture('echo', false, { full: true, virtualTime: true });
  try {
    const ack = await sendCommand(f, 1, 'selectPoint', commandPoint(f));
    const reference = await sendCommand(full, 1, 'selectPoint', commandPoint(full));
    expectPrefixSelection(f, ack, []);
    expectAnsweredSelectionParity(ack, reference, false);
    for (const [index, unit] of (['character', 'word'] as const).entries()) {
      const payload = { unit, direction: 'forward' as const, extend: false };
      const moved = await sendCommand(f, index + 2, 'navigate', payload);
      const movedFull = await sendCommand(full, index + 2, 'navigate', payload);
      expectPrefixSelection(f, moved, []);
      expectAnsweredSelectionParity(moved, movedFull, false);
    }
  } finally { f.destroy(); full.destroy(); }
});

test('selectPoint on a certified page acknowledges the exact prefix with echo context without completion', async () => {
  const f = await commandFixture('echo', false, { virtualTime: true });
  const full = await commandFixture('echo', false, { full: true, virtualTime: true });
  try {
    const ack = await sendCommand(f, 1, 'selectPoint', commandPoint(f));
    const reference = await sendCommand(full, 1, 'selectPoint', commandPoint(full));
    expectPrefixSelection(f, ack, []);
    expectSelectionParity(ack, reference);
  } finally { f.destroy(); full.destroy(); }
});

test('selectPoint rebuilds a released certified page without completing layout', async () => {
  const f = await commandFixture('echo', false, { virtualTime: true });
  const full = await commandFixture('echo', false, { full: true, virtualTime: true });
  try {
    for (const fixture of [f, full]) {
      await fixture.w.send({ type: 'commandViewport', request: { start: 0, end: 1, frameEpoch: fixture.frame.frameEpoch } });
      expect(commandFramesAfter(fixture.frame, fixture.w.frames).pages[1]!.page.unbuilt).toBe(true);
    }
    const ack = await sendCommand(f, 1, 'selectPoint', commandPoint(f, 1));
    const reference = await sendCommand(full, 1, 'selectPoint', commandPoint(full, 1));
    expect(ack.frame!.payload.layoutProvisional).toBe(true);
    expect(ack.frame!.payload.finalizedPages).toBe(3);
    expect(f.calls).toEqual([]);
    const frame = commandFramesAfter(f.frame, [...f.w.frames, ...f.acks].sort((a, b) => a.revision - b.revision));
    expect(frame.pages[1]!.page).toEqual(f.full.pages[1]!.page);
    expectSelectionParity(ack, reference);
  } finally { f.destroy(); full.destroy(); }
});

for (const unit of ['character', 'word', 'line'] as const) {
  for (const direction of ['backward', 'forward'] as const) {
    test(`navigate ${unit} ${direction} on a certified page acknowledges the exact prefix with echo context`, async () => {
      const f = await commandFixture('echo', false, { virtualTime: true });
      const full = await commandFixture('echo', false, { full: true, virtualTime: true });
      try {
        await sendCommand(f, 1, 'selectPoint', commandPoint(f));
        await sendCommand(full, 1, 'selectPoint', commandPoint(full));
        const calls = f.calls.slice();
        const navigation = { unit, direction, extend: false };
        const ack = await sendCommand(f, 2, 'navigate', navigation);
        const reference = await sendCommand(full, 2, 'navigate', navigation);
        expectPrefixSelection(f, ack, calls);
        expectSelectionParity(ack, reference);
      } finally { f.destroy(); full.destroy(); }
    });
  }
}

test('extendSelection on certified pages completes before answering the full selection frame', async () => {
  const f = await commandFixture('flow-selection', false, { virtualTime: true });
  const full = await commandFixture('flow-selection', false, { full: true, virtualTime: true });
  try {
    await sendCommand(f, 1, 'selectPoint', commandPoint(f));
    await sendCommand(full, 1, 'selectPoint', commandPoint(full));
    const ack = await sendCommand(f, 2, 'extendSelection', commandPoint(f, 0, 24));
    const reference = await sendCommand(full, 2, 'extendSelection', commandPoint(full, 0, 24));
    expect(reference.selection!.anchor.offset).not.toBe(reference.selection!.focus.offset);
    expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
    expect(f.calls).toContain('layoutDocumentWithRegionsRetained');
    expectAnsweredSelectionParity(ack, reference);
    expectSelectionParity(ack, reference, false);
  } finally { f.destroy(); full.destroy(); }
});

for (const kind of ['selectPoint', 'extendSelection', 'line down', 'document end'] as const) {
  test(`${kind} beyond the certified prefix completes before acknowledging selection`, async () => {
    const f = await commandFixture('echo', false, { virtualTime: true });
    const full = await commandFixture('echo', false, { full: true, virtualTime: true });
    try {
      await sendCommand(f, 1, 'selectPoint', commandPoint(f, 2));
      await sendCommand(full, 1, 'selectPoint', commandPoint(full, 2));
      const commandKind = kind === 'line down' || kind === 'document end' ? 'navigate' : kind;
      const payload = kind === 'line down' ? { unit: 'line' as const, direction: 'down' as const, extend: false, goalX: commandPoint(f, 2).x }
        : kind === 'document end' ? { unit: 'document-boundary' as const, direction: 'forward' as const, extend: false }
        : { ...commandPoint(f), page: { pageId: '0', pageIndex: 3 } };
      const ack = await sendCommand(f, 2, commandKind, payload);
      const reference = await sendCommand(full, 2, commandKind, payload);
      expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
      expect(commandFramesAfter(f.frame, f.acks).pages).toHaveLength(f.full.pages.length);
      expect(f.calls).toContain('layoutDocumentWithRegionsRetained');
      expect(selectionPositions(ack.selection)).toEqual(selectionPositions(reference.selection));
    } finally { f.destroy(); full.destroy(); }
  });
}

test('a computed paragraph boundary outside the prefix completes and recomputes selection', async () => {
  const f = await commandFixture('flow-paragraph', false, { virtualTime: true });
  const full = await commandFixture('flow-paragraph', false, { full: true, virtualTime: true });
  try {
    await sendCommand(f, 1, 'selectPoint', commandPoint(f, 2));
    await sendCommand(full, 1, 'selectPoint', commandPoint(full, 2));
    const navigation = { unit: 'paragraph' as const, direction: 'forward' as const, extend: true };
    const ack = await sendCommand(f, 2, 'navigate', navigation);
    const reference = await sendCommand(full, 2, 'navigate', navigation);
    expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
    expect(f.calls).toContain('layoutDocumentWithRegionsRetained');
    expect(selectionPositions(ack.selection)).toEqual(selectionPositions(reference.selection));
  } finally { f.destroy(); full.destroy(); }
});

for (const unit of ['character', 'word'] as const) {
  test(`a computed ${unit} result beyond the prefix completes and recomputes`, async () => {
    const f = await commandFixture('echo', false, { virtualTime: true });
    const full = await commandFixture('echo', false, { full: true, virtualTime: true });
    try {
      const selected = await sendCommand(f, 1, 'selectPoint', commandPoint(f, 2));
      await sendCommand(full, 1, 'selectPoint', commandPoint(full, 2));
      expectPrefixSelection(f, selected, []);
      const end = { unit: 'line' as const, direction: 'forward' as const, extend: false };
      const boundary = await sendCommand(f, 2, 'navigate', end);
      await sendCommand(full, 2, 'navigate', end);
      expectPrefixSelection(f, boundary, []);
      const navigation = { unit, direction: 'forward' as const, extend: true };
      const ack = await sendCommand(f, 3, 'navigate', navigation);
      const reference = await sendCommand(full, 3, 'navigate', navigation);
      expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
      expect(f.calls).toContain('layoutDocumentWithRegionsRetained');
      expect(selectionPositions(ack.selection)).toEqual(selectionPositions(reference.selection));
    } finally { f.destroy(); full.destroy(); }
  });
}

test('a triple click whose paragraph crosses the prefix completes and recomputes', async () => {
  const f = await commandFixture('flow-paragraph', false, { virtualTime: true });
  const full = await commandFixture('flow-paragraph', false, { full: true, virtualTime: true });
  try {
    const ack = await sendCommand(f, 1, 'selectPoint', { ...commandPoint(f, 2), clickCount: 3 });
    const reference = await sendCommand(full, 1, 'selectPoint', { ...commandPoint(full, 2), clickCount: 3 });
    expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
    expect(f.calls).toContain('layoutDocumentWithRegionsRetained');
    expect(selectionPositions(ack.selection)).toEqual(selectionPositions(reference.selection));
  } finally { f.destroy(); full.destroy(); }
});

test('insertText after a prefix selection completes before editing and matches a full open', async () => {
  const f = await commandFixture('echo', false, { virtualTime: true });
  const full = await commandFixture('echo', false, { full: true, virtualTime: true });
  try {
    const selected = await sendCommand(f, 1, 'selectPoint', commandPoint(f));
    const selectedFull = await sendCommand(full, 1, 'selectPoint', commandPoint(full));
    expectPrefixSelection(f, selected, []);
    const ack = await sendCommand(f, 2, 'insertText', { text: 'Edit ' });
    const reference = await sendCommand(full, 2, 'insertText', { text: 'Edit ' });
    expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
    expect(f.calls.filter(call => call === 'layoutDocumentWithRegionsRetained' || call === f.editLayout))
      .toEqual(['layoutDocumentWithRegionsRetained', f.editLayout]);
    const document = f.engine.proposalEngine.readParagraphs({ story: 'body', view: 'accepted' });
    const referenceDocument = full.engine.proposalEngine.readParagraphs({ story: 'body', view: 'accepted' });
    expect(document.ok).toBe(true);
    expect(referenceDocument.ok).toBe(true);
    expect(document.ok && document.paragraphs).toEqual(referenceDocument.ok && referenceDocument.paragraphs);
    const frame = commandFramesAfter(f.frame, [selected.frame!, ack.frame!]);
    const referenceFrame = commandFramesAfter(full.frame, [selectedFull.frame!, reference.frame!]);
    expect(frame.displayList.pages).toEqual(referenceFrame.displayList.pages);
    expectSelectionParity(ack, reference);
  } finally { f.destroy(); full.destroy(); }
});

test('background completion resumes after a prefix selection without restarting', async () => {
  const f = await commandFixture('echo', false, { virtualTime: true });
  try {
    await f.present();
    await f.w.turn();
    expect(f.calls).toEqual(['beginRegionLayout']);
    const ack = await sendCommand(f, 1, 'selectPoint', commandPoint(f));
    expectPrefixSelection(f, ack, ['beginRegionLayout']);
    await f.w.advance(INPUT_IDLE_HOLD_MS);
    expect(f.calls.filter(call => call === 'beginRegionLayout')).toHaveLength(1);
    expect(f.calls).not.toContain('layoutDocumentWithRegionsRetained');
    expect(f.calls.filter(call => call === 'resumeRegionLayout').length).toBeGreaterThan(1);
    expect(f.w.frames.at(-1)!.payload.layoutProvisional).not.toBe(true);
    const frames = [...f.w.frames, ack.frame!].sort((a, b) => a.revision - b.revision);
    const completed = commandFramesAfter(f.frame, frames);
    expect(completed.pages).toHaveLength(f.full.pages.length);
    expect(completed.displayList.pages).toEqual(f.full.displayList.pages);
  } finally { f.destroy(); }
});

test('an edit overtakes sliced completion and finishes it before changing the document', async () => {
  const f = await commandFixture();
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    await f.present();
    await f.w.turn();
    const point = { story: 'body', paraId: f.engine.geometryReader.paragraphs('body')[0]!.paraId, offset: 0 };
    const edited = await f.w.interleave({ type: 'edit', command: {
      session: f.init.session, incarnation: 1, seq: 1, basis: { epoch: 1, version: 0 },
      kind: 'insertText', payload: { text: 'Edit ', selection: { anchor: point, focus: point } },
    } });
    expect(edited.ok && edited.commandAck?.outcomes[0]?.status).toBe('applied');
    expect(edited.ok && edited.commandAck?.frame?.payload.layoutProvisional).not.toBe(true);
    if (!edited.ok || !edited.commandAck?.frame) throw new Error('Missing edit frame');
    const reference = await createResidentEngineSession();
    try {
      reference.loadState(f.engine.encodeState());
      reference.registerFont(FONT);
      reference.layoutDocumentWithRegionsRetained(f.init.layoutInput);
      reference.setDisplayWindow(0, 3);
      reference.setDisplayRetainBuiltPages(false);
      reference.setWindowedIncrementalBuilds(true);
      const cold = applyFrameDelta(null, decodeFrameDelta(reference.buildDisplayListFrame(f.init.extras, 0)));
      const final = commandFramesAfter(commandFramesAfter(f.frame, f.w.frames), [edited.commandAck.frame]);
      expect(final.displayList.pages).toEqual(cold.displayList.pages);
    } finally { reference.destroy(); }
    const count = f.w.frames.length;
    await f.w.drain();
    expect(f.w.frames).toHaveLength(count);
  } finally { f.destroy(); }
});

/** Each engine mints its own revision ids; numbers them by first appearance so two engines' pages compare. */
function revisionOrdinals<T>(value: T): T {
  const ids = new Map<unknown, number>();
  return JSON.parse(JSON.stringify(value, (key, item: unknown) => {
    if (key !== 'revisionId') return item;
    if (!ids.has(item)) ids.set(item, ids.size);
    return ids.get(item);
  })) as T;
}

async function proposeFirstParagraph(f: CommandFixture): Promise<DocxCommandAck> {
  const paragraph = f.engine.paragraphIdentities().paragraphs[0]!.session!;
  const reply = await f.w.interleave({ type: 'edit', command: {
    session: f.init.session, incarnation: 1, seq: 1, basis: { epoch: 1, version: 0 }, kind: 'propose',
    payload: { expectVersion: f.engine.proposalEngine.version(), proposals: [{ id: 'one', paragraph, op: 'replaceText',
      search: 'Paragraph 1.', replaceWith: 'Omega 1.', suggest: { author: 'Ada', date: '2026-01-01' } }] },
  } });
  if (!reply.ok || !reply.commandAck?.frame) throw new Error(JSON.stringify(reply));
  return reply.commandAck;
}

for (const begun of [false, true]) {
  test(`a proposal round on the prefix${begun ? ' while its completion runs' : ''} lays out the proposed document once`, async () => {
    const f = await commandFixture('plain', false, { virtualTime: true });
    const full = await commandFixture('plain', false, { full: true, virtualTime: true });
    try {
      expect(f.initial.frame.payload.layoutProvisional).toBe(true);
      if (begun) {
        await f.present();
        await f.w.turn();
        await full.present();
        expect(f.calls).toEqual(['beginRegionLayout']);
      }
      const before = f.calls.length;
      const ack = await proposeFirstParagraph(f);
      const reference = await proposeFirstParagraph(full);
      expect(ack.outcomes[0]?.status).toBe('applied');
      expect(f.calls.slice(before)).toEqual([f.editLayout]);
      expect(ack.frame!.payload.layoutProvisional).not.toBe(true);
      const frames = [...f.w.frames, ack.frame!].sort((a, b) => a.revision - b.revision);
      const referenceFrames = [...full.w.frames, reference.frame!].sort((a, b) => a.revision - b.revision);
      const proposed = commandFramesAfter(f.frame, frames).displayList.pages;
      const expected = commandFramesAfter(full.frame, referenceFrames).displayList.pages;
      expect(proposed).toHaveLength(expected.length);
      expect(revisionOrdinals(proposed.slice(0, 3))).toEqual(revisionOrdinals(expected.slice(0, 3)));
      await f.w.advance(INPUT_IDLE_HOLD_MS);
      expect(f.calls.slice(before)).toEqual([f.editLayout]);
    } finally { f.destroy(); full.destroy(); }
  });
}

/** Longer than the worker holds background work after user input. */
const INPUT_IDLE_HOLD_MS = 400;

test('a command refused while the prefix is pending leaves it to reach the completed frame', async () => {
  for (const stale of ['epoch', 'incarnation', 'session', 'version'] as const) {
    const f = await commandFixture();
    try {
      expect(f.initial.frame.payload.layoutProvisional).toBe(true);
      await f.present();
      await f.w.turn();
      const point = { story: 'body', paraId: f.engine.geometryReader.paragraphs('body')[0]!.paraId, offset: 0 };
      const refused = await f.w.interleave({ type: 'edit', command: {
        session: stale === 'session' ? 'another' : f.init.session,
        incarnation: stale === 'incarnation' ? 2 : 1, seq: 1,
        basis: { epoch: stale === 'epoch' ? 0 : 1, version: stale === 'version' ? 7 : 0 },
        ...(stale === 'version' ? { expectVersion: true } : {}),
        kind: 'insertText', payload: { text: 'Edit ', selection: { anchor: point, focus: point } },
      } });
      if (!refused.ok || !refused.commandAck) throw new Error(JSON.stringify(refused));
      expect(refused.commandAck.outcomes[0]?.status).toBe('refused');
      await f.w.drain();
      await new Promise((resolve) => setTimeout(resolve, INPUT_IDLE_HOLD_MS));
      await f.w.drain();
      const frames = [...f.w.frames, ...(refused.commandAck.frame ? [refused.commandAck.frame] : [])]
        .sort((left, right) => left.revision - right.revision);
      expect(frames.at(-1)!.payload.layoutProvisional).not.toBe(true);
      expect(commandFramesAfter(f.frame, frames).displayList.pages).toEqual(f.full.displayList.pages);
    } finally { f.destroy(); }
  }
});

const EXPORT_OPTIONS: DocxPageExportOptions = { revisionView: 'markup', stories: ['body'] };

for (const request of ['export', 'completeLayout', 'sliced completeLayout'] as const) {
  test(`${request} finishes a pending prefix inside its own frame work and publishes the completed frame`, async () => {
    const f = await commandFixture('plain', request === 'export');
    try {
      expect(f.initial.frame.payload.layoutProvisional).toBe(true);
      await f.present();
      await f.w.turn();
      const count = f.w.frames.length;
      const reply = request === 'export'
        ? await f.w.interleave({ type: 'documentRead', read: {
          kind: 'exportStructuredWithPages', options: EXPORT_OPTIONS,
          currentRequest: f.init.layoutInput,
        } })
        : await f.w.interleave({
          type: 'completeLayout', expectedFrameEpoch: f.w.frames.at(-1)!.revision, paintCaret: false,
          ...(request === 'sliced completeLayout' ? { sliceBlocks: 8 } : {}),
        });
      if (!reply.ok) throw new Error(JSON.stringify(reply));
      if (request === 'export') {
        type Exported = { ok: boolean; content?: { layout: { pages: unknown[]; fragments: DocxPageFragment[] } } };
        const placed = ({ content }: Exported) => ({
          pages: content?.layout.pages,
          fragments: content?.layout.fragments.map(({ pageIndex, nodeId, blockId, slice, continuedFromPrevious,
            continuedOnNext, repeatedTableHeader }) =>
            ({ pageIndex, nodeId, blockId, slice, continuedFromPrevious, continuedOnNext, repeatedTableHeader })),
        });
        const exported = JSON.parse((reply as { read?: { value?: string } }).read!.value!) as Exported;
        f.cold.layoutDocumentWithRegionsRetainedJson(f.init.layoutInput);
        const cold = JSON.parse(f.cold.exportStructuredWithPagesJson(EXPORT_OPTIONS, f.init.layoutInput)) as Exported;
        expect(exported).toMatchObject({ ok: true });
        expect(cold).toMatchObject({ ok: true });
        expect(placed(exported).pages).toHaveLength(f.full.displayList.pages.length);
        expect(placed(exported).fragments!.length).toBeGreaterThan(f.full.displayList.pages.length);
        expect(placed(exported)).toEqual(placed(cold));
      }
      await f.w.drain();
      expect(f.w.frames.length).toBeGreaterThan(count);
      expect(f.w.frames.at(-1)!.payload.layoutProvisional).not.toBe(true);
      expect(commandFramesAfter(f.frame, f.w.frames).displayList.pages).toEqual(f.full.displayList.pages);
    } finally { f.destroy(); }
  });
}

test('a print completion publishes the whole layout before the next page build and keeps windowed builds', async () => {
  const f = await commandFixture();
  try {
    expect(f.initial.frame.payload.layoutProvisional).toBe(true);
    const windowed: boolean[] = [];
    const setWindowed = f.engine.setWindowedIncrementalBuilds.bind(f.engine);
    f.engine.setWindowedIncrementalBuilds = (enabled: boolean) => { windowed.push(enabled); setWindowed(enabled); };
    await f.present();
    await f.w.turn();
    const completed = await f.w.interleave({
      type: 'completeLayout', expectedFrameEpoch: f.w.frames.at(-1)!.revision, paintCaret: false,
    });
    expect(completed.ok).toBe(true);
    const built = await f.w.interleave({
      type: 'buildPages', pages: [0], expectedFrameEpoch: f.w.frames.at(-1)!.revision, paintCaret: false,
    });
    expect(built.ok).toBe(true);
    expect(f.w.frames.at(-2)!.payload.layoutProvisional).not.toBe(true);
    expect(f.w.frames.at(-1)!.payload.layoutProvisional).not.toBe(true);
    expect(commandFramesAfter(f.frame, f.w.frames).displayList.pages).toEqual(f.full.displayList.pages);
    expect(windowed).not.toContain(false);
  } finally { f.destroy(); }
});

test('collaboration and replacement cannot publish an abandoned command completion', async () => {
  for (const replace of [false, true]) {
    const f = await commandFixture();
    try {
      expect(f.initial.frame.payload.layoutProvisional).toBe(true);
      await f.present();
      await f.w.turn();
      if (replace) {
        const replacement = await createResidentEngineSession();
        f.w.replaceSession(replacement);
        const reply = await f.w.interleave({ type: 'bootstrap', snapshot: {
          clientId: 1, state: f.cold.encodeState(), fontsRevision: 1, fonts: [FONT], renderInputs: [], measureInputs: [],
          layoutInput: f.init.layoutInput, layoutWithRegions: true, layoutRevision: 1, selection: null,
        }, extras: f.init.extras, expectedFrameEpoch: 0, displayWindow: [0, 3] });
        expect(reply.ok && reply.commandInitial?.epoch).toBe(2);
        expect(reply.ok && reply.commandInitial?.frame.payload.layoutProvisional).toBe(true);
        replacement.destroy();
      } else {
        const reply = await f.w.interleave({ type: 'commandUpdate', update: f.engine.encodeState() });
        expect(reply.ok).toBe(true);
        expect(f.w.frames.at(-1)!.payload.layoutProvisional).not.toBe(true);
      }
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
