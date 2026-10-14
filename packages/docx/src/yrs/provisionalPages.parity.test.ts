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
    '<w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>' +
    `Paragraph ${index + 1}. ${'alpha beta gamma delta epsilon zeta eta theta '.repeat(8 + index % 5)}` +
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
  const harness = { session, encodeDisplayListFrameExtras, loadRustDisplayListQueryEngine, loadedRustDisplayListQueryEngine };
  startWorker(scope, harness, controlled ? class {
    port1 = { onmessage: () => {} };
    port2 = { postMessage: () => turns.push(() => this.port1.onmessage()) };
  } : MessageChannel);
  return {
    frames,
    replaceSession(next: ResidentEngineSession) { harness.session = next; },
    async turn() { turns.shift()?.(); await new Promise<void>((resolve) => setTimeout(resolve, 0)); },
    async drain() { for (let i = 0; i < 2000 && turns.length; i += 1) { turns.shift()!(); await new Promise<void>((resolve) => setTimeout(resolve, 0)); } },
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
    case 'flow-image':
      types.push('<Default Extension="png" ContentType="image/png"/>');
      relationships.push('<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>');
      parts.set('word/media/image1.png', PNG);
      body = opening + flowParagraph(48, 'Filler', third) +
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

const MEASUREMENT = {
  defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
  compat: { noLeading: false, doNotExpandShiftReturn: false }, authoritativeShaping: true as const,
};

/** `exportable` registers the font for the default family too, as an export requires. */
async function commandFixture(feature = 'plain', exportable = false) {
  const bytes = syntheticDocx(feature);
  const cold = await createResidentEngineSession();
  const engine = await createResidentEngineSession();
  const w = worker(engine, true);
  const host = decodeDocxHostJson(cold.openDocx(bytes), bytes);
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
