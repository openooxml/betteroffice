import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { decodeDocxHostJson, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { residentWorkerFactory } from '../../../../../docx/src/yrs/__fixtures__/residentWorker';
import { createPeerlessShellController } from './shellController';
import { createPeerlessViewStore } from './viewStore';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

const ROOT = resolve(import.meta.dir, '../../../../../..');
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))));
  startWorker = await residentWorkerFactory();
});

function documentBytes() {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', toBytes(`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>First page</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second page</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)],
  ])));
}

async function fixture(readOnly = false) {
  const port = startWorker();
  const client = new ResidentEngineWorkerClient(port);
  const bytes = documentBytes();
  const opened = await client.open(bytes);
  const resident = port.sessions[0]!;
  const document = decodeDocxHostJson(opened.hostJson, bytes).document;
  const font = resident.registerFont(new Uint8Array(readFileSync(resolve(ROOT, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))));
  const layout = buildResidentRegionLayoutRequest(document, 24, {});
  const requirements = JSON.parse(resident.layoutFontRequirementsJson(JSON.stringify(layout))) as { key: string }[];
  layout.measurement = { fontChains: Object.fromEntries(requirements.map(({ key }) => [key, [font]])),
    defaults: { fontSize: 11, fontFamily: 'Calibri' }, compat: { noLeading: false, doNotExpandShiftReturn: false }, authoritativeShaping: true };
  const request = JSON.stringify(layout);
  const relayout = () => resident.layoutDocumentWithRegionsRetainedJson(request);
  relayout();
  const viewStore = createPeerlessViewStore({ identity: { session: 'worker', incarnation: 1, epoch: 1 },
    version: 0, throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null, geometry: null, editState: undefined });
  const binding = { viewStore, admission: {}, session: { flush: async () => {} } } as unknown as PeerlessEditSessionBinding;
  const settings = { read: client.documentRead.bind(client), layoutRequest: () => request, zoom: 1, readOnly, focus() {} };
  const controller = createPeerlessShellController(settings);
  controller.attach(binding);
  return { controller, binding, resident, relayout, request, settings, viewStore,
    close: async () => { controller.dispose(); await client.dispose(); } };
}

for (const readOnly of [false, true]) {
  test(`${readOnly ? 'viewer' : 'editor'} ref exports two synthetic pages through the worker read`, async () => {
    const h = await fixture(readOnly);
    try {
      const result = await h.controller.ref.exportStructuredWithPages({ revisionView: 'markup', stories: ['body'], includeGeometry: true });
      expect(result.ok).toBe(true);
      if (!result.ok) throw new Error(result.failure.message);
      expect(result.content.layout.pages).toHaveLength(2);
      expect(JSON.stringify(result.content.structured)).toContain('Second page');
      expect(() => h.controller.ref.getPageContent(1)).toThrow('getPageContent');
      const pinned = await h.controller.ref.exportStructuredWithPages({ revisionView: 'markup', stories: ['body'], expectLayoutVersion: 'missing' });
      expect(pinned).toMatchObject({ ok: false, failure: { code: 'stale-layout' } });
    } finally { await h.close(); }
  });
}

test('paged worker export flushes an edit pending at the ref call', async () => {
  const h = await fixture();
  let release!: () => void;
  const pending = new Promise<void>(resolve => { release = resolve; }).then(() => {
    const paragraph = h.resident.paragraphs('body')[0]!;
    const result = h.resident.proposalEngine.applyEdits({ expectVersion: h.resident.proposalEngine.version(), steps: [
      { op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: paragraph.paraId }, at: 'end', text: ' Pending edit' },
    ] });
    expect(result.ok).toBe(true);
    h.relayout();
  });
  h.binding.session.flush = () => pending;
  try {
    const result = h.controller.ref.exportStructuredWithPages({ revisionView: 'markup', stories: ['body'] });
    release();
    const exported = await result;
    expect(exported.ok).toBe(true);
    if (!exported.ok) throw new Error(exported.failure.message);
    expect(JSON.stringify(exported.content.structured)).toContain('Pending edit');
    expect(exported.content.layout.pages).toHaveLength(2);
  } finally { release(); await pending; await h.close(); }
});

test('paged worker export waits for the layout already in progress', async () => {
  const h = await fixture();
  let request: string | null = null;
  let subscribed!: () => void;
  const listening = new Promise<void>(resolve => { subscribed = resolve; });
  const subscribe = h.viewStore.subscribe;
  h.viewStore.subscribe = listener => { const stop = subscribe(listener); subscribed(); return stop; };
  h.controller.update({ ...h.settings, layoutRequest: () => request });
  try {
    const pending = h.controller.ref.exportStructuredWithPages({ revisionView: 'markup', stories: ['body'] });
    await listening;
    request = h.request;
    h.viewStore.installFrameGeometry({ identity: h.viewStore.getSnapshot().identity, coversThroughSeq: 0,
      frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [] });
    const result = await pending;
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error(result.failure.message);
    expect(result.content.layout.pages).toHaveLength(2);
  } finally { await h.close(); }
});
