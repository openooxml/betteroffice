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
import { admitIntent, createInputAdmission } from './inputAdmission';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

const ROOT = resolve(import.meta.dir, '../../../../../..');
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))));
  startWorker = await residentWorkerFactory();
});

function documentBytes(revisions: boolean) {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', toBytes(`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>First page</w:t></w:r>${revisions ? '<w:ins w:id="1" w:author="Writer" w:date="2026-01-01T00:00:00Z"><w:r><w:t> Revised text</w:t></w:r></w:ins>' : ''}</w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second page</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)],
  ])));
}

async function fixture(readOnly = false, commandFonts = false, revisions = false) {
  const port = startWorker();
  const client = new ResidentEngineWorkerClient(port);
  const bytes = documentBytes(revisions);
  const opened = await client.open(bytes);
  const resident = port.sessions[0]!;
  const document = decodeDocxHostJson(opened.hostJson, bytes).document;
  const fontBytes = new Uint8Array(readFileSync(resolve(ROOT, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')));
  const font = resident.registerFont(fontBytes);
  const layout = buildResidentRegionLayoutRequest(document, 24, {});
  const requirements = JSON.parse(resident.layoutFontRequirementsJson(JSON.stringify(layout))) as { key: string }[];
  layout.measurement = { fontChains: Object.fromEntries(requirements.map(({ key }) => [key, [font]])),
    defaults: { fontSize: 11, fontFamily: 'Calibri' }, compat: { noLeading: false, doNotExpandShiftReturn: false }, authoritativeShaping: true };
  const request = JSON.stringify(layout);
  const relayout = () => resident.layoutDocumentWithRegionsRetainedJson(request);
  relayout();
  const baseExport = JSON.parse(resident.exportStructuredWithPagesJson({ revisionView: 'markup', includeGeometry: true }, request));
  const aliasBytes = new Uint8Array(readFileSync(resolve(ROOT, 'crates/docx-raster/tests/assets/Carlito-Regular.ttf')));
  if (commandFonts) await client.commandInit({
    session: 'page-export', incarnation: 1, epoch: 1, readOnly,
    fonts: [fontBytes, aliasBytes, { substituteOf: 1, family: 'MS Mincho' }],
    fontRegistrationSlots: { fontChains: Object.fromEntries(requirements.map(({ key }) => [key, [2]])) },
    layoutInput: request, layoutWithRegions: true, layoutRevision: 1, extras: '{}', provisionalPages: 3, displayWindow: [0, 3],
  });
  const viewStore = createPeerlessViewStore({ identity: { session: 'worker', incarnation: 1, epoch: 1 },
    version: 0, throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null, geometry: null, editState: undefined });
  if (commandFonts) viewStore.installFrameGeometry({ identity: viewStore.getSnapshot().identity, coversThroughSeq: 0,
    frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [] });
  const binding = { viewStore, admission: {}, session: { flush: async () => {} } } as unknown as PeerlessEditSessionBinding;
  const settings = { read: client.documentRead.bind(client), layoutRequest: () => request, zoom: 1, readOnly, focus() {} };
  const controller = createPeerlessShellController(settings);
  controller.attach(binding);
  return { controller, binding, resident, relayout, request, settings, viewStore, client, baseExport,
    close: async () => { controller.dispose(); client.destroy(); } };
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

test('paged worker export uses worker font registrations while still checking the render environment', async () => {
  const h = await fixture(true, true);
  const options = { revisionView: 'markup' as const, stories: ['body' as const], includeGeometry: true };
  try {
    const strict = JSON.parse((await h.client.documentRead({ kind: 'exportStructuredWithPages', options, currentRequest: h.request })).value);
    expect(strict).toMatchObject({ ok: false, failure: { code: 'stale-layout' } });
    const result = await h.controller.ref.exportStructuredWithPages(options);
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error(result.failure.message);
    expect(result.content.layout.pages).toHaveLength(2);
    expect(result.content.layout.fragments.map(fragment => fragment.geometry)).not.toEqual(
      h.baseExport.content.layout.fragments.map((fragment: { geometry: unknown }) => fragment.geometry)
    );
    expect(result.content.layout.layoutVersion).not.toBe(h.baseExport.content.layout.layoutVersion);
    const input = JSON.parse(h.request);
    input.renderEnv = { ...input.renderEnv, showHiddenText: true };
    h.controller.update({ ...h.settings, layoutRequest: () => JSON.stringify(input) });
    expect(await h.controller.ref.exportStructuredWithPages({ ...options, expectLayoutVersion: result.content.layout.layoutVersion }))
      .toMatchObject({ ok: false, failure: { code: 'stale-layout' } });
  } finally { await h.close(); }
});

function presentLayout(h: Awaited<ReturnType<typeof fixture>>) {
  h.viewStore.installFrameGeometry({ identity: h.viewStore.getSnapshot().identity, coversThroughSeq: 0,
    frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [] });
}

for (const revisionView of ['accepted', 'original'] as const) {
  test(`paged worker export waits for markup layout when exporting ${revisionView} revisions`, async () => {
    const h = await fixture(true);
    const preview = JSON.parse(h.request);
    preview.renderEnv = { ...preview.renderEnv, revisionPreview: { '1': 'accepted' } };
    h.resident.layoutDocumentWithRegionsRetainedJson(JSON.stringify(preview));
    const subscribe = h.viewStore.subscribe;
    h.viewStore.subscribe = listener => {
      const stop = subscribe(listener);
      queueMicrotask(() => { h.relayout(); presentLayout(h); });
      return stop;
    };
    try {
      expect(JSON.parse(h.resident.exportStructuredWithPagesJson({ revisionView }, h.request)))
        .toMatchObject({ ok: false, failure: { code: 'unsupported-revision-layout' } });
      const result = await h.controller.ref.exportStructuredWithPages({ revisionView });
      expect(result.ok).toBe(true);
      if (!result.ok) throw new Error(result.failure.message);
    } finally { await h.close(); }
  });
}

for (const previewView of ['accepted', 'rejected'] as const) {
  test(`paged worker export restores markup after a revised document's ${previewView} preview`, async () => {
    const h = await fixture(true, false, true);
    const preview = JSON.parse(h.request);
    preview.renderEnv = { ...preview.renderEnv, revisionPreview: { '1': previewView } };
    h.resident.layoutDocumentWithRegionsRetainedJson(JSON.stringify(preview));
    const subscribe = h.viewStore.subscribe;
    h.viewStore.subscribe = listener => {
      const stop = subscribe(listener);
      queueMicrotask(() => { h.relayout(); presentLayout(h); });
      return stop;
    };
    try {
      const result = await h.controller.ref.exportStructuredWithPages({ revisionView: 'markup' });
      if (!result.ok) throw new Error(result.failure.message);
      expect(JSON.stringify(result.content.structured)).toContain('Revised text');
      expect(result.content.layout.revisionView).toBe('markup');
    } finally { await h.close(); }
  });
}

test('paged worker export preserves an explicit revision preview refusal without waiting', async () => {
  const h = await fixture(true);
  const preview = JSON.parse(h.request);
  preview.renderEnv = { ...preview.renderEnv, revisionPreview: { '1': 'accepted' } };
  let subscriptions = 0;
  const subscribe = h.viewStore.subscribe;
  h.viewStore.subscribe = listener => { subscriptions += 1; return subscribe(listener); };
  h.controller.update({ ...h.settings, layoutRequest: () => JSON.stringify(preview) });
  try {
    expect(await h.controller.ref.exportStructuredWithPages({ revisionView: 'accepted' }))
      .toMatchObject({ ok: false, failure: { code: 'unsupported-revision-layout' } });
    expect(subscriptions).toBe(0);
  } finally { await h.close(); }
});

test('paged worker export polls while layout converges after the first completion', async () => {
  const h = await fixture(true);
  const next = JSON.parse(h.request);
  next.renderEnv = { ...next.renderEnv, showHiddenText: true };
  const request = JSON.stringify(next);
  let reads = 0;
  h.controller.update({ ...h.settings, layoutRequest: () => request, read: async (...args) => {
    const result = await h.settings.read(...args);
    reads += 1;
    if (reads === 2) h.resident.layoutDocumentWithRegionsRetainedJson(request);
    return result;
  } });
  const subscribe = h.viewStore.subscribe;
  h.viewStore.subscribe = listener => { const stop = subscribe(listener); queueMicrotask(() => presentLayout(h)); return stop; };
  try {
    expect((await h.controller.ref.exportStructuredWithPages({ revisionView: 'markup' })).ok).toBe(true);
    expect(reads).toBe(3);
  } finally { await h.close(); }
});

test('paged worker export polls for font readiness after layout completion', async () => {
  const h = await fixture(true);
  let requests = 0;
  let fontsReady = false;
  h.controller.update({ ...h.settings, layoutRequest: () => {
    requests += 1;
    if (requests === 2) queueMicrotask(() => {
      h.relayout();
      const state = h.viewStore.getSnapshot();
      h.viewStore.acceptWorkerState(state.identity, { epoch: 1, version: state.version + 1, throughSeq: 0,
        presentedThroughSeq: 0, selection: null, editState: state.editState });
      fontsReady = true;
    });
    return fontsReady ? h.request : null;
  } });
  const subscribe = h.viewStore.subscribe;
  h.viewStore.subscribe = listener => { const stop = subscribe(listener); queueMicrotask(() => presentLayout(h)); return stop; };
  try {
    expect((await h.controller.ref.exportStructuredWithPages({ revisionView: 'markup' })).ok).toBe(true);
    expect(requests).toBe(3);
  } finally { await h.close(); }
});

for (const replace of [false, true]) {
  test(`paged worker export ${replace ? 'throws on replacement' : 'refuses an edit admitted'} while its read is pending`, async () => {
    const h = await fixture();
    let release!: () => void;
    const held = new Promise<void>(resolve => { release = resolve; });
    let entered!: () => void;
    const reading = new Promise<void>(resolve => { entered = resolve; });
    h.controller.update({ ...h.settings, read: async (...args) => {
      const result = await h.settings.read(...args);
      entered();
      await held;
      return result;
    } });
    Object.assign(h.binding.session, { state: { selection: null, editState: { readOnly: false } } });
    h.binding.session.submit = (() => {
      const paragraph = h.resident.paragraphs('body')[0]!;
      const result = h.resident.proposalEngine.applyEdits({ expectVersion: h.resident.proposalEngine.version(), steps: [
        { op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: paragraph.paraId }, at: 'end', text: ' Concurrent edit' },
      ] });
      expect(result.ok).toBe(true);
      const state = h.viewStore.getSnapshot();
      h.viewStore.acceptWorkerState(state.identity, { epoch: 1, version: state.version + 1, throughSeq: 1,
        presentedThroughSeq: 0, selection: null, editState: state.editState });
      return { id: 1, seq: 1, ack: Promise.resolve({ status: 'applied' }), presented: Promise.resolve('presented') };
    }) as typeof h.binding.session.submit;
    h.binding.admission = createInputAdmission({ session: h.binding.session, viewStore: h.viewStore,
      echoPainter: { disable() {}, preview: () => undefined } as never });
    try {
      const pending = h.controller.ref.exportStructuredWithPages({ revisionView: 'markup' });
      await reading;
      if (replace) { h.controller.detach(h.binding); h.controller.attach({ ...h.binding }); }
      else expect(admitIntent(h.binding.admission, { kind: 'insertText', payload: { text: ' Concurrent edit' } }).status).toBe('submitted');
      release();
      if (replace) await expect(pending).rejects.toThrow('Document was replaced');
      else expect(await pending).toMatchObject({ ok: false, failure: { code: 'stale-document' } });
    } finally { release(); h.binding.admission.dispose(); await h.close(); }
  });
}
