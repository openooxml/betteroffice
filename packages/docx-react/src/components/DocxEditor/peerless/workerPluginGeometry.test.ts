import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createElement, memo, useSyncExternalStore } from 'react';
import { buildResidentRegionLayoutRequest } from '@betteroffice/docx/editor';
import {
  applyFrameDelta,
  createDisplayListQueries,
  decodeFrameDelta,
  encodeDisplayListFrameExtras,
  type DisplayListBuildInputs,
} from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as display from '@betteroffice/docx/layout/render';
import { preloadLayoutWasm } from '@betteroffice/docx/wasm/layout';
import {
  computeAnchorTargetGeometry,
  createYrsSession,
  proposalRevisionPreview,
  resolveReservations,
  type DocxReservationsPayload,
  type ProposalGeometryTarget,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { createResidentEngineSession, type ResidentEngineSession } from '@betteroffice/docx/yrs/residentEngineSession';
import { storyAnchorsDocx } from '../../../plugins/__fixtures__/storyAnchorsDocx';
import { createDocxPluginHost } from '../../../plugins/createDocxPluginHost';
import { defineDocxPlugin } from '../../../plugins/defineDocxPlugin';
import { scopePluginGeometry } from '../../../plugins/geometry';
import { PluginOverlays } from '../../../plugins/PluginOverlays';
import type { DocxAnchorGeometryResult, DocxGeometryTarget, DocxPluginGeometry } from '../../../plugins/types';
import { createWorkerPluginGeometry } from './workerPluginGeometry';
import { inputs, proposalHost } from './workerProposalHost.testing';
import type { PeerlessFrameGeometry } from './viewStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, render, waitFor } = await import('@testing-library/react');
const wasm = (name: string) => new Uint8Array(readFileSync(resolve(import.meta.dir, `../../../../../docx/src/wasm/generated/${name}_bg.wasm`)));
const FONT = new Uint8Array(readFileSync(resolve(import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')));
const PAGE_GAP = 24;
const HOST = { left: 10, top: 48 };
const destroy: { destroy(): void }[] = [];
beforeAll(async () => {
  await preloadLayoutWasm(wasm('layout/docx_layout'));
  await preloadEditWasm(wasm('edit/docx_edit'));
});
afterEach(() => { for (const session of destroy.splice(0)) session.destroy(); });
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const bounds = (x: number, y: number, width: number, height: number) => ({ x, y, left: x, top: y, width, height, right: x + width, bottom: y + height, toJSON() {} }) as DOMRect;

test('worker proposal anchors use page-local frame geometry at every zoom', async () => {
  const h = proposalHost();
  await h.submit('propose', { proposals: inputs }, {}).ack;
  const pages = document.createElement('div');
  pages.getBoundingClientRect = () => bounds(100, 100, 600, 2140);
  const host = document.createElement('div');
  host.setAttribute('data-docx-pages-host', '');
  host.getBoundingClientRect = pages.getBoundingClientRect;
  pages.appendChild(host);
  const layer = document.createElement('div');
  layer.getBoundingClientRect = () => bounds(80, 90, 600, 2140);
  for (let index = 0; index < 2; index++) {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    canvas.getBoundingClientRect = () => bounds(100, 100 + index * 1040, 600, 1000);
    host.appendChild(canvas);
  }
  const frame = h.binding.viewStore.getSnapshot().geometry!;
  h.binding.viewStore.installFrameGeometry({ ...frame, frameEpoch: frame.frameEpoch + 1, pages: Array.from({ length: 2 }, (_, pageIndex) => ({
    page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: pageIndex * 520, width: 300, height: 500 },
    transform: [1, 0, 0, 1, 0, pageIndex * 520],
  })) });
  const queries = { pageCount: () => 2, pageSize: () => ({ width: 300, height: 500 }),
    pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, width: 300, height: 500 }),
    rangeRects: (from: number, to: number) => from === 20 && to === 25 ? [{ pageIndex: 1, x: 30, y: 40, width: 50, height: 12 }] : [],
    anchorRect: () => null, dispose() {} } as unknown as display.DisplayListQueries;
  const load = spyOn(display, 'loadRustDisplayListQueryEngine').mockResolvedValue({} as never);
  const create = spyOn(display, 'createDisplayListQueries').mockReturnValue(queries);
  const geometry = createWorkerPluginGeometry(h.controller, () => pages, () => layer, () => {});
  try {
    h.controller.ref.setZoom(2);
    await geometry.update();
    const placed = geometry.geometry()!.getAnchorGeometry({ kind: 'proposal', id: inputs[2]!.id });
    expect(placed).toMatchObject({ ok: true, version: h.snapshot().version, previewVersion: h.snapshot().previewVersion,
      rects: [{ pageIndex: 1, x: 80, y: 1130, width: 100, height: 24 }], anchor: { pageIndex: 1, x: 180, y: 1130, width: 0, height: 24 } });
  } finally { geometry.dispose(); load.mockRestore(); create.mockRestore(); h.controller.dispose(); }
});

test('a viewport frame of the presented layout keeps every proposal anchor without waiting for another review read', async () => {
  const h = proposalHost();
  await h.submit('propose', { proposals: inputs }, {}).ack;
  const pages = document.createElement('div');
  pages.getBoundingClientRect = () => bounds(100, 100, 600, 2140);
  const host = document.createElement('div');
  host.setAttribute('data-docx-pages-host', '');
  host.getBoundingClientRect = pages.getBoundingClientRect;
  pages.appendChild(host);
  const layer = document.createElement('div');
  layer.getBoundingClientRect = () => bounds(80, 90, 600, 2140);
  const present = () => {
    const frame = h.binding.viewStore.getSnapshot().geometry!;
    h.binding.viewStore.installFrameGeometry({ ...frame, frameEpoch: frame.frameEpoch + 1, pages: Array.from({ length: 2 }, (_, pageIndex) => ({
      page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: pageIndex * 520, width: 300, height: 500 },
      transform: [1, 0, 0, 1, 0, pageIndex * 520],
    })) });
  };
  present();
  const queries = { pageCount: () => 2, pageSize: () => ({ width: 300, height: 500 }),
    pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, width: 300, height: 500 }),
    rangeRects: (from: number) => [{ pageIndex: from < 20 ? 0 : 1, x: 30, y: 40, width: 50, height: 12 }],
    anchorRect: () => null, dispose() {} } as unknown as display.DisplayListQueries;
  const load = spyOn(display, 'loadRustDisplayListQueryEngine').mockResolvedValue({} as never);
  const create = spyOn(display, 'createDisplayListQueries').mockReturnValue(queries);
  const geometry = createWorkerPluginGeometry(h.controller, () => pages, () => layer, () => {});
  const anchors = () => inputs.map(({ id }) => geometry.geometry()!.getAnchorGeometry({ kind: 'proposal', id }))
    .map(result => result.ok ? result.anchor.pageIndex : result.failure.code);
  try {
    await geometry.update();
    expect(anchors()).toEqual([0, 0, 1, 1]);
    const reads = h.readAt.mock.calls.length;
    h.readAt.mockImplementation(() => new Promise(() => {}));
    present();
    void geometry.update();
    for (let turn = 0; turn < 50; turn++) await Promise.resolve();
    expect(anchors()).toEqual([0, 0, 1, 1]);
    expect(h.readAt.mock.calls.length).toBe(reads);
  } finally { geometry.dispose(); load.mockRestore(); create.mockRestore(); h.controller.dispose(); }
});

async function builtProposalGeometry() {
  const h = proposalHost();
  await h.submit('propose', { proposals: inputs }, {}).ack;
  const pages = document.createElement('div');
  pages.getBoundingClientRect = () => bounds(100, 100, 600, 2140);
  const host = document.createElement('div');
  host.setAttribute('data-docx-pages-host', '');
  host.getBoundingClientRect = pages.getBoundingClientRect;
  pages.appendChild(host);
  const layer = document.createElement('div');
  layer.getBoundingClientRect = () => bounds(80, 90, 600, 2140);
  const store = h.binding.viewStore;
  const present = (change: Partial<PeerlessFrameGeometry> = {}, top = 0) => {
    const frame = store.getSnapshot().geometry!;
    store.installFrameGeometry({ ...frame, frameEpoch: frame.frameEpoch + 1, layoutEpoch: 1, coversThroughSeq: store.getSnapshot().throughSeq,
      pages: Array.from({ length: 2 }, (_, pageIndex) => ({ page: { pageId: String(pageIndex + 1), pageIndex },
        bounds: { x: 0, y: top + pageIndex * 520, width: 300, height: 500 }, transform: [1, 0, 0, 1, 0, top + pageIndex * 520] })), ...change });
  };
  const acknowledge = (edit: boolean) => {
    const view = store.getSnapshot();
    store.acceptWorkerState(view.identity, { epoch: 1, version: view.version + (edit ? 1 : 0), throughSeq: view.throughSeq + 1,
      presentedThroughSeq: view.presentedThroughSeq, selection: null,
      editState: edit ? { ...view.editState!, documentVersion: 'native-edited' } : view.editState! } as never);
  };
  present();
  const queries = { pageCount: () => 2, pageSize: () => ({ width: 300, height: 500 }),
    pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, width: 300, height: 500 }),
    rangeRects: (from: number) => [{ pageIndex: from < 20 ? 0 : 1, x: 30, y: 40, width: 50, height: 12 }],
    anchorRect: () => null, dispose() {} } as unknown as display.DisplayListQueries;
  const load = spyOn(display, 'loadRustDisplayListQueryEngine').mockResolvedValue({} as never);
  const create = spyOn(display, 'createDisplayListQueries').mockReturnValue(queries);
  const geometry = createWorkerPluginGeometry(h.controller, () => pages, () => layer, () => {});
  await geometry.update();
  const built = geometry.geometry()!;
  const anchors = (source = built) => inputs.map(({ id }) => source.getAnchorGeometry({ kind: 'proposal', id }))
    .map(result => result.ok ? result.anchor.pageIndex : result.failure.code);
  const dispose = () => { geometry.dispose(); load.mockRestore(); create.mockRestore(); h.controller.dispose(); };
  return { h, geometry, built, anchors, present, acknowledge, dispose };
}

test('a selection ack and a caret frame at the built document version keep answering from the built geometry until its rebuild lands', async () => {
  const { geometry, built, anchors, present, acknowledge, dispose } = await builtProposalGeometry();
  try {
    expect(anchors()).toEqual([0, 0, 1, 1]);
    acknowledge(false);
    expect(anchors()).toEqual([0, 0, 1, 1]);
    expect(geometry.layout()).toBe(built.layout);
    present();
    expect(anchors()).toEqual([0, 0, 1, 1]);
    await geometry.update();
    expect(geometry.geometry()).not.toBe(built);
    expect(anchors(geometry.geometry()!)).toEqual([0, 0, 1, 1]);
  } finally { dispose(); }
});

test('an edit, a relayout, a moved page, another pixel ratio or zoom answers layout-unavailable until the geometry is rebuilt', async () => {
  const unavailable = Array(4).fill('layout-unavailable');
  const pixelRatio = globalThis.devicePixelRatio;
  const changes: [string, (built: Awaited<ReturnType<typeof builtProposalGeometry>>) => void][] = [
    ['edit ack', ({ acknowledge }) => acknowledge(true)],
    ['edit ack then caret frame', ({ acknowledge, present }) => { acknowledge(true); present(); }],
    ['preview the worker reports', ({ h }) => {
      const view = h.binding.viewStore.getSnapshot();
      const review = view.editState!.review!;
      h.binding.viewStore.acceptWorkerState(view.identity, { epoch: 1, version: view.version, throughSeq: view.throughSeq + 1,
        presentedThroughSeq: view.presentedThroughSeq, selection: null,
        editState: { ...view.editState!, review: { ...review, previewVersion: review.previewVersion + 1 } } } as never);
    }],
    ['frame of another document version', ({ present }) => present({ documentVersion: 'native-remote' })],
    ['render environment relayout', ({ present }) => present({ layoutEpoch: 2 })],
    ['fonts relayout', ({ h, present }) => present({ layoutRevision: h.binding.viewStore.getSnapshot().geometry!.layoutRevision + 1 })],
    ['moved pages', ({ present }) => present({}, 40)],
    ['provisional layout', ({ present }) => present({ layoutProvisional: true })],
    ['pixel ratio', ({ present }) => { Object.defineProperty(globalThis, 'devicePixelRatio', { value: pixelRatio * 2, configurable: true }); present(); }],
    ['zoom', ({ h }) => h.controller.ref.setZoom(2)],
  ];
  for (const [name, change] of changes) {
    const built = await builtProposalGeometry();
    try {
      expect([name, built.anchors()]).toEqual([name, [0, 0, 1, 1]]);
      change(built);
      expect([name, built.anchors()]).toEqual([name, unavailable]);
      expect([name, built.geometry.layout()]).toEqual([name, null]);
    } finally {
      Object.defineProperty(globalThis, 'devicePixelRatio', { value: pixelRatio, configurable: true });
      built.dispose();
    }
  }
});

test('worker paragraph navigation passes focus intent without moving selection for a scroll', async () => {
  const h = proposalHost();
  const navigate = mock(async (_target: unknown, _focus?: boolean, _version?: string) => true);
  const readAt = h.readAt.mockImplementation(async () => ({ status: 'ok', version: h.snapshot().version,
    value: { loc: { story: 'body', paraId: 'p0', offset: 0 }, position: 0 } as never }));
  const geometry = createWorkerPluginGeometry(h.controller, () => null, () => null, () => {}, navigate);
  try {
    expect(await geometry.navigation.scrollToParagraph({ story: 'body', paraId: 'p0' }, { expectVersion: 'native-0' })).toEqual({ ok: true });
    expect(navigate).toHaveBeenLastCalledWith({ story: 'body', paraId: 'p0' }, false, 'native-0');
    expect(h.submit).not.toHaveBeenCalled();
    expect(await geometry.navigation.scrollToParagraph({ story: 'body', paraId: 'p0' }, { expectVersion: 'native-0', focus: true })).toEqual({ ok: true });
    expect(navigate).toHaveBeenLastCalledWith({ story: 'body', paraId: 'p0' }, true, 'native-0');
  } finally { readAt.mockClear(); geometry.dispose(); h.controller.dispose(); }
});

test('plugin selection ranges label the native worker caret with its accepted offset', async () => {
  const identity = { session: 's', incarnation: 1, epoch: 1 };
  const caret = { story: 'body', paraId: 'p', offset: 3 };
  const view = { identity, version: 0, throughSeq: 0, presentedThroughSeq: 0, workerSelection: { anchor: caret, focus: caret }, anchors: null, editState: undefined,
    geometry: { identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [] } };
  const reads: { kind: string; targets?: { range: unknown }[] }[] = [];
  const readAt = async (request: { kind: string; selection?: { anchor: { offset: number } } }) => {
    reads.push(request);
    return request.kind === 'rawSelectionInfo' ? { paraId: 'p', selectedText: '', paragraphText: 'abc', before: 'a', after: 'bc' }
      : [{ ok: false, failure: { code: 'missing-target', message: 'unused' } }];
  };
  const controller = { getSnapshot: () => ({ status: 'ready', view }), readAt,
    review: async () => ({ version: 'v1', previewVersion: 0 }), describedReview: () => null, ref: { getZoom: () => 1 } };
  const geometry = createWorkerPluginGeometry(controller as never, () => document.createElement('div'), () => document.createElement('div'), () => {});
  await geometry.update();
  expect(reads.find(read => read.kind === 'anchorTargets')?.targets?.[0]?.range).toEqual({ story: 'body', view: 'accepted',
    start: { paraId: 'p', offset: 1 }, end: { paraId: 'p', offset: 1 } });
  geometry.dispose();
});

interface LaidOut { main: YrsSession; engine: ResidentEngineSession; frame: ReturnType<typeof applyFrameDelta>; window: readonly [number, number] }

/** The fixture, changed by `prepare`, laid out as the worker frames it, with pages outside `[first, end)` unbuilt. */
async function laidOut(first = 0, end = 2 ** 32 - 1, reservations?: DocxReservationsPayload | ((main: YrsSession) => DocxReservationsPayload),
  prepare?: (main: YrsSession) => void): Promise<LaidOut> {
  const main = await createYrsSession();
  const engine = await createResidentEngineSession();
  destroy.push(main, engine);
  const host = main.openDocx(await storyAnchorsDocx(), true);
  prepare?.(main);
  const payload = typeof reservations === 'function' ? reservations(main) : reservations;
  const request = buildResidentRegionLayoutRequest(host.document, PAGE_GAP, { revisionPreview: proposalRevisionPreview(main.getProposals()),
    ...(payload ? { reservations: resolveReservations(main, payload, main.version()) } : {}) });
  const requirements = JSON.parse(main.layoutFontRequirementsJson(JSON.stringify(request))) as { key: string }[];
  main.clearFonts();
  const font = main.registerFont(FONT);
  const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [font]]));
  request.measurement = { fontChains, defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
    compat: { noLeading: false, doNotExpandShiftReturn: false }, authoritativeShaping: true } as typeof request.measurement;
  const { headersFooters } = JSON.parse(main.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request))) as Pick<DisplayListBuildInputs, 'headersFooters'>;
  main.setDisplayWindow(first, end);
  main.setDisplayRetainBuiltPages(false);
  const extras = encodeDisplayListFrameExtras({ fontChains, headersFooters } as DisplayListBuildInputs);
  const frame = applyFrameDelta(null, decodeFrameDelta(main.buildDisplayListFrame(extras, 0)));
  engine.loadState(main.encodeState());
  return { main, engine, frame, window: [first, end] };
}

interface ShellOptions {
  zoom?: number; window?: readonly [number, number]; answer?(result: ProposalGeometryTarget): ProposalGeometryTarget;
  /** The reservation set the frame was laid out with, and the one the editor sends. */
  laidOutWith?: { version: string; key: string | null }; sent?(): DocxReservationsPayload | null;
}

/** The peer-less plugin geometry over `laid`'s frame, its targets answered by the worker read. */
async function shellOf({ engine, frame, window: built }: LaidOut, { zoom = 1, window, answer = (result) => result, laidOutWith, sent = () => null }: ShellOptions = {}) {
  const shown = window ?? built;
  const container = document.createElement('div');
  const host = document.createElement('div');
  host.setAttribute('data-docx-pages-host', '');
  container.appendChild(host);
  let top = 0;
  const pages = frame.displayList.pages.map((displayPage, pageIndex) => {
    const bounds = { x: 0, y: top, width: displayPage.width, height: displayPage.height };
    top += displayPage.height + PAGE_GAP;
    if (!displayPage.unbuilt || (pageIndex >= shown[0] && pageIndex < shown[1])) {
      const element = document.createElement('div');
      element.className = 'canvas-page';
      element.dataset.pageIndex = String(pageIndex);
      element.getBoundingClientRect = () => new DOMRect(HOST.left + bounds.x * zoom, HOST.top + bounds.y * zoom, bounds.width * zoom, bounds.height * zoom);
      host.appendChild(element);
    }
    return { page: { pageId: String(pageIndex + 1), pageIndex }, bounds, transform: [1, 0, 0, 1, 0, bounds.y], displayPage };
  });
  const width = Math.max(...pages.map(({ bounds }) => bounds.width));
  host.getBoundingClientRect = () => new DOMRect(HOST.left, HOST.top, width * zoom, top * zoom);
  container.getBoundingClientRect = () => new DOMRect(0, 0, HOST.left + width * zoom, HOST.top + top * zoom);
  const layer = document.createElement('div');
  layer.getBoundingClientRect = container.getBoundingClientRect;
  const identity = { session: 's', incarnation: 1, epoch: 1 };
  const view = { identity, version: 0, throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null, editState: undefined,
    geometry: { identity, coversThroughSeq: 0, frameEpoch: frame.frameEpoch, layoutRevision: 1, pages, caret: null, selectionRects: [],
      ...(laidOutWith ? { reservations: laidOutWith } : {}) } };
  const version = engine.geometryReader.version();
  const review = { version, previewVersion: 0 };
  const reads: unknown[] = [];
  const controller = { getSnapshot: () => ({ status: 'ready', view }),
    review: async () => review, describedReview: () => review,
    readAt: async (request: { targets: Parameters<typeof computeAnchorTargetGeometry>[1] }) => {
      reads.push(request);
      return computeAnchorTargetGeometry(engine.geometryReader, request.targets).map(answer);
    },
    ref: { getZoom: () => zoom, readParagraphs: async () => ({ ok: true, version, paragraphs: [] }) } };
  const plugin = createWorkerPluginGeometry(controller as never, () => container, () => layer, () => {}, undefined, sent);
  if (window) plugin.setDisplayWindow(window);
  await plugin.update();
  return {
    plugin,
    version,
    reads,
    bounds: pages.map(({ bounds }) => bounds),
    async anchor(target: DocxGeometryTarget) {
      return (await plugin.geometry()!.readAnchorGeometries([target]))[0]!;
    },
  };
}

function range({ main }: LaidOut, version: string, story: string, paraId: string, text: string) {
  const offset = main.paragraphs(story).find((paragraph) => paragraph.paraId === paraId)!.text.indexOf(text);
  return { kind: 'range', version, range: { story, start: { paraId, offset }, end: { paraId, offset: offset + text.length }, view: 'accepted' } } as const;
}

function ok(result: DocxAnchorGeometryResult) {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result;
}

const pageIndices = (result: DocxAnchorGeometryResult) => [...new Set(ok(result).rects.map(({ pageIndex }) => pageIndex))];

test('a header target has rects on every built page that paints it and its anchor on the earliest', async () => {
  const laid = await laidOut();
  const shell = await shellOf(laid);
  const odd = ok(await shell.anchor(range(laid, shell.version, 'hf:rIdH1', '10000001', '{{odd}}')));
  expect(pageIndices(odd)).toEqual([2, 4]);
  expect(odd.unbuiltPages).toBeUndefined();
  expect(odd.anchor).toMatchObject({ pageIndex: 2, width: 0 });
  expect(odd.anchor.x).toBeCloseTo(odd.rects[0]!.x + odd.rects[0]!.width, 6);
  expect(pageIndices(await shell.anchor(range(laid, shell.version, 'hf:rIdH2', '10000011', '{{even}}')))).toEqual([1, 3, 5]);
  expect(pageIndices(await shell.anchor(range(laid, shell.version, 'hf:rIdF1', '10000021', '{{first}}')))).toEqual([0]);

  const partial = await laidOut(0, 3);
  const early = await shellOf(partial);
  const built = ok(await early.anchor(range(partial, early.version, 'hf:rIdH1', '10000001', '{{odd}}')));
  expect(pageIndices(built)).toEqual([2]);
  expect(built.anchor.pageIndex).toBe(2);
  expect(built.unbuiltPages).toEqual([4]);
  const even = ok(await early.anchor(range(partial, early.version, 'hf:rIdH2', '10000011', '{{even}}')));
  expect(pageIndices(even)).toEqual([1]);
  expect(even.unbuiltPages).toEqual([3, 5]);
});

test('a body table cell target has its own rects, anchored at its end', async () => {
  const laid = await laidOut();
  const shell = await shellOf(laid);
  const intro = ok(await shell.anchor(range(laid, shell.version, 'body', '00000001', 'Intro')));
  const cell = ok(await shell.anchor(range(laid, shell.version, 'body:t0:r0c0', '00000002', '{{cell}}')));
  expect(cell.rects).toHaveLength(1);
  const [rect] = cell.rects;
  expect(rect).toMatchObject({ pageIndex: 0 });
  expect(rect!.x).toBeGreaterThan(intro.rects[0]!.x);
  expect(rect!.y).toBeGreaterThan(intro.rects[0]!.y);
  expect(cell.anchor).toMatchObject({ pageIndex: 0, y: rect!.y, width: 0 });
  expect(cell.anchor.x).toBeCloseTo(rect!.x + rect!.width, 6);
  expect(cell.unbuiltPages).toBeUndefined();
});

test('hidden ranges of a header target are left out of its rects', async () => {
  const laid = await laidOut();
  const version = laid.engine.geometryReader.version();
  const token = range(laid, version, 'hf:rIdH1', '10000001', '{{odd}}');
  const [hidden] = computeAnchorTargetGeometry(laid.engine.geometryReader, [token]);
  if (!hidden?.ok) throw new Error('the token has no display range');
  const tokenRects = ok(await (await shellOf(laid)).anchor(token)).rects;
  const shell = await shellOf(laid, { answer: (result) => result.ok ? { ...result, hidden: hidden.ranges } : result });
  const line = ok(await shell.anchor(range(laid, version, 'hf:rIdH1', '10000001', 'Odd {{odd}} header')));
  expect(pageIndices(line)).toEqual([2, 4]);
  for (const hole of tokenRects) {
    const pieces = line.rects.filter(({ pageIndex }) => pageIndex === hole.pageIndex);
    expect(pieces.every(({ x, width }) => x + width <= hole.x + 0.01 || x >= hole.x + hole.width - 0.01)).toBe(true);
    expect(pieces.some(({ x }) => x < hole.x)).toBe(true);
    expect(pieces.some(({ x }) => x >= hole.x + hole.width - 0.01)).toBe(true);
  }
});

test('a part no page paints is missing once every page is built and waits while pages are unbuilt', async () => {
  for (const [end, failure] of [[2 ** 32 - 1, { code: 'missing-target', message: 'No page paints this header or footer' }],
    [3, { code: 'layout-unavailable', message: 'No laid-out page paints this header or footer' }]] as const) {
    const laid = await laidOut(0, end);
    for (const page of laid.frame.displayList.pages) {
      delete page.footer;
      if (page.hfParts) delete page.hfParts.footer;
    }
    const shell = await shellOf(laid);
    expect(await shell.anchor(range(laid, shell.version, 'hf:rIdF1', '10000021', '{{first}}'))).toEqual({ ok: false, failure });
  }
});

test('a header target waits while a page in the display window that paints it is unbuilt', async () => {
  const laid = await laidOut(0, 3);
  const shell = await shellOf(laid, { window: [3, 5] });
  expect(await shell.anchor(range(laid, shell.version, 'hf:rIdH1', '10000001', '{{odd}}'))).toEqual({
    ok: false, failure: { code: 'layout-unavailable', message: 'No rendered layout shows this target yet' },
  });
  expect(pageIndices(await shell.anchor(range(laid, shell.version, 'hf:rIdF1', '10000021', '{{first}}')))).toEqual([0]);
});

test('moving the display window onto an unbuilt page that paints a header target notifies and waits', async () => {
  const laid = await laidOut(0, 3);
  const shell = await shellOf(laid, { window: [0, 3] });
  const target = range(laid, shell.version, 'hf:rIdH1', '10000001', '{{odd}}');
  expect(pageIndices(await shell.anchor(target))).toEqual([2]);
  let published = 0;
  shell.plugin.subscribe(() => { published += 1; });
  shell.plugin.setDisplayWindow([3, 5]);
  expect(published).toBe(1);
  expect(shell.plugin.geometry()!.getAnchorGeometry(target)).toEqual({
    ok: false, failure: { code: 'layout-unavailable', message: 'No rendered layout shows this target yet' },
  });
});

test('a memoized plugin overlay reads a header target again when its answer arrives and when the window moves', async () => {
  const laid = await laidOut(0, 3);
  const shell = await shellOf(laid, { window: [0, 3] });
  const target = range(laid, shell.version, 'hf:rIdH1', '10000001', '{{odd}}');
  const seen: DocxAnchorGeometryResult[] = [];
  const overlay = memo(function HeaderOverlay({ geometry }: { context: unknown; geometry: DocxPluginGeometry }) {
    seen.push(geometry.getAnchorGeometry(target));
    return null;
  });
  const host = createDocxPluginHost({ pagedEditorRef: { current: null }, writeMode: () => 'viewing', commands: () => null,
    layout: () => ({ queries: null, complete: false, failed: false }), subscribeLayout: () => () => {},
    geometry: shell.plugin.geometry, subscribeGeometry: shell.plugin.subscribe, translate: (key) => key });
  host.setPlugins([defineDocxPlugin({ id: 'test.header-overlay', createState: () => null, overlay })]);
  host.openWorker(shell.version);
  await waitFor(() => expect(host.activations()).toHaveLength(1));
  const layout = shell.plugin.layout()!;
  host.layoutChanged(layout);
  host.geometryPresented(layout);
  function Overlays() {
    const activations = useSyncExternalStore(host.subscribe, host.activations, host.activations);
    return createElement(PluginOverlays, { host, activations, target: document.body, layerRef: () => {} });
  }
  const view = render(createElement(Overlays));
  await waitFor(() => expect(seen.at(-1)?.ok).toBe(true));
  expect(pageIndices(seen.at(-1)!)).toEqual([2]);
  const answered = seen.length;
  await act(async () => { shell.plugin.setDisplayWindow([3, 5]); });
  expect(seen.length).toBeGreaterThan(answered);
  expect(seen.at(-1)).toEqual({ ok: false, failure: { code: 'layout-unavailable', message: 'No rendered layout shows this target yet' } });
  view.unmount();
  host.close('unmounted');
});

test('header rects are page-local at 100% and 150% zoom', async () => {
  const laid = await laidOut();
  const version = laid.engine.geometryReader.version();
  const target = range(laid, version, 'hf:rIdH1', '10000001', '{{odd}}');
  const [token] = computeAnchorTargetGeometry(laid.engine.geometryReader, [target]);
  if (!token?.ok) throw new Error('the token has no display range');
  const queries = createDisplayListQueries(laid.frame.displayList, laid.main as never);
  const raw = queries.hfRangeRects('header', 'rIdH1', token.ranges[0]!.from, token.ranges[0]!.to).filter(({ width }) => width > 0);
  expect(raw.map(({ pageIndex }) => pageIndex)).toEqual([2, 4]);
  for (const zoom of [1, 1.5]) {
    const shell = await shellOf(laid, { zoom });
    const result = ok(await shell.anchor(target));
    expect(result.rects).toHaveLength(raw.length);
    result.rects.forEach((rect, index) => {
      const page = shell.bounds[rect.pageIndex]!;
      const expected = raw[index]!;
      expect(rect.pageIndex).toBe(expected.pageIndex);
      expect((rect.x - HOST.left) / zoom - page.x).toBeCloseTo(expected.x, 6);
      expect((rect.y - HOST.top) / zoom - page.y).toBeCloseTo(expected.y, 6);
      expect(rect.width / zoom).toBeCloseTo(expected.width, 6);
      expect(rect.height / zoom).toBeCloseTo(expected.height, 6);
    });
    const page = shell.bounds[2]!;
    expect(result.pageRect.y).toBeCloseTo(HOST.top + page.y * zoom, 6);
    expect(result.pageRect.height).toBeCloseTo(page.height * zoom, 6);
  }
  queries.dispose();
});

test('an anchor on an unbuilt page with no page element follows the frame placement at 150% zoom', async () => {
  const laid = await laidOut(3, 6);
  const zoom = 1.5;
  const shell = await shellOf(laid, { zoom });
  const content = laid.frame.displayList.pages[0]!.contentBounds;
  if (!content) throw new Error('the unbuilt page has no content box');
  const footer = ok(await shell.anchor(range(laid, shell.version, 'hf:rIdF1', '10000021', '{{first}}')));
  expect(footer.rects).toEqual([]);
  const page = shell.bounds[0]!;
  for (const [actual, expected] of [
    [footer.anchor, { pageIndex: 0, x: HOST.left + (page.x + content.x) * zoom, y: HOST.top + (page.y + content.y + content.height) * zoom, width: 0, height: 0 }],
    [footer.pageRect, { x: HOST.left + page.x * zoom, y: HOST.top + page.y * zoom, width: page.width * zoom, height: page.height * zoom }],
  ] as const) {
    for (const [key, value] of Object.entries(expected)) expect(actual[key as keyof typeof actual]).toBeCloseTo(value, 6);
  }
});

test('a header answer without its relationship ids is refused', async () => {
  const laid = await laidOut();
  const shell = await shellOf(laid, { answer: (result) => { if (result.ok) delete result.rIds; return result; } });
  expect(await shell.anchor(range(laid, shell.version, 'hf:rIdH1', '10000001', '{{odd}}'))).toEqual({
    ok: false, failure: { code: 'unsupported', message: 'The target has no display position' },
  });
});

test('pages released to unbuilt keep the parts their build paints across a frame delta', async () => {
  const laid = await laidOut();
  const painted = laid.frame.displayList.pages.map(({ header, footer }) => [header?.rId ?? null, footer?.rId ?? null]);
  const released = laid.main.releaseDisplayPagesFrame([0, 1, 2], laid.frame.frameEpoch);
  if (!released) throw new Error('the release was superseded');
  const delta = decodeFrameDelta(released);
  expect(delta.full).toBe(false);
  const { displayList } = applyFrameDelta(laid.frame, delta);
  expect(displayList.pages.map(({ unbuilt, header, footer, hfParts }) =>
    [unbuilt ?? false, ...(unbuilt ? [hfParts?.header ?? null, hfParts?.footer ?? null] : [header?.rId ?? null, footer?.rId ?? null])]))
    .toEqual(painted.map((parts, index) => [index < 3, ...parts]));
});

const at = (paraId: string, offset: number) => ({ paraId, offset });
/** The probe sets with their anchors read at `version`. */
function reservedAt(version: string): DocxReservationsPayload {
  const span = (paraId: string, from: number, to: number) =>
    ({ kind: 'range', version, range: { story: 'body', start: at(paraId, from), end: at(paraId, to), view: 'accepted' } }) as const;
  return { version, sets: [
    { owner: 'probe', items: [
      { id: 'chip', kind: 'inline', width: 40, height: 20, anchor: span('00000004', 7, 7) },
      { id: 'gap', kind: 'block', side: 'after', height: 60, anchor: span('00000005', 0, 0) },
      { id: 'gone', kind: 'hide', anchor: span('00000005', 6, 11) },
      { id: 'late', kind: 'inline', width: 30, height: 10, anchor: span('00000008', 0, 0) },
    ] },
    { owner: 'other', items: [{ id: 'chip', kind: 'inline', width: 24, height: 12, anchor: span('00000001', 0, 0) }] },
  ] };
}
const LAID_OUT = { version: 'v', key: '1' };
const sentAs = (key: string, version = 'v') => ({ ...reservedAt(version), key });
const reservation = (id: string) => ({ kind: 'reservation', id }) as const;

async function reservedShell(first?: number, end?: number) {
  const laid = await laidOut(first, end, reservedAt('v'));
  const version = laid.engine.geometryReader.version();
  return { laid, shell: await shellOf(laid, { laidOutWith: LAID_OUT, sent: () => sentAs('1', version) }) };
}

/** `owner`'s box `id` on the frame pages, which the answer must place in overlay pixels. */
function expectBox(result: DocxAnchorGeometryResult, { frame }: LaidOut, { bounds }: { bounds: { x: number; y: number }[] }, owner: string, id: string) {
  const pageIndex = frame.displayList.pages.findIndex((page) => page.reservations?.some((box) => box.owner === owner && box.id === id));
  const box = frame.displayList.pages[pageIndex]!.reservations!.find((entry) => entry.owner === owner && entry.id === id)!;
  const expected = { x: HOST.left + bounds[pageIndex]!.x + box.x, y: HOST.top + bounds[pageIndex]!.y + box.y, width: box.width, height: box.height };
  const answered = ok(result);
  expect(answered.rects).toHaveLength(1);
  for (const rect of [answered.rects[0]!, answered.anchor]) {
    expect(rect.pageIndex).toBe(pageIndex);
    for (const key of ['x', 'y', 'width', 'height'] as const) expect(rect[key]).toBeCloseTo(expected[key], 6);
  }
}

test('the worker-only layout reports the reservation set its frame was laid out with', async () => {
  expect((await reservedShell()).shell.plugin.layout()!.reservationsVersion).toBe('v');
  expect((await shellOf(await laidOut())).plugin.layout()!.reservationsVersion).toBeNull();
});

test('a plugin reads its own inline and block boxes from the frame pages, batched with other targets, without a worker read', async () => {
  const { laid, shell } = await reservedShell();
  const geometry = shell.plugin.geometry()!;
  const probe = scopePluginGeometry(geometry, 'probe');
  const [chip, gap] = await probe.readAnchorGeometries([reservation('chip'), reservation('gap')]);
  expect(shell.reads).toHaveLength(0);
  expectBox(chip!, laid, shell, 'probe', 'chip');
  expectBox(gap!, laid, shell, 'probe', 'gap');
  expect(ok(gap!).rects[0]!.height).toBeCloseTo(60, 6);
  expectBox((await scopePluginGeometry(geometry, 'other').readAnchorGeometries([reservation('chip')]))[0]!, laid, shell, 'other', 'chip');
  expect(probe.getAnchorGeometry(reservation('chip'))).toEqual(chip!);
  const [mixed, intro] = await probe.readAnchorGeometries([reservation('chip'), range(laid, shell.version, 'body', '00000001', 'Intro')]);
  expect(mixed).toEqual(chip!);
  ok(intro!);
  expect(geometry.getAnchorGeometry(reservation('chip'))).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
  expect(probe.getPageRect(0)).toEqual(geometry.getPageRect(0));
  expect(probe.isPageBuilt(0)).toBe(true);
});

test('a frame laid out with another reservation set than the one the editor sends answers layout-unavailable', async () => {
  let sent = sentAs('2');
  const laid = await laidOut(undefined, undefined, reservedAt('v'));
  const shell = await shellOf(laid, { laidOutWith: LAID_OUT, sent: () => sent });
  const probe = scopePluginGeometry(shell.plugin.geometry()!, 'probe');
  const unavailable = { ok: false, failure: { code: 'layout-unavailable' } };
  expect(probe.getAnchorGeometry(reservation('chip'))).toMatchObject(unavailable);
  expect(await probe.readAnchorGeometries([reservation('chip'), reservation('gone')])).toMatchObject([unavailable, unavailable]);
  const unstamped = await shellOf(laid, { sent: () => sentAs('1') });
  expect(scopePluginGeometry(unstamped.plugin.geometry()!, 'probe').getAnchorGeometry(reservation('chip'))).toMatchObject(unavailable);
  sent = sentAs('1');
  expectBox(probe.getAnchorGeometry(reservation('chip')), laid, shell, 'probe', 'chip');
});

test('a hidden reservation answers the boundary its hidden text leaves, read with the other targets', async () => {
  const { laid, shell } = await reservedShell();
  const probe = scopePluginGeometry(shell.plugin.geometry()!, 'probe');
  const [gone, before] = await probe.readAnchorGeometries([reservation('gone'), range(laid, shell.version, 'body', '00000005', 'After ')]);
  const hidden = ok(gone!);
  const visible = ok(before!);
  expect(shell.reads).toHaveLength(1);
  expect(hidden.rects).toEqual([]);
  expect(hidden.anchor).toMatchObject({ pageIndex: visible.anchor.pageIndex, width: 0 });
  expect(hidden.anchor.x).toBeCloseTo(visible.anchor.x, 6);
  expect(hidden.anchor.y).toBeCloseTo(visible.anchor.y, 6);
});

const span = (version: string, paraId: string, from: number, to: number) =>
  ({ kind: 'range', version, range: { story: 'body', start: at(paraId, from), end: at(paraId, to), view: 'accepted' } }) as const;

/** `sets` laid out with the fixture as one sent set, their paragraph anchors naming each session's `sessionId`. */
async function setShell(sets: (version: string, sessionId: string) => DocxReservationsPayload['sets'], prepare?: (main: YrsSession) => void) {
  const laid = await laidOut(undefined, undefined, (main) => ({ version: 'v', sets: sets('v', main.paragraphIdentities().sessionId) }), prepare);
  const version = laid.engine.geometryReader.version();
  const sent = { version, key: '1', sets: sets(version, laid.engine.paragraphIdentities().sessionId) };
  return { laid, shell: await shellOf(laid, { laidOutWith: LAID_OUT, sent: () => sent }) };
}

test.each(['another owner', 'the document'] as const)('a hidden reservation after text %s hides answers the edge of the visible text before both', async (hider) => {
  const { laid, shell } = await setShell((version) => [
    { owner: 'probe', items: [{ id: 'gone', kind: 'hide' as const, anchor: span(version, '00000005', 6, 11) }] },
    ...(hider === 'another owner' ? [{ owner: 'other', items: [{ id: 'cut', kind: 'hide' as const, anchor: span(version, '00000005', 3, 6) }] }] : []),
  ], hider === 'the document' ? (main) => {
    const { end } = main.locateParagraph('body', '00000005');
    main.applyRawOps('body', [{ op: 'format', index: end - 'er split'.length, len: 'er '.length, attrs: { hidden: true } }]);
  } : undefined);
  const [gone, kept] = await scopePluginGeometry(shell.plugin.geometry()!, 'probe').readAnchorGeometries([
    reservation('gone'), range(laid, shell.version, 'body', '00000005', 'Aft')]);
  const hidden = ok(gone!);
  const visible = ok(kept!);
  expect(hidden.rects).toEqual([]);
  expect(hidden.anchor).toMatchObject({ pageIndex: visible.anchor.pageIndex, width: 0 });
  expect(hidden.anchor.x).toBeCloseTo(visible.anchor.x, 6);
  expect(hidden.anchor.y).toBeCloseTo(visible.anchor.y, 6);
});

test('a hidden whole paragraph answers the start of the visible paragraph after it', async () => {
  const { laid, shell } = await setShell((_, sessionId) => [{ owner: 'probe', items: [{ id: 'gone', kind: 'hide',
    anchor: { kind: 'paragraph', paragraph: { kind: 'session', sessionId, story: 'body', paraId: '00000005' } } }] }]);
  const [gone, next] = await scopePluginGeometry(shell.plugin.geometry()!, 'probe').readAnchorGeometries([
    reservation('gone'), range(laid, shell.version, 'body', '00000006', 'Page four')]);
  const hidden = ok(gone!);
  const first = ok(next!).rects[0]!;
  expect(hidden.rects).toEqual([]);
  expect(hidden.anchor).toMatchObject({ pageIndex: first.pageIndex, width: 0 });
  expect(hidden.anchor.x).toBeCloseTo(first.x, 6);
  expect(hidden.anchor.y).toBeCloseTo(first.y, 6);
});

test('a reservation box on an unbuilt page waits for its layout, and an unknown one is missing once laid out', async () => {
  const unavailable = { ok: false, failure: { code: 'layout-unavailable' } };
  const early = await reservedShell(0, 3);
  expect(await scopePluginGeometry(early.shell.plugin.geometry()!, 'probe').readAnchorGeometries([reservation('late')])).toMatchObject([unavailable]);
  const probe = scopePluginGeometry((await reservedShell()).shell.plugin.geometry()!, 'probe');
  ok((await probe.readAnchorGeometries([reservation('late')]))[0]!);
  expect(await probe.readAnchorGeometries([reservation('nope')])).toMatchObject([{ ok: false, failure: { code: 'missing-target' } }]);
});
