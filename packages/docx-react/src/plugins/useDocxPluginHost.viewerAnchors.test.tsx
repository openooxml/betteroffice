import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import {
  computeAnchorDisplayTarget,
  createYrsSession,
  type AnchorDisplayTarget,
  type ResidentDocumentRead,
  type ResidentEngineWorkerClient,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from '@betteroffice/docx/yrs/residentEngineSession';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import { createRenderedDomContext } from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  markPresented,
  stampRevisionPreviewKey,
  stampSourceVersion,
  stampWorkerFrameVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import { defineDocxPlugin } from './defineDocxPlugin';
import { currentPreviewKey } from './proposalPreview';
import type { DocxAnchorGeometryResult, DocxGeometryTarget } from './types';
import { useDocxPluginHost, type UseDocxPluginHostOptions } from './useDocxPluginHost';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const elements: HTMLElement[] = [];
const sessions: Array<{ destroy(): void }> = [];
afterEach(() => {
  cleanup();
  for (const element of elements.splice(0)) element.remove();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});
beforeAll(async () => {
  const { preloadEditWasm } = await import('@betteroffice/docx/wasm/edit');
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  );
});

async function trackedDocument() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  );
  zip.file(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Alpha</w:t></w:r><w:del w:id="1" w:author="Reviewer"><w:r><w:delText>gone</w:delText></w:r></w:del><w:r><w:t>Beta</w:t></w:r></w:p><w:p w14:paraId="00000003"><w:r><w:t>AB</w:t></w:r><w:ins w:id="3" w:author="Reviewer"><w:r><w:t>xy</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>'
  );
  return zip.generateAsync({ type: 'uint8array' });
}

function residentRead(engine: ResidentEngineSession, fail = () => false) {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    if (fail()) throw new Error('worker restarting');
    const version = engine.proposalEngine.version();
    if (request.kind === 'resolveParagraphAnchors') {
      return {
        version,
        value: { results: request.anchors.map((anchor) => engine.geometryReader.resolveParagraphAnchor(anchor)) },
      };
    }
    if (request.kind !== 'anchorTarget') throw new Error(`unexpected ${request.kind} read`);
    return {
      version,
      value:
        version === request.expectVersion
          ? computeAnchorDisplayTarget(engine.geometryReader, request.target, request.revisionPreview)
          : null,
    };
  }) as ResidentEngineWorkerClient['documentRead'];
  return { read, requests };
}

function queries(
  version: string,
  workerVersion: string,
  previewKey: string,
  unbuilt = false
): DisplayListQueries {
  const pages = [
    { pageIndex: 0, width: 100, height: 200 },
    ...(unbuilt ? [{ pageIndex: 1, width: 100, height: 200, unbuilt: true, positionSpan: [0, 1000] }] : []),
  ];
  const result = {
    displayList: { pages },
    pageCount: () => pages.length,
    pageSize: () => ({ width: 100, height: 200 }),
    rangeRects: (from: number, to: number) =>
      from < to ? [{ pageIndex: 0, x: from, y: 10, width: to - from, height: 12 }] : [],
    anchorRect: (position: number) => ({ pageIndex: 0, x: position, y: 10, width: 1, height: 12 }),
    pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: pageIndex * 200, width: 100, height: 200 }),
    hitTestRegions: () => null,
    sourceState: () => ({ status: 'ready' }),
    whenReady: () => Promise.resolve(),
  } as unknown as DisplayListQueries;
  stampSourceVersion(result, version);
  stampWorkerFrameVersion(result, workerVersion);
  stampRevisionPreviewKey(result, previewKey);
  return result;
}

async function host(
  session: YrsSession,
  displayQueries: DisplayListQueries,
  viewerDocumentRead?: ResidentEngineWorkerClient['documentRead'],
  { presented = true, onEvent }: Pick<Parameters<typeof defineDocxPlugin>[0], 'onEvent'> & { presented?: boolean } = {}
) {
  const pages = document.createElement('div');
  for (const pageIndex of displayQueries.displayList.pages.keys()) {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(pageIndex);
    canvas.getBoundingClientRect = () => new DOMRect(0, pageIndex * 200, 100, 200);
    pages.append(canvas);
  }
  const target = document.createElement('div');
  document.body.append(pages, target);
  elements.push(pages, target);
  pages.getBoundingClientRect = target.getBoundingClientRect = () => new DOMRect(0, 0, 100, 200);
  const editor = {
    getYrsSession: () => (viewerDocumentRead ? null : session),
    getLayout: () => null,
    getSelectionRange: () => null,
    hasPendingInput: () => false,
    yrsLocToDisplayPosition: () => 2,
  } as unknown as PagedEditorRef;
  const options: UseDocxPluginHostOptions = {
    plugins: [defineDocxPlugin({ id: 'test.viewer-anchors', createState: () => null, overlay: () => null, onEvent })],
    pagedEditorRef: { current: editor },
    writeModeRef: { current: 'viewing' },
    mode: 'viewing',
    readOnly: true,
    commands: createDocxCommandController(),
    session,
    loadGeneration: 0,
    queries: displayQueries,
    viewerDocumentRead,
    layoutError: null,
    zoom: 1,
    canvasHostRef: { current: pages },
    overlayTarget: target,
    selectionChangeSubscribersRef: { current: new Set() },
    i18n: undefined,
    onRenderedDomContextReady: undefined,
  };
  if (presented) markPresented(pages, displayQueries.displayList);
  const view = renderHook(useDocxPluginHost, { initialProps: options });
  await act(async () => {
    view.result.current.overlayLayerRef(target as HTMLDivElement);
    view.result.current.onRenderedDomContext(createRenderedDomContext(pages, 1), displayQueries);
  });
  return Object.assign(view, { present: () => markPresented(pages, displayQueries.displayList) });
}

async function geometryOf(result: Awaited<ReturnType<typeof host>>['result']) {
  await waitFor(() => expect(result.current.activations[0]?.context.geometry).toBeTruthy());
  return result.current.activations[0]!.context.geometry!;
}

function shownSpans(answer: AnchorDisplayTarget) {
  if (!answer.ok) throw new Error(answer.failure.message);
  const spans: Array<{ x: number; width: number }> = [];
  for (const { from, to } of [...answer.ranges].sort((a, b) => a.from - b.from)) {
    let start = from;
    for (const hidden of [...answer.hidden].sort((a, b) => a.from - b.from)) {
      if (hidden.to <= start || hidden.from >= to) continue;
      if (hidden.from > start) spans.push({ x: start, width: hidden.from - start });
      start = Math.max(start, hidden.to);
    }
    if (start < to) spans.push({ x: start, width: to - start });
  }
  return spans;
}

const persisted = {
  kind: 'persisted' as const,
  story: { partUri: '/word/document.xml', kind: 'body' as const },
  paraId: '00000003',
};

async function workerViewer(
  options: { unbuilt?: boolean; fail?: () => boolean } & Parameters<typeof host>[3] = {}
) {
  const main = await createYrsSession({ clientId: 901 });
  const engine = await createResidentEngineSession(undefined, 902);
  sessions.push(main, engine);
  main.openDocx(await trackedDocument(), true);
  engine.loadState(main.encodeState());
  const found = main.findText({
    text: 'Beta',
    within: { kind: 'paragraph', story: 'body', paraId: '00000001' },
    view: 'accepted',
  });
  if (!found.ok) throw new Error(found.failure.message);
  const targets: Exclude<DocxGeometryTarget, { kind: 'proposal' }>[] = [
    { kind: 'paragraph', paragraph: persisted },
    { kind: 'search', paragraph: persisted, text: 'xy' },
    { kind: 'range', version: main.version(), range: found.matches[0]!.range },
    { kind: 'revision', revisionId: main.listRevisions()[0]!.revisionId },
  ];
  const { read, requests } = residentRead(engine, options.fail);
  const workerVersion = engine.proposalEngine.version();
  const view = await host(
    main,
    queries(main.version(), workerVersion, currentPreviewKey(main), options.unbuilt),
    read,
    options
  );
  return { main, engine, targets, requests, workerVersion, view, geometry: await geometryOf(view.result) };
}

test('a worker viewer reads paragraph, search, range and revision geometry from the worker', async () => {
  const { main, targets, requests, workerVersion, geometry } = await workerViewer();
  expect(workerVersion).not.toBe(main.version());
  for (const target of targets) {
    const answer = await geometry.readAnchorGeometry(target);
    expect(answer).toMatchObject({ ok: true, version: main.version(), unbuiltPages: [] });
    if (!answer.ok) throw new Error(answer.failure.message);
    expect(answer.rects.map(({ x, width }) => ({ x, width }))).toEqual(
      shownSpans(computeAnchorDisplayTarget(main, target, undefined))
    );
  }
  expect(requests).toHaveLength(targets.length);
  for (const request of requests) {
    expect(request).toMatchObject({ kind: 'anchorTarget', expectVersion: workerVersion });
  }
  expect(requests[2]).toMatchObject({ target: { kind: 'range', version: workerVersion } });
});

test('a worker viewer reports unbuilt pages its targets reach', async () => {
  const { main, targets, geometry } = await workerViewer({ unbuilt: true });
  for (const target of targets) {
    const answer = await geometry.readAnchorGeometry(target);
    expect(answer).toMatchObject({ ok: true, version: main.version(), unbuiltPages: [1] });
  }
});

test('a worker viewer read before the first paint refuses at once and succeeds on the repeated layout-change', async () => {
  const answers: Array<Promise<DocxAnchorGeometryResult> | null> = [];
  const paragraph = { kind: 'paragraph', paragraph: persisted } as const;
  const { requests, view, geometry } = await workerViewer({
    presented: false,
    onEvent(context, event) {
      if (event.type === 'layout-change' && event.layout) {
        answers.push(context.geometry?.readAnchorGeometry(paragraph) ?? null);
      }
    },
  });
  expect(await geometry.readAnchorGeometry(paragraph)).toMatchObject({
    ok: false,
    failure: { code: 'layout-unavailable' },
  });
  expect(requests).toHaveLength(0);
  const before = answers.length;
  await act(async () => view.present());
  await waitFor(() => expect(answers.length).toBeGreaterThan(before));
  expect(await answers.at(-1)).toMatchObject({ ok: true, unbuiltPages: [] });
  expect(requests).toHaveLength(1);
});

test('a worker viewer refuses stale targets and superseded worker replies', async () => {
  const { main, engine, targets, requests, geometry } = await workerViewer();
  const range = targets[2]!;
  expect(await geometry.readAnchorGeometry({ ...range, version: 'superseded' } as typeof range)).toMatchObject({
    ok: false,
    failure: { code: 'stale-version' },
  });
  expect(requests).toHaveLength(0);
  const peer = await createYrsSession({ clientId: 903 });
  sessions.push(peer);
  peer.applyUpdate(main.encodeState());
  peer.insertText({ story: 'body', paraId: '00000003', offset: 0 }, 'z');
  engine.applyUpdate(peer.encodeState());
  for (const target of targets) {
    expect(await geometry.readAnchorGeometry(target)).toMatchObject({
      ok: false,
      failure: { code: 'stale-version' },
    });
  }
  expect(requests).toHaveLength(targets.length);
});

test('a failed worker read refuses instead of placing rects', async () => {
  const { targets, requests, geometry } = await workerViewer({ fail: () => true });
  for (const [index, target] of targets.entries()) {
    expect(await geometry.readAnchorGeometry(target)).toEqual({
      ok: false,
      failure: { code: 'layout-unavailable', message: expect.any(String) },
    });
    expect(requests).toHaveLength(index + 1);
  }
});

test('a main-thread session answers readAnchorGeometry exactly like getAnchorGeometry', async () => {
  const main = await createYrsSession({ clientId: 904 });
  sessions.push(main);
  main.openDocx(await trackedDocument(), true);
  const found = main.findText({
    text: 'Beta',
    within: { kind: 'paragraph', story: 'body', paraId: '00000001' },
    view: 'accepted',
  });
  if (!found.ok) throw new Error(found.failure.message);
  const view = await host(main, queries(main.version(), 'w1', currentPreviewKey(main)));
  const geometry = await geometryOf(view.result);
  const targets: DocxGeometryTarget[] = [
    { kind: 'paragraph', paragraph: persisted },
    { kind: 'search', paragraph: persisted, text: 'xy' },
    { kind: 'range', version: main.version(), range: found.matches[0]!.range },
    { kind: 'range', version: 'superseded', range: found.matches[0]!.range },
    { kind: 'revision', revisionId: main.listRevisions()[0]!.revisionId },
  ];
  let ok = 0;
  for (const target of targets) {
    const expected = geometry.getAnchorGeometry(target);
    if (expected.ok) ok += 1;
    expect(await geometry.readAnchorGeometry(target)).toEqual(expected);
  }
  expect(ok).toBeGreaterThan(0);
});
