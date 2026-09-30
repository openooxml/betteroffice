import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useCallback, useRef, useState } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsDocxHost, type YrsSession } from '@betteroffice/docx/yrs';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import type { ResidentEngineWorkerRequest } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { LayoutSelectionGate } from '@betteroffice/docx/layout';
import { useCanvasRenderer, type OpenInWorker } from './useDisplayList';
import { useLayoutPipeline } from './useLayoutPipeline';
import { useYrsCoreSession } from './useYrsCoreSession';
import type { DocxEditorCollaborationOptions } from '../types';
import { awaitWorkerOpenReplica, ensureWorkerOpenReplica } from '../internals/workerOpenReplica';
import { sourceVersionOf } from '../internals/layoutProvenance';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;
const bytes = new Uint8Array(readFileSync(resolve(
  import.meta.dir,
  '../../../../../../crates/docx-edit/tests/fixtures/page-fragments/pages.docx'
)));
const font = new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
)));
const sessions: YrsSession[] = [];
let startWorker!: () => InProcessResidentWorker;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});
afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function installWorker(options: { failOpen?: boolean; failState?: boolean; holdState?: boolean; holdOpen?: boolean } = {}) {
  const workers: InProcessResidentWorker[] = [];
  const posted: ResidentEngineWorkerRequest[] = [];
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      const send = worker.postMessage.bind(worker);
      worker.postMessage = (request, transfer) => {
        posted.push(request);
        if ((options.holdState && request.type === 'encodeState') ||
            (options.holdOpen && request.type === 'open')) worker.hold();
        if ((options.failOpen && request.type === 'open') ||
            (options.failState && request.type === 'encodeState')) {
          queueMicrotask(() => worker.onmessage?.({
            data: { id: request.id, ok: false, error: 'open failed', terminal: true },
          } as MessageEvent));
        } else send(request, transfer);
      };
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
  return { workers, posted };
}

interface HarnessProps {
  experimentalWorkerOpen: boolean;
  source: Uint8Array;
  generation: number;
  collaboration?: DocxEditorCollaborationOptions;
}

function useHarness(props: HarnessProps) {
  const renderer = useCanvasRenderer();
  const [host, setHost] = useState<YrsDocxHost | null>(null);
  const mainOpens = useRef<boolean[]>([]);
  const errors = useRef<Error[]>([]);
  const openInWorker = useCallback<OpenInWorker>((session, source, digest, generation) => {
    const open = session.openDocx.bind(session);
    session.openDocx = (input, seed, options) => {
      mainOpens.current.push(seed);
      return open(input, seed, options);
    };
    return renderer.openInWorker(session, source, digest, generation);
  }, [renderer.openInWorker]);
  const core = useYrsCoreSession(
    true, host?.document ?? null, null, props.source, props.generation, props.collaboration,
    {
      onHostDocument: setHost,
      onError: (error) => errors.current.push(error),
    },
    props.experimentalWorkerOpen ? {
      openInWorker,
      renderedFrame: renderer.status === 'ready' ? renderer.displayList : null,
    } : undefined
  );
  const syncCoordinator = useRef(new LayoutSelectionGate());
  const element = useRef<HTMLDivElement | null>(null);
  const registeredFont = useRef<{ session: YrsSession; id: number } | null>(null);
  const pipeline = useLayoutPipeline({
    document: host?.document ?? null,
    session: core.session,
    renderEnv: {},
    pageGap: 24,
    zoom: 1,
    residentMeasurementConfig: (requirements) => {
      const session = core.session;
      if (!session) return null;
      if (registeredFont.current?.session !== session) {
        registeredFont.current = { session, id: session.registerFont(font) };
      }
      const id = registeredFont.current.id;
      return {
        fontChains: Object.fromEntries(requirements.map((requirement) => [requirement.key, [id]])),
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: { noLeading: false, doNotExpandShiftReturn: false },
        authoritativeShaping: true,
      };
    },
    deferLayoutPass: () => false,
    pagesContainerRef: element,
    viewportLayoutRef: element,
    syncCoordinator: syncCoordinator.current,
    getScrollContainer: () => null,
    onLayoutComputed: (layout) => renderer.onLayoutComputed(layout, core.session),
    layoutInWorker: renderer.layoutInWorker,
    fontRequirementsInWorker: props.experimentalWorkerOpen ? renderer.fontRequirementsInWorker : undefined,
    onError: (error) => errors.current.push(error),
  });
  return { core, renderer, pipeline, host, mainOpens: mainOpens.current, errors: errors.current };
}

const initialProps: HarnessProps = { experimentalWorkerOpen: true, source: bytes, generation: 1 };

function texts(session: YrsSession) {
  return Object.fromEntries(session.storyIds().sort().map((story) => [
    story, session.paragraphs(story).map((paragraph) => paragraph.text),
  ]));
}

test('the default open calls no worker open, font preflight or state handoff', async () => {
  const { workers, posted } = installWorker();
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, experimentalWorkerOpen: false },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.core.replicaReady).toBe(true);
  expect(workers).toHaveLength(0);
  expect(posted).toHaveLength(0);
});

test('worker font preflight and the first frame precede the save-capable main replica', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const replicas: Array<YrsSession | null> = [];
  const { result } = renderHook(useHarness, {
    initialProps: {
      ...initialProps,
      collaboration: { clientId: 9401, onReplica: (replica) => replicas.push(replica as YrsSession | null) },
    },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  const session = result.current.core.session!;
  expect(session.clientId).toBe(9401);
  expect(session.storyIds()).toEqual([]);
  expect(result.current.mainOpens).toEqual([]);
  expect(replicas).toEqual([]);
  act(() => result.current.pipeline.runLayoutPipeline());
  await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
  await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
  expect(result.current.errors).toEqual([]);
  expect(posted.map((request) => request.type).slice(0, 3)).toEqual(['open', 'fontRequirements', 'bootstrap']);
  expect(posted.find((request) => request.type === 'bootstrap')).toMatchObject({ opened: true });
  expect(result.current.renderer.frame).not.toBeNull();
  expect(result.current.mainOpens).toEqual([]);
  expect(result.current.core.replicaReady).toBe(false);
  expect(session.storyIds()).toEqual([]);

  await act(async () => {
    workers[0].release();
    await awaitWorkerOpenReplica(session);
  });
  await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
  expect(result.current.mainOpens).toEqual([false]);
  expect(replicas).toEqual([session]);
  expect(sourceVersionOf(result.current.renderer.displayList)).toBe(session.version());
  const direct = await createYrsSession();
  sessions.push(direct);
  direct.openDocx(bytes, true);
  expect(texts(session)).toEqual(texts(direct));
  expect(result.current.core.documentFromYrs()).not.toBeNull();
});

test('a failed worker open falls back to the existing main open', async () => {
  const { posted } = installWorker({ failOpen: true });
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(result.current.mainOpens).toEqual([true]);
  expect(result.current.core.replicaReady).toBe(true);
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.errors).toEqual([]);
  expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
});

test('an unavailable worker falls back before publishing host metadata', async () => {
  globalThis.Worker = undefined as unknown as typeof Worker;
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(result.current.mainOpens).toEqual([true]);
  expect(result.current.core.replicaReady).toBe(true);
  expect(result.current.errors).toEqual([]);
});

test('a worker lost during the handoff opens a full main replica', async () => {
  installWorker({ failState: true });
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  act(() => result.current.pipeline.runLayoutPipeline());
  await waitFor(() => expect(result.current.mainOpens).toEqual([true]));
  await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.errors).toEqual([]);
});

test('a failed fallback reports the same document error as a normal open', async () => {
  const invalid = Uint8Array.of(1, 2, 3);
  const direct = await createYrsSession();
  sessions.push(direct);
  let expected = '';
  try { direct.openDocx(invalid, true); }
  catch (error) { expected = error instanceof Error ? error.message : String(error); }
  expect(expected).not.toBe('');
  installWorker({ failOpen: true });
  const { result } = renderHook(useHarness, { initialProps: { ...initialProps, source: invalid } });
  await waitFor(() => expect(result.current.errors).toHaveLength(1));
  expect(result.current.errors[0].message).toBe(expected);
  expect(result.current.core.session).toBeNull();
});

test('a shared collaboration update keeps the existing join path', async () => {
  const shared = await createYrsSession();
  sessions.push(shared);
  shared.openDocx(bytes, true);
  const { workers } = installWorker();
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, collaboration: { initialUpdate: shared.encodeState() } },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(workers).toHaveLength(0);
  expect(result.current.core.replicaReady).toBe(true);
  expect(texts(result.current.core.session!)).toEqual(texts(shared));
});

test('main-thread layout fallback opens the pending replica before measuring it', async () => {
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  act(() => result.current.pipeline.runLayoutPipeline({ onHost: true }));
  await waitFor(() => expect(result.current.pipeline.layout).not.toBeNull());
  expect(result.current.mainOpens).toEqual([true]);
  expect(result.current.core.replicaReady).toBe(true);
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.errors).toEqual([]);
});

test('a replaced worker open never publishes its host or revives its replica', async () => {
  const { workers, posted } = installWorker({ holdOpen: true });
  const { result, rerender } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(workers).toHaveLength(1));
  act(() => rerender({ ...initialProps, generation: 2 }));
  await waitFor(() => expect(workers).toHaveLength(2));
  await act(async () => {
    workers[0].release();
    workers[1].release();
  });
  await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
  expect(posted.filter((request) => request.type === 'open').map((request) => request.generation)).toEqual(['1', '2']);
  expect(result.current.mainOpens).toEqual([]);
  expect(result.current.errors).toEqual([]);
  act(() => ensureWorkerOpenReplica(result.current.core.session!));
});
