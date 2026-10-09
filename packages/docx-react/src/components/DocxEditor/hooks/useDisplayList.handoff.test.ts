import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DisplayList, ImageResolver } from '@betteroffice/docx/layout/render';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { useCanvasRenderer, useRustDisplayList } from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);

afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const REQUEST = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
  renderEnv: {},
});

class FakeWorker {
  static created: FakeWorker[] = [];
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  constructor() {
    FakeWorker.created.push(this);
  }
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
  }
  reply(response: ResidentEngineWorkerResponse): void {
    this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  terminate(): void {}
}

/** A host engine whose worker lays out `text`, with that layout and first frame. */
function hostWithPage(clientId: number, text: string) {
  const native = createEditSession(clientId);
  native.create_story('body', text, 'Normal', 'left');
  const layoutJson = native.layout_document_with_regions_retained_json(REQUEST);
  const frame = native.build_display_list_frame(JSON.stringify({}), 0);
  const engine = {
    adoptResidentWorkerLayout: () => 1,
    residentLayoutInWorker: () => true,
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({
      state: new Uint8Array(),
      fonts: [],
      fontsRevision: 0,
      layoutRevision: 1,
    }),
    encodeStateVector: () => new Uint8Array([1]),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
    outlineGlyphJson: () => text,
  } as unknown as YrsSession;
  return { native, layoutJson, frame, engine };
}

function text(list: DisplayList | null): string {
  return JSON.stringify(list?.pages[0]?.primitives ?? []);
}

test('a session handed over keeps its worker and shows the old pages until the new ones', async () => {
  FakeWorker.created = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const preview = hostWithPage(9510, 'Preview page');
  const full = hostWithPage(9511, 'Full page');
  const handoffFrom = { current: null as YrsSession | null };
  const shown: Array<DisplayList | null> = [];
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { preloadLayoutWasm } = await import('@betteroffice/docx/wasm/layout');
    await preloadLayoutWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm')
        )
      )
    );
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => {
        const hook = useRustDisplayList(
          layout,
          undefined,
          undefined,
          undefined,
          source,
          undefined,
          undefined,
          handoffFrom
        );
        shown.push(hook.displayList);
        return hook;
      },
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const layOut = async (host: ReturnType<typeof hostWithPage>, index: number) => {
      const pending = result.current.layoutInWorker(host.engine, REQUEST);
      const worker = FakeWorker.created[0]!;
      worker.reply({
        id: worker.posted[index]!.id,
        ok: true,
        frame: host.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: null,
        layoutRevision: 1,
        layoutJson: host.layoutJson,
      });
      const computation = await pending!;
      await act(async () => {
        rerender({ layout: computation!.layout, source: host.engine });
      });
    };

    await layOut(preview, 0);
    await waitFor(() => expect(text(result.current.displayList)).toContain('Preview'));
    const firstShown = shown.findIndex((list) => list !== null);
    const stale = result.current.queries!;
    await stale.whenReady();
    stale.prime();
    const previewRects = stale.rangeRects(1, 2);
    const previewHits = stale.hitTestRegions(0, 100, 100);
    expect(previewRects.length).toBeGreaterThan(0);

    handoffFrom.current = preview.engine;
    await layOut(full, 1);
    await waitFor(() => expect(text(result.current.displayList)).toContain('Full'));
    const live = result.current.queries!;
    await live.whenReady();
    live.prime();
    expect(live).not.toBe(stale);
    expect(text(stale.displayList)).toContain('Preview');
    expect(text(live.displayList)).toContain('Full');
    expect(live.rangeRects(1, 2).length).toBeGreaterThan(0);

    // A line is one session's: the preview's facade keeps its handle and answers from its own pages.
    warnings.mockClear();
    expect(stale.rangeRects(1, 2)).toEqual(previewRects);
    expect(stale.hitTestRegions(0, 100, 100)).toEqual(previewHits);
    expect(warnings).not.toHaveBeenCalled();

    // What hosts read from here on (search, geometry, plugin queries, hit tests) is the full layout's.
    await act(async () => {
      await new Promise((done) => setTimeout(done, 50));
    });
    const current = result.current.queries!;
    expect(current).toBe(live);
    expect(text(current.displayList)).toContain('Full');
    expect(text(current.displayList)).not.toContain('Preview');
    expect(text(result.current.displayList)).not.toContain('Preview');
    expect(text(shown.at(-1)!)).toContain('Full');

    expect(FakeWorker.created).toHaveLength(1);
    const worker = FakeWorker.created[0]!;
    expect(worker.posted.map((request) => request.type)).toEqual(['bootstrap', 'bootstrap']);
    expect(worker.posted[1]).toMatchObject({ keepSurfaces: true });
    expect(worker.posted[0]).not.toHaveProperty('keepSurfaces');
    // No render between the two documents shows an empty page list.
    expect(shown.slice(firstShown).every((list) => list !== null)).toBe(true);
    stale.dispose();
    live.dispose();
    unmount();
  } finally {
    warnings.mockRestore();
    preview.native.free();
    full.native.free();
  }
});

test('a session handed over never takes the worker back from its successor', async () => {
  FakeWorker.created = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const preview = hostWithPage(9516, 'Preview page');
  const full = hostWithPage(9517, 'Full page');
  const handoffFrom = { current: null as YrsSession | null };
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source, resolved }) =>
        useRustDisplayList(
          layout,
          undefined,
          undefined,
          resolved,
          source,
          undefined,
          undefined,
          handoffFrom
        ),
      {
        initialProps: {
          layout: null as Layout | null,
          source: null as YrsSession | null,
          resolved: undefined as ReadonlySet<number> | undefined,
        },
      }
    );
    const reply = (host: ReturnType<typeof hostWithPage>, index: number) => {
      const worker = FakeWorker.created[0]!;
      worker.reply({
        id: worker.posted[index]!.id,
        ok: true,
        frame: host.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: null,
        layoutRevision: 1,
        layoutJson: host.layoutJson,
      });
    };
    const previewPending = result.current.layoutInWorker(preview.engine, REQUEST);
    reply(preview, 0);
    const previewLayout = (await previewPending!)!.layout;
    await act(async () => {
      rerender({ layout: previewLayout, source: preview.engine, resolved: undefined });
    });
    await waitFor(() => expect(text(result.current.displayList)).toContain('Preview'));

    handoffFrom.current = preview.engine;
    const fullPending = result.current.layoutInWorker(full.engine, REQUEST);
    // A preview redraw after the handover, such as a changed set of resolved comments.
    await act(async () => {
      rerender({ layout: previewLayout, source: preview.engine, resolved: new Set([1]) });
    });
    expect(result.current.layoutInWorker(preview.engine, REQUEST)).toBeNull();
    expect(FakeWorker.created).toHaveLength(1);
    const worker = FakeWorker.created[0]!;
    expect(worker.posted.map((request) => request.type)).toEqual(['bootstrap', 'bootstrap']);
    expect(text(result.current.displayList)).toContain('Preview');

    reply(full, 1);
    const fullLayout = (await fullPending!)!.layout;
    await act(async () => {
      rerender({ layout: fullLayout, source: full.engine, resolved: undefined });
    });
    await waitFor(() => expect(text(result.current.displayList)).toContain('Full'));
    expect(FakeWorker.created).toHaveLength(1);
    expect(worker.posted.map((request) => request.type)).not.toContain('destroy');
    unmount();
  } finally {
    preview.native.free();
    full.native.free();
  }
});

test('glyph outlines and decoded images come from the session whose pages are shown', async () => {
  FakeWorker.created = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const preview = hostWithPage(9520, 'Preview page');
  const full = hostWithPage(9521, 'Full page');
  const handoffFrom = { current: null as YrsSession | null };
  const shown: Array<[string, string | undefined, ImageResolver]> = [];
  try {
    const { result, unmount } = renderHook(() => {
      const renderer = useCanvasRenderer(undefined, undefined, undefined, undefined, handoffFrom);
      shown.push([
        text(renderer.displayList),
        (renderer.glyphOutlineProvider as ((json: string) => string) | null)?.(''),
        renderer.resolveImage,
      ]);
      return renderer;
    });
    const layOut = async (host: ReturnType<typeof hostWithPage>, index: number) => {
      const pending = result.current.layoutInWorker(host.engine, REQUEST);
      const worker = FakeWorker.created[0]!;
      worker.reply({
        id: worker.posted[index]!.id,
        ok: true,
        frame: host.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: null,
        layoutRevision: 1,
        layoutJson: host.layoutJson,
      });
      const computation = await pending!;
      await act(async () => {
        result.current.onLayoutComputed(
          computation!.layout,
          host.engine as unknown as Parameters<typeof result.current.onLayoutComputed>[1]
        );
      });
    };
    await layOut(preview, 0);
    await waitFor(() => expect(text(result.current.displayList)).toContain('Preview'));
    handoffFrom.current = preview.engine;
    await layOut(full, 1);
    await waitFor(() => expect(text(result.current.displayList)).toContain('Full'));
    const previewImages = shown.find(([pages]) => pages.includes('Preview'))![2];
    for (const [pages, outlines, images] of shown) {
      if (pages.includes('Preview')) {
        expect(outlines).toBe('Preview page');
        expect(images).toBe(previewImages);
      }
      if (pages.includes('Full')) {
        expect(outlines).toBe('Full page');
        expect(images).not.toBe(previewImages);
      }
    }
    unmount();
  } finally {
    preview.native.free();
    full.native.free();
  }
});

test('a request of the preview failing after the handover leaves the new session its pages', async () => {
  FakeWorker.created = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const preview = hostWithPage(9512, 'Preview page');
  const full = hostWithPage(9513, 'Full page');
  const handoffFrom = { current: null as YrsSession | null };
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) =>
        useRustDisplayList(
          layout,
          undefined,
          undefined,
          undefined,
          source,
          undefined,
          undefined,
          handoffFrom
        ),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const layOut = async (
      host: ReturnType<typeof hostWithPage>,
      index: number,
      provisional: boolean
    ) => {
      const pending = result.current.layoutInWorker(host.engine, REQUEST);
      const worker = FakeWorker.created[0]!;
      worker.reply({
        id: worker.posted[index]!.id,
        ok: true,
        frame: host.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: null,
        layoutRevision: 1,
        layoutJson: host.layoutJson,
        ...(provisional ? { layoutProvisional: true } : {}),
      });
      const computation = await pending!;
      await act(async () => {
        rerender({ layout: computation!.layout, source: host.engine });
      });
    };

    await layOut(preview, 0, true);
    const worker = FakeWorker.created[0]!;
    await waitFor(() =>
      expect(worker.posted.map((request) => request.type)).toContain('completeLayout')
    );
    handoffFrom.current = preview.engine;
    await layOut(full, 2, false);
    await waitFor(() => expect(text(result.current.displayList)).toContain('Full'));
    expect(result.current.workerSurfacesActive).toBe(true);

    await act(async () => {
      worker.reply({ id: worker.posted[1]!.id, ok: false, error: 'late' });
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    expect(text(result.current.displayList)).toContain('Full');
    expect(result.current.frame).not.toBeNull();
    expect(result.current.workerSurfacesActive).toBe(true);
    unmount();
  } finally {
    preview.native.free();
    full.native.free();
  }
});

test('input for a session the worker does not serve takes the host path', async () => {
  FakeWorker.created = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const preview = hostWithPage(9514, 'Preview page');
  const at = { story: 'body', paraId: '00000001', offset: 0 };
  Object.assign(preview.engine, { selection: () => ({ anchor: at, head: at }) });
  const handoffFrom = { current: null as YrsSession | null };
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) =>
        useRustDisplayList(
          layout,
          undefined,
          undefined,
          undefined,
          source,
          undefined,
          undefined,
          handoffFrom
        ),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const pending = result.current.layoutInWorker(preview.engine, REQUEST);
    const worker = FakeWorker.created[0]!;
    worker.reply({
      id: worker.posted[0]!.id,
      ok: true,
      frame: preview.frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson: preview.layoutJson,
    });
    const computation = await pending!;
    await act(async () => {
      rerender({ layout: computation!.layout, source: preview.engine });
    });
    await waitFor(() => expect(text(result.current.displayList)).toContain('Preview'));

    // The shown session is no longer the one the worker holds.
    await act(async () => {
      rerender({ layout: computation!.layout, source: {} as YrsSession });
    });
    const outcome = await Promise.race([
      result.current.applyInput('x'),
      new Promise((resolve) => setTimeout(() => resolve('pending'), 200)),
    ]);
    expect(outcome).toBeNull();
    expect(worker.posted.map((request) => request.type)).not.toContain('applyInput');
    unmount();
  } finally {
    preview.native.free();
  }
});

test('a display-only preview never asks the worker for the rest of its layout', async () => {
  FakeWorker.created = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const preview = hostWithPage(9515, 'Preview page');
  Object.assign(preview.engine, { isDisplayOnly: () => true });
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) =>
        useRustDisplayList(layout, undefined, undefined, undefined, source, undefined, undefined, {
          current: null,
        }),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const pending = result.current.layoutInWorker(preview.engine, REQUEST);
    const worker = FakeWorker.created[0]!;
    worker.reply({
      id: worker.posted[0]!.id,
      ok: true,
      frame: preview.frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson: preview.layoutJson,
      layoutProvisional: true,
    });
    const computation = await pending!;
    expect(computation!.complete).toBeUndefined();
    await act(async () => {
      rerender({ layout: computation!.layout, source: preview.engine });
      await new Promise((resolve) => setTimeout(resolve, 300));
    });
    expect(worker.posted.map((request) => request.type)).not.toContain('completeLayout');
    unmount();
  } finally {
    preview.native.free();
  }
});

test('releasing lets go of the worker and the engine the pages showed', async () => {
  FakeWorker.created = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const preview = hostWithPage(9520, 'Preview page');
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) =>
        useRustDisplayList(layout, undefined, undefined, undefined, source),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const pending = result.current.layoutInWorker(preview.engine, REQUEST);
    const worker = FakeWorker.created[0]!;
    worker.reply({
      id: worker.posted[0]!.id,
      ok: true,
      frame: preview.frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson: preview.layoutJson,
    });
    const computation = await pending!;
    await act(async () => {
      rerender({ layout: computation!.layout, source: preview.engine });
    });
    await waitFor(() => expect(result.current.presentedEngine).toBe(preview.engine));

    await act(async () => {
      result.current.release();
      rerender({ layout: null, source: null });
    });
    expect(worker.posted.at(-1)?.type).toBe('destroy');
    expect(result.current.presentedEngine).toBeNull();
    expect(result.current.displayList).toBeNull();
    unmount();
  } finally {
    preview.native.free();
  }
});

const lazyRequest = (fontId: number) =>
  JSON.stringify({
    bodyStory: 'body',
    regions: {
      sections: [
        {
          sectionId: 'main',
          properties: {
            pageWidth: 4320,
            pageHeight: 2880,
            marginTop: 300,
            marginBottom: 300,
            marginLeft: 300,
            marginRight: 300,
          },
        },
      ],
    },
    measurement: {
      fontChains: { 'calibri|0|0': [fontId] },
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      authoritativeShaping: true,
    },
    renderEnv: {},
  });

/** Like `hostWithPage`, over several pages of which only the first is built. */
function hostWithLazyPages(clientId: number, text: string) {
  const native = createEditSession(clientId);
  native.create_story('body', `${text} `.repeat(400), 'Normal', 'left');
  const fontId = native.register_measure_font(
    new Uint8Array(
      readFileSync(
        resolve(
          import.meta.dir,
          '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
        )
      )
    )
  );
  native.set_display_window(0, 1);
  const request = lazyRequest(fontId);
  const layoutJson = native.layout_document_with_regions_retained_json(request);
  const frame = native.build_display_list_frame(JSON.stringify({}), 0);
  const host = hostWithPage(clientId, text);
  host.native.free();
  return { native, layoutJson, frame, engine: host.engine, request };
}

test("a worker handed to another session builds no pages of the old session's frame", async () => {
  FakeWorker.created = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const preview = hostWithLazyPages(9530, 'Preview');
  const full = hostWithLazyPages(9531, 'Full');
  const handoffFrom = { current: null as YrsSession | null };
  const reply = (
    worker: FakeWorker,
    request: ResidentEngineWorkerRequest,
    host: ReturnType<typeof hostWithLazyPages>,
    frame: Uint8Array
  ) =>
    worker.reply({
      id: request.id,
      ok: true,
      frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson: host.layoutJson,
    });
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) =>
        useRustDisplayList(
          layout,
          undefined,
          undefined,
          undefined,
          source,
          undefined,
          undefined,
          handoffFrom
        ),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const first = result.current.layoutInWorker(preview.engine, preview.request)!;
    const worker = FakeWorker.created[0]!;
    reply(worker, worker.posted[0]!, preview, preview.frame);
    const previewLayout = await first;
    await act(async () => {
      rerender({ layout: previewLayout!.layout, source: preview.engine });
    });
    const last = result.current.displayList!.pages.length - 1;
    expect(result.current.displayList!.pages[last]!.unbuilt).toBe(true);

    handoffFrom.current = preview.engine;
    const second = result.current.layoutInWorker(full.engine, full.request)!;
    const bootstrap = worker.posted.at(-1)!;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
      await new Promise((done) => setTimeout(done, 50));
    });
    expect(worker.posted.map((request) => request.type)).not.toContain('buildPages');

    reply(worker, bootstrap, full, full.frame);
    const fullLayout = await second;
    await act(async () => {
      rerender({ layout: fullLayout!.layout, source: full.engine });
    });
    await waitFor(() => expect(worker.posted.at(-1)!.type).toBe('buildPages'));
    const build = worker.posted.at(-1)! as ResidentEngineWorkerRequest & {
      pages: number[];
      expectedFrameEpoch: number;
    };
    expect(build.expectedFrameEpoch).toBe(result.current.frame!.frameEpoch);
    await act(async () => {
      reply(
        worker,
        build,
        full,
        full.native.build_display_pages_frame(
          Uint32Array.from(build.pages),
          build.expectedFrameEpoch
        )
      );
    });
    await waitFor(() =>
      expect(build.pages.every((index) => !result.current.displayList!.pages[index]!.unbuilt)).toBe(
        true
      )
    );
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    preview.native.free();
    full.native.free();
  }
});

test.each([false, true])(
  'a provisional preview handoff builds every full-document page with worker-open=%s',
  async (experimentalWorkerOpen) => {
    FakeWorker.created = [];
    globalThis.Worker = FakeWorker as unknown as typeof Worker;
    const preview = hostWithLazyPages(9532, 'Preview');
    const full = hostWithLazyPages(9533, 'Full document');
    const handoffFrom = { current: null as YrsSession | null };
    try {
      const { result, rerender, unmount } = renderHook(
        ({ layout, source }) =>
          useRustDisplayList(
            layout,
            undefined,
            undefined,
            undefined,
            source,
            undefined,
            undefined,
            handoffFrom,
            experimentalWorkerOpen
          ),
        { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
      );
      const first = result.current.layoutInWorker(preview.engine, preview.request)!;
      const worker = FakeWorker.created[0]!;
      let host = preview;
      const reply = (request: ResidentEngineWorkerRequest, frame: Uint8Array) =>
        worker.reply({
          id: request.id,
          ok: true,
          frame: frame.slice().buffer,
          caret: JSON.parse(host.native.resident_caret_snapshot_json()),
          selection: null,
          layoutRevision: 1,
          layoutJson: host.layoutJson,
          ...(host === preview ? { layoutProvisional: true } : {}),
        });
      worker.postMessage = (request) => {
        worker.posted.push(request);
        if (request.type === 'buildPages') {
          const pages = host === preview ? request.pages.filter((index) => index < 1) : request.pages;
          const frame = host.native.build_display_pages_frame(
            Uint32Array.from(pages),
            request.expectedFrameEpoch
          );
          queueMicrotask(() => reply(request, frame));
        }
      };
      reply(worker.posted[0]!, preview.frame);
      const previewLayout = await first;
      await act(async () => {
        rerender({ layout: previewLayout!.layout, source: preview.engine });
        result.current.setRetainBuiltPages!(true);
      });
      const previewFrame = result.current.frame!;
      await act(async () => {
        result.current.setDisplayWindow(1, previewFrame.displayList.pages.length);
      });
      await waitFor(() => expect(result.current.frame).not.toBe(previewFrame));

      handoffFrom.current = preview.engine;
      host = full;
      const second = result.current.layoutInWorker(full.engine, full.request)!;
      full.native.reset_frame_base();
      reply(
        worker.posted.at(-1)!,
        full.native.build_display_pages_frame(new Uint32Array(), result.current.frame!.frameEpoch)
      );
      const fullLayout = await second;
      await act(async () => {
        rerender({ layout: fullLayout!.layout, source: full.engine });
      });
      await waitFor(() => expect(result.current.presentedEngine).toBe(full.engine));
      await act(async () => {
        const settled = result.current.settledDisplayList(null, 1_000);
        await expect(settled).resolves.toBeDefined();
        expect((await settled).pages.every((page) => !page.unbuilt)).toBe(true);
      });
      unmount();
    } finally {
      preview.native.free();
      full.native.free();
    }
  }
);

test("a display-only preview's failed build fails no wait for the document", async () => {
  const failure = new Error('preview build failed');
  const overrides = {
    build: async () => {
      throw failure;
    },
    getInputs: () => ({ measured: [], options: {} }) as never,
  };
  const layout = { pageSize: { w: 816, h: 1056 }, pages: [] } as unknown as Layout;
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    for (const displayOnly of [true, false]) {
      const engine = { isDisplayOnly: () => displayOnly } as unknown as YrsSession;
      const { result, unmount } = renderHook(() =>
        useRustDisplayList(layout, overrides, undefined, undefined, engine)
      );
      let failed = false;
      void result.current.settledDisplayList(null, null).catch(() => (failed = true));
      await waitFor(() => expect(result.current.error).toBe(failure));
      // The editor fails the load only for its full session's own errors.
      expect(result.current.errorEngine).toBe(engine);
      await act(async () => {});
      expect(failed).toBe(!displayOnly);
      unmount();
    }
  } finally {
    errors.mockRestore();
  }
});
