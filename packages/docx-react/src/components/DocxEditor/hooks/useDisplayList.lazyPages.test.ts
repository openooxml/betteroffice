import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { useRustDisplayList } from './useDisplayList';

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

/** A worker that answers from a real engine session. */
class EngineWorker {
  static engine: ReturnType<typeof createEditSession> | null = null;
  static last: EngineWorker | null = null;
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  constructor() {
    EngineWorker.last = this;
  }
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
    const engine = EngineWorker.engine!;
    let frame: Uint8Array;
    if (request.type === 'bootstrap') {
      if (request.displayWindow) engine.set_display_window(...request.displayWindow);
      frame = engine.build_display_list_frame(request.extras, 0);
    } else if (request.type === 'buildPages') {
      frame = engine.build_display_pages_frame(
        Uint32Array.from(request.pages),
        request.expectedFrameEpoch
      );
    } else {
      return;
    }
    const caret = JSON.parse(engine.resident_caret_snapshot_json());
    queueMicrotask(() =>
      this.onmessage?.({
        data: {
          id: request.id,
          ok: true,
          frame: frame.slice().buffer,
          caret,
          selection: null,
          layoutRevision: 1,
        },
      } as MessageEvent<ResidentEngineWorkerResponse>)
    );
  }
  terminate(): void {}
}

test('a worker frame builds only the pages near the viewport', async () => {
  const engine = createEditSession(9401);
  engine.create_story('body', 'Lazy pages. '.repeat(400), 'Normal', 'left');
  const fontId = engine.register_measure_font(
    new Uint8Array(
      readFileSync(
        resolve(
          import.meta.dir,
          '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
        )
      )
    )
  );
  const inputs = JSON.parse(
    engine.layout_document_with_regions_json(
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
                marginRight: 300,
                marginBottom: 300,
                marginLeft: 300,
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
      })
    )
  );
  EngineWorker.engine = engine;
  globalThis.Worker = EngineWorker as unknown as typeof Worker;
  const host = {
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host)
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const pages = () => result.current.frame!.displayList.pages;
    const first = pages();
    expect(first.length).toBeGreaterThan(5);
    expect(first.slice(0, 5).every((page) => !page.unbuilt)).toBe(true);
    expect(EngineWorker.last!.posted[0]).toMatchObject({ displayWindow: [0, 5] });

    const last = first.length - 1;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
    });
    await waitFor(() => expect(pages()[last]?.unbuilt).toBeFalsy());
    expect(pages().slice(5, last).every((page) => page.unbuilt)).toBe(true);
    const span = pages()[5]!.positionSpan!;
    expect(span[0]).toBeLessThanOrEqual(span[1]);
    expect(
      EngineWorker.last!.posted.filter((request) => request.type === 'buildPages')
    ).toEqual([expect.objectContaining({ pages: [last] })]);

    let settled: Awaited<ReturnType<typeof result.current.settledDisplayList>> | undefined;
    await act(async () => {
      settled = await result.current.settledDisplayList(() => {});
    });
    expect(settled!.pages.some((page) => page.unbuilt)).toBe(false);
    const full = JSON.parse(engine.build_display_list_json(JSON.stringify(inputs))) as {
      pages: unknown[];
    };
    expect(settled!.pages.map((page) => page.primitives.length)).toEqual(
      full.pages.map((page) => (page as { primitives: unknown[] }).primitives.length)
    );
    unmount();
  } finally {
    engine.free();
  }
});
