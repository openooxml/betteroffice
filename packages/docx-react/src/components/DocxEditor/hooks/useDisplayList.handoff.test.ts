import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DisplayList } from '@betteroffice/docx/layout/render';
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
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => {
        const hook = useRustDisplayList(
          layout,
          undefined,
          undefined,
          undefined,
          source,
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

    handoffFrom.current = preview.engine;
    await layOut(full, 1);
    await waitFor(() => expect(text(result.current.displayList)).toContain('Full'));

    expect(FakeWorker.created).toHaveLength(1);
    const worker = FakeWorker.created[0]!;
    expect(worker.posted.map((request) => request.type)).toEqual(['bootstrap', 'bootstrap']);
    expect(worker.posted[1]).toMatchObject({ keepSurfaces: true });
    expect(worker.posted[0]).not.toHaveProperty('keepSurfaces');
    // No render between the two documents shows an empty page list.
    expect(shown.slice(firstShown).every((list) => list !== null)).toBe(true);
    unmount();
  } finally {
    preview.native.free();
    full.native.free();
  }
});
