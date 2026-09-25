import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef, useRef, type RefObject } from 'react';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type DocxEditRequest,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { YrsInput, type YrsInputRef } from '../YrsInput';
import { createCommentIdAllocator } from '../commentFactories';
import {
  useRustDisplayList,
  type RustDisplayListHookOverrides,
  type UseRustDisplayListResult,
} from './useDisplayList';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';
import { usePagedEditorRefApi } from './usePagedEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');

const WASM = resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});
const originalWorker = globalThis.Worker;
const sessions: YrsSession[] = [];
let startWorker: () => InProcessResidentWorker;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
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

interface HarnessProps {
  session: YrsSession;
  layout: Layout;
  overrides: RustDisplayListHookOverrides;
  pagedRef: RefObject<PagedEditorRef | null>;
  docxRef: RefObject<DocxEditorRef | null>;
  display: { current: UseRustDisplayListResult | null };
}

/** The editor's input, resident-frame and ref wiring, without its canvas. */
function Harness({ session, layout, overrides, pagedRef, docxRef, display }: HarnessProps) {
  const frames = useRustDisplayList(layout, overrides, undefined, undefined, session);
  display.current = frames;
  const inputRef = useRef<YrsInputRef>(null);
  const onReadyRef = useRef<((ref: PagedEditorRef) => void) | undefined>(undefined);
  const map = () =>
    createYrsInputPositionMap(
      'body',
      session.paragraphs('body').map((paragraph) => ({
        paraId: paragraph.paraId,
        length: paragraph.text.length,
      }))
    );
  usePagedEditorRefApi({
    ref: pagedRef,
    yrsInputRef: inputRef,
    layout: null,
    runLayoutPipeline: () => {},
    getLayoutRequest: () => null,
    scrollToPositionImpl: () => {},
    scrollToParaIdImpl: () => false,
    scrollToPageImpl: () => {},
    setIsFocused: () => {},
    onReadyRef,
    documentFromYrs: () => null,
    yrsSession: session,
    yrsLocToDisplayPosition: () => null,
    syncYrsInputState: () => true,
    applyYrsFormatting: () => false,
    applyYrsCommand: () => false,
    getYrsPositionProjection: () => null,
    displayPositionToYrsLoc: () => null,
  });
  useDocxEditorRefApi({
    ref: docxRef,
    document: null,
    documentFromYrs: () => null,
    historyStateRef: { current: null },
    pagedEditorRef: pagedRef,
    handleSave: async () => null,
    handleDirectPrint: () => {},
    zoom: 1,
    setZoom: () => {},
    scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
    loadParsedDocument: () => {},
    loadBuffer: async () => {},
    comments: [],
    setComments: () => {},
    setShowCommentsSidebar: () => {},
    contentChangeSubscribersRef: { current: new Set() },
    selectionChangeSubscribersRef: { current: new Set() },
    getCachedStyleResolver: (() => {
      throw new Error('unused');
    }) as never,
    commentIdAllocator: createCommentIdAllocator(),
    modeRef: { current: 'editing' },
  });
  // Position callbacks are rebuilt on every render, as the editor's are after each new frame,
  // so every render also rebuilds the input's imperative handle.
  return (
    <YrsInput
      ref={inputRef}
      enabled
      readOnly={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
      applyResidentInput={frames.applyInput}
    />
  );
}

async function until(done: () => boolean): Promise<void> {
  const deadline = Date.now() + 5000;
  while (!done() && Date.now() < deadline) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

function texts(session: YrsSession): string[] {
  return session.paragraphs('body').map((paragraph) => paragraph.text);
}

test('applyEdits waits for worker input in flight, then refuses the batch that typing made stale', async () => {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed');
  session.registerFont(new Uint8Array(readFileSync(FONT)));
  const inputs = JSON.parse(session.layoutDocumentWithRegionsJson(LAYOUT));
  session.setSelection({ story: 'body', paraId, offset: 4 });
  let worker!: InProcessResidentWorker;
  globalThis.Worker = class {
    constructor() {
      worker = startWorker();
      return worker;
    }
  } as unknown as typeof Worker;
  const props: HarnessProps = {
    session,
    layout: inputs.layout as Layout,
    overrides: { getInputs: () => inputs },
    pagedRef: createRef<PagedEditorRef>(),
    docxRef: createRef<DocxEditorRef>(),
    display: { current: null },
  };
  const view = render(<Harness {...props} />);
  await until(() => props.display.current?.workerSurfacesActive === true);
  const request = (expectVersion: string): DocxEditRequest => ({
    expectVersion,
    steps: [
      {
        op: 'replaceText',
        target: { kind: 'search', text: 'Seed', within: { kind: 'paragraph', story: 'body', paraId }, view: 'accepted' },
        text: 'Batch',
      },
    ],
  });
  const readBeforeTyping = session.version();

  worker.hold();
  act(() => props.pagedRef.current!.insertText(' typed'));
  await until(() => worker.requests.includes('applyInput'));
  let settled = false;
  const stale = props.docxRef.current!.applyEdits(request(readBeforeTyping)).finally(() => {
    settled = true;
  });
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 20));
  });
  expect(settled).toBe(false);
  expect(texts(session)).toEqual(['Seed']);
  act(() => view.rerender(<Harness {...props} />));
  await act(async () => {
    worker.release();
    await stale.catch(() => {});
  });
  expect(await stale).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
  expect(texts(session)).toEqual(['Seed typed']);

  const forwarded = worker.requests.filter((type) => type === 'applyUpdate').length;
  const applied = await act(() => props.docxRef.current!.applyEdits(request(session.version())));
  expect(applied).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
  expect(texts(session)).toEqual(['Batch typed']);
  expect(worker.requests.filter((type) => type === 'applyUpdate')).toHaveLength(forwarded + 1);
});
