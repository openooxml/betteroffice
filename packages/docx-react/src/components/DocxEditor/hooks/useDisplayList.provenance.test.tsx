import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DisplayList } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsRenderEnv, type YrsSession } from '@betteroffice/docx/yrs';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import {
  revisionPreviewKey,
  revisionPreviewKeyOf,
  sourceVersionOf,
  stampRevisionPreviewKey,
  stampSourceVersion,
} from '../internals/layoutProvenance';
import {
  useRustDisplayList,
  type RustDisplayListHookOverrides,
  type UseRustDisplayListResult,
} from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');

const WASM = resolve(
  import.meta.dir,
  '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
);
const FONT = resolve(
  import.meta.dir,
  '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const layoutRequest = (renderEnv: YrsRenderEnv) =>
  JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
    renderEnv,
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

function Harness({
  session,
  layout,
  overrides,
  display,
  requestLayout,
  previewKeys,
}: {
  session: YrsSession;
  layout: Layout;
  overrides: RustDisplayListHookOverrides;
  display: { current: UseRustDisplayListResult | null };
  requestLayout: () => void;
  previewKeys: [string | null, string][];
}) {
  display.current = useRustDisplayList(
    layout,
    overrides,
    undefined,
    undefined,
    session,
    requestLayout
  );
  previewKeys.push([
    revisionPreviewKeyOf(display.current.queries),
    paintedText(display.current.displayList),
  ]);
  return null;
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

async function setup(renderEnv: { current: YrsRenderEnv } = { current: {} }) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed');
  session.registerFont(new Uint8Array(readFileSync(FONT)));
  const layOut = () => {
    const inputs = JSON.parse(session.layoutDocumentWithRegionsJson(layoutRequest(renderEnv.current)));
    stampSourceVersion(inputs.layout, session.version());
    stampRevisionPreviewKey(inputs.layout, revisionPreviewKey(renderEnv.current.revisionPreview));
    return inputs as { layout: Layout };
  };
  let inputs = layOut();
  session.setSelection({ story: 'body', paraId, offset: 4 });
  let worker!: InProcessResidentWorker;
  globalThis.Worker = class {
    constructor() {
      worker = startWorker();
      return worker;
    }
  } as unknown as typeof Worker;
  const display: { current: UseRustDisplayListResult | null } = { current: null };
  const layoutRequests: number[] = [];
  const previewKeys: [string | null, string][] = [];
  const overrides: RustDisplayListHookOverrides = { getInputs: () => inputs as never };
  const harness = () => (
    <Harness
      session={session}
      layout={inputs.layout}
      overrides={overrides}
      display={display}
      requestLayout={() => layoutRequests.push(Date.now())}
      previewKeys={previewKeys}
    />
  );
  const view = render(harness());
  await until(() => display.current?.workerSurfacesActive === true && !!display.current.queries);
  const relayout = () => {
    inputs = layOut();
    view.rerender(harness());
  };
  return {
    session,
    paraId,
    display,
    worker: () => worker,
    layoutRequests,
    previewKeys,
    relayout,
  };
}

test('a frame is stamped with the version its layout and typed input produced', async () => {
  const { session, display } = await setup();
  expect(sourceVersionOf(display.current!.queries)).toBe(session.version());
  await act(async () => {
    await display.current!.applyInput('!');
  });
  expect(session.paragraphs('body')[0].text).toBe('Seed!');
  expect(sourceVersionOf(display.current!.queries)).toBe(session.version());
});

test('a worker frame overtaken by a remote change stays unpublished and unsettled until a fresh layout', async () => {
  const { session, display, worker, layoutRequests, relayout } = await setup();
  const replica = await createYrsSession({ clientId: 7777 });
  sessions.push(replica);
  replica.applyUpdate(session.encodeStateAsUpdate());
  worker().hold();
  let pending!: Promise<unknown>;
  act(() => {
    pending = display.current!.applyInput('?');
  });
  await until(() => worker().requests.includes('applyInput'));
  const [paragraph] = replica.paragraphs('body');
  replica.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Remote ');
  await act(async () => {
    session.applyUpdate(replica.encodeStateAsUpdate(session.encodeStateVector()));
  });
  let settled = false;
  void display
    .current!.settledDisplayList(() => {})
    .then(() => {
      settled = true;
    });
  await act(async () => {
    worker().release();
    await pending;
  });
  expect(session.paragraphs('body')[0].text).toBe('Remote Seed?');
  expect(display.current!.displayList).not.toBeNull();
  expect(display.current!.queries).toBeNull();
  await until(() => layoutRequests.length > 0);
  expect(settled).toBe(false);

  await act(async () => relayout());
  await until(() => settled && sourceVersionOf(display.current!.queries) === session.version());
});

/** Painted text, each tracked stretch followed by its revision kind. */
function paintedText(displayList: DisplayList | null): string {
  const segments: [string, string][] = [];
  for (const primitive of (displayList?.pages ?? []).flatMap((page) => page.primitives)) {
    if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun') continue;
    const kind = primitive.revision?.kind ?? '';
    const last = segments.at(-1);
    if (last && last[1] === kind) last[0] += primitive.text;
    else segments.push([primitive.text, kind]);
  }
  return segments.map(([text, kind]) => (kind ? `${text}[${kind}]` : text)).join('');
}

/** "Seed" with a suggested " more", shown through a preview the test switches by relayout. */
async function previewSetup() {
  const renderEnv: { current: YrsRenderEnv } = { current: {} };
  const harness = await setup(renderEnv);
  const { session, paraId, display, relayout } = harness;
  const applied = session.applyEdits({
    expectVersion: session.version(),
    history: 'none',
    steps: [
      {
        op: 'insertText',
        target: { kind: 'paragraph', story: 'body', paraId },
        at: 'end',
        text: ' more',
        suggest: { author: 'Ann', date: '2026-09-29T12:00:00Z' },
      },
    ],
  });
  if (!applied.ok) throw new Error(applied.failure.message);
  const [revision] = applied.receipts[0].revisionIds;
  const key = (decision?: 'accepted' | 'rejected') =>
    revisionPreviewKey(decision ? { [revision]: decision } : undefined);
  const show = async (decision?: 'accepted' | 'rejected') => {
    renderEnv.current = decision ? { revisionPreview: { [revision]: decision } } : {};
    await act(async () => relayout());
  };
  const shown = () => [
    revisionPreviewKeyOf(display.current!.queries),
    paintedText(display.current!.displayList),
  ];
  await show('accepted');
  await until(() => shown()[0] === key('accepted'));
  expect(shown()).toEqual([key('accepted'), 'Seed more']);
  return { ...harness, revision, key, show, shown };
}

test('a frame carries the revision preview of its layout, typed frames included', async () => {
  const { display, previewKeys, worker, key, show, shown } = await previewSetup();
  const acceptedQueries = display.current!.queries;
  await act(async () => {
    await display.current!.applyInput('!');
  });
  expect(shown()).toEqual([key('accepted'), 'Seed more!']);

  worker().hold();
  await show('rejected');
  await show();
  const released = previewKeys.length;
  await act(async () => worker().release());
  await until(() => shown()[0] === key());
  expect(shown()).toEqual([key(), 'Seed more[ins]!']);
  expect(previewKeys.slice(released).map(([shownKey]) => shownKey)).not.toContain(key('rejected'));
  expect(revisionPreviewKeyOf(acceptedQueries)).toBe(key('accepted'));
  expect(revisionPreviewKey({ b: 'rejected', a: 'accepted' })).toBe(
    revisionPreviewKey({ a: 'accepted', b: 'rejected' })
  );
});

test('a typed frame from the previous preview never publishes under the new one', async () => {
  const { display, previewKeys, worker, key, show, shown } = await previewSetup();
  worker().hold();
  let typed!: Promise<unknown>;
  act(() => {
    typed = display.current!.applyInput('!');
  });
  await until(() => worker().requests.includes('applyInput'));
  await show('rejected');
  const released = previewKeys.length;
  await act(async () => {
    worker().release();
    await typed;
  });
  await until(() => shown()[0] === key('rejected'));
  expect(shown()).toEqual([key('rejected'), 'Seed!']);
  for (const [shownKey, text] of previewKeys.slice(released)) {
    if (shownKey === key('rejected')) expect(text).toBe('Seed!');
  }
});

test('settling waits for the frame of the current preview', async () => {
  const { display, worker, key, show, shown } = await previewSetup();
  worker().hold();
  await show('rejected');
  let settled: DisplayList | null = null;
  void display.current!.settledDisplayList(() => {}).then((list) => {
    settled = list;
  });
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 20));
  });
  expect(settled).toBeNull();
  await act(async () => worker().release());
  await until(() => settled !== null);
  expect(paintedText(settled)).toBe('Seed');
  expect(shown()).toEqual([key('rejected'), 'Seed']);
});

test('frames the worker builds after a preview layout was sent carry that preview', async () => {
  const { session, display, worker, previewKeys, revision, key, shown } = await previewSetup();
  worker().hold();
  let laidOut!: ReturnType<UseRustDisplayListResult['layoutInWorker']>;
  let typed!: Promise<unknown>;
  act(() => {
    laidOut = display.current!.layoutInWorker(
      session,
      layoutRequest({ revisionPreview: { [revision]: 'rejected' } })
    );
    typed = display.current!.applyInput('!');
  });
  expect(laidOut).not.toBeNull();
  const released = previewKeys.length;
  await act(async () => {
    worker().release();
    await laidOut;
    await typed;
  });
  await until(() => shown()[1] === 'Seed!');
  expect(shown()[0]).not.toBe(key('accepted'));
  for (const [shownKey, text] of previewKeys.slice(released)) {
    if (shownKey === key('accepted')) expect(text).toContain('more');
  }
});
