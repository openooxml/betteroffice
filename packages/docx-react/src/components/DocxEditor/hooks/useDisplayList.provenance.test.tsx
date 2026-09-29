import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
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
  previewKeys: (string | null)[];
}) {
  display.current = useRustDisplayList(
    layout,
    overrides,
    undefined,
    undefined,
    session,
    requestLayout
  );
  previewKeys.push(revisionPreviewKeyOf(display.current.queries));
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
  const previewKeys: (string | null)[] = [];
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
function paintedText(display: { current: UseRustDisplayListResult | null }): string {
  const segments: [string, string][] = [];
  for (const primitive of (display.current?.displayList?.pages ?? []).flatMap(
    (page) => page.primitives
  )) {
    if (primitive.kind !== 'text' && primitive.kind !== 'glyphRun') continue;
    const kind = primitive.revision?.kind ?? '';
    const last = segments.at(-1);
    if (last && last[1] === kind) last[0] += primitive.text;
    else segments.push([primitive.text, kind]);
  }
  return segments.map(([text, kind]) => (kind ? `${text}[${kind}]` : text)).join('');
}

test('a frame carries the revision preview of its layout, and a superseded preview never lands', async () => {
  const renderEnv: { current: YrsRenderEnv } = { current: {} };
  const { session, paraId, display, worker, previewKeys, relayout } = await setup(renderEnv);
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
  const show = async (revisionPreview?: YrsRenderEnv['revisionPreview']) => {
    renderEnv.current = revisionPreview ? { revisionPreview } : {};
    await act(async () => relayout());
  };

  await show({ [revision]: 'accepted' });
  const acceptedKey = revisionPreviewKey({ [revision]: 'accepted' });
  await until(() => revisionPreviewKeyOf(display.current!.queries) === acceptedKey);
  expect(paintedText(display)).toBe('Seed more');
  const acceptedQueries = display.current!.queries;
  await act(async () => {
    await display.current!.applyInput('!');
  });
  expect(revisionPreviewKeyOf(display.current!.queries)).toBe(acceptedKey);

  worker().hold();
  await show({ [revision]: 'rejected' });
  await show();
  const released = previewKeys.length;
  await act(async () => worker().release());
  await until(
    () =>
      revisionPreviewKeyOf(display.current!.queries) === '' &&
      sourceVersionOf(display.current!.queries) === session.version()
  );
  expect(paintedText(display)).toBe('Seed more[ins]!');
  expect(previewKeys.slice(released)).not.toContain(revisionPreviewKey({ [revision]: 'rejected' }));
  expect(revisionPreviewKeyOf(acceptedQueries)).toBe(acceptedKey);
  expect(revisionPreviewKey({ b: 'rejected', a: 'accepted' })).toBe(
    revisionPreviewKey({ a: 'accepted', b: 'rejected' })
  );
  expect(revisionPreviewKey({})).toBe('');
});
