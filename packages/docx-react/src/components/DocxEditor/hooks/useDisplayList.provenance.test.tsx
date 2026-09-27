import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { sourceVersionOf, stampSourceVersion } from '../internals/layoutProvenance';
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

function Harness({
  session,
  layout,
  overrides,
  display,
  requestLayout,
}: {
  session: YrsSession;
  layout: Layout;
  overrides: RustDisplayListHookOverrides;
  display: { current: UseRustDisplayListResult | null };
  requestLayout: () => void;
}) {
  display.current = useRustDisplayList(
    layout,
    overrides,
    undefined,
    undefined,
    session,
    requestLayout
  );
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

async function setup() {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed');
  session.registerFont(new Uint8Array(readFileSync(FONT)));
  const layOut = () => {
    const inputs = JSON.parse(session.layoutDocumentWithRegionsJson(LAYOUT));
    stampSourceVersion(inputs.layout, session.version());
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
  const overrides: RustDisplayListHookOverrides = { getInputs: () => inputs as never };
  const harness = () => (
    <Harness
      session={session}
      layout={inputs.layout}
      overrides={overrides}
      display={display}
      requestLayout={() => layoutRequests.push(Date.now())}
    />
  );
  const view = render(harness());
  await until(() => display.current?.workerSurfacesActive === true && !!display.current.queries);
  const relayout = () => {
    inputs = layOut();
    view.rerender(harness());
  };
  return { session, paraId, display, worker: () => worker, layoutRequests, relayout };
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
