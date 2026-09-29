import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput } from './YrsInput';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function session(text: string): Promise<YrsSession> {
  const created = await createYrsSession();
  sessions.push(created);
  created.createStory('body', text);
  return created;
}

function input(current: YrsSession, onStateChange: () => void, story = 'body') {
  const map = () =>
    createYrsInputPositionMap(
      'body',
      current.paragraphs('body').map((p) => ({ paraId: p.paraId, length: p.text.length }))
    );
  return (
    <YrsInput
      enabled
      readOnly={false}
      session={current}
      story={story}
      inputPositionMap={map}
      displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(), loc)}
      onStateChange={onStateChange}
      onDirectInput={() => {}}
    />
  );
}

const idle = () => act(() => new Promise((resolve) => setTimeout(resolve, 600)));

test("a session's first selection event waits for idle, its selection and later events do not", async () => {
  const first = await session('First');
  let events = 0;
  const count = () => {
    events += 1;
  };
  const view = render(input(first, count));
  expect(events).toBe(0);
  expect(first.selection()?.anchor.offset).toBe(0);

  // Re-renders before it fires leave the one pending event.
  view.rerender(input(first, () => count()));
  expect(events).toBe(0);
  await idle();
  expect(events).toBe(1);

  view.rerender(input(first, () => count()));
  expect(events).toBe(2);

  const second = await session('Second');
  view.rerender(input(second, count));
  expect(events).toBe(2);
  await idle();
  expect(events).toBe(3);
});

test('an unmounted input sends no first selection event', async () => {
  const current = await session('Gone');
  let events = 0;
  const view = render(
    input(current, () => {
      events += 1;
    })
  );
  view.unmount();
  await idle();
  expect(events).toBe(0);
});
