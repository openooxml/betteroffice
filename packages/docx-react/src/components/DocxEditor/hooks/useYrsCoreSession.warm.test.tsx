import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { useEffect } from 'react';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsDocxHost, YrsSession } from '@betteroffice/docx/yrs';
import { useCompatibilityWarm, useYrsCoreSession } from './useYrsCoreSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const FIXTURE = resolve(
  import.meta.dir,
  '../../../../../../crates/docx-edit/tests/fixtures/paragraph-identities'
);

function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (dir: string, prefix: string) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, entry.name);
      if (entry.isDirectory()) add(path, `${prefix}${entry.name}/`);
      else parts.set(`${prefix}${entry.name}`, toBytes(readFileSync(path, 'utf8')));
    }
  };
  add(FIXTURE, '');
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function openedSession() {
  const bytes = fixture();
  let host: YrsDocxHost | null = null;
  const hook = renderHook(() =>
    useYrsCoreSession(true, null, null, bytes, 1, undefined, {
      onHostDocument: (opened) => {
        host = opened;
      },
    })
  );
  await waitFor(() => expect(hook.result.current.session).not.toBeNull());
  const session = hook.result.current.session!;
  let materializations = 0;
  const materializeDocx = session.materializeDocx.bind(session);
  session.materializeDocx = () => {
    materializations += 1;
    return materializeDocx();
  };
  const project = () => hook.result.current.documentFromYrs(host!.document);
  return { hook, project, materializations: () => materializations };
}

const idle = () => new Promise((resolve) => setTimeout(resolve, 300));

test('a session whose document fails to open is freed at once', async () => {
  const probe = createEditSession(1);
  const free = spyOn(Object.getPrototypeOf(probe) as { free(): void }, 'free');
  const errors: Error[] = [];
  const logged = spyOn(console, 'error').mockImplementation(() => {});
  try {
    renderHook(() =>
      useYrsCoreSession(true, null, null, Uint8Array.of(1, 2, 3), 1, undefined, {
        onError: (error) => errors.push(error),
      })
    );
    await waitFor(() => expect(errors).toHaveLength(1));
    expect(free).toHaveBeenCalledTimes(1);
  } finally {
    logged.mockRestore();
    free.mockRestore();
    probe.free();
  }
});

test('a replaced session stays usable until the editor renders without it', async () => {
  const errors: unknown[] = [];
  const hook = renderHook(
    ({ bytes, generation }: { bytes: Uint8Array; generation: number }) => {
      const { session } = useYrsCoreSession(true, null, null, bytes, generation);
      // A consumer that reads its session whenever the document changes, as the sidebar does.
      useEffect(() => {
        try {
          session?.version();
        } catch (error) {
          errors.push(error);
        }
      }, [session, bytes]);
      return session;
    },
    { initialProps: { bytes: fixture(), generation: 1 } }
  );
  await waitFor(() => expect(hook.result.current).not.toBeNull());
  const replaced = hook.result.current!;
  hook.rerender({ bytes: fixture(), generation: 2 });
  expect(errors).toEqual([]);
  expect(() => replaced.version()).toThrow();
  await waitFor(() => expect(hook.result.current).not.toBeNull());
  expect(hook.result.current).not.toBe(replaced);
});

test('opening never materializes the compatibility document by itself', async () => {
  const { project, materializations } = await openedSession();
  await act(idle);
  expect(materializations()).toBe(0);

  expect(project()).not.toBeNull();
  expect(materializations()).toBe(1);
  expect(project()).not.toBeNull();
  expect(materializations()).toBe(1);
});

test('a requested warm materializes once when the main thread is idle', async () => {
  const { hook, project, materializations } = await openedSession();
  hook.result.current.scheduleCompatibilityWarm();
  hook.result.current.scheduleCompatibilityWarm();
  expect(materializations()).toBe(0);
  await act(idle);
  expect(materializations()).toBe(1);
  hook.result.current.scheduleCompatibilityWarm();
  await act(idle);
  expect(project()).not.toBeNull();
  expect(materializations()).toBe(1);
});

describe('useCompatibilityWarm', () => {
  const sessions = [{ name: 'first' }, { name: 'replacement' }] as unknown as YrsSession[];
  const frames = [{ frame: 0 }, { frame: 1 }, { frame: 2 }];

  function warmer(initial: {
    session: YrsSession | null;
    frame: object | null;
    projects: boolean;
  }) {
    const calls: string[] = [];
    const schedule = () => calls.push('schedule');
    const cancel = () => calls.push('cancel');
    const hook = renderHook(
      ({ session, frame, projects }) =>
        useCompatibilityWarm(session, frame, projects, schedule, cancel),
      { initialProps: initial }
    );
    return { calls, rerender: hook.rerender };
  }

  test("waits for the session's own first frame", () => {
    const { calls, rerender } = warmer({ session: null, frame: null, projects: true });
    rerender({ session: sessions[0]!, frame: null, projects: true });
    expect(calls.filter((call) => call === 'schedule')).toEqual([]);
    rerender({ session: sessions[0]!, frame: frames[0]!, projects: true });
    expect(calls.filter((call) => call === 'schedule')).toHaveLength(1);
  });

  test('a replacement showing the previous frame waits for a frame of its own', () => {
    const { calls, rerender } = warmer({ session: sessions[0]!, frame: null, projects: true });
    rerender({ session: sessions[0]!, frame: frames[0]!, projects: true });
    calls.length = 0;
    rerender({ session: sessions[1]!, frame: frames[0]!, projects: true });
    expect(calls).toEqual(['cancel']);
    rerender({ session: sessions[1]!, frame: frames[1]!, projects: true });
    expect(calls).toEqual(['cancel', 'schedule']);
    rerender({ session: sessions[1]!, frame: frames[2]!, projects: true });
    expect(calls).toEqual(['cancel', 'schedule']);
  });

  test('losing the last content listener cancels a pending warm', () => {
    const { calls, rerender } = warmer({ session: sessions[0]!, frame: null, projects: true });
    rerender({ session: sessions[0]!, frame: frames[0]!, projects: true });
    calls.length = 0;
    rerender({ session: sessions[0]!, frame: frames[0]!, projects: false });
    expect(calls).toEqual(['cancel']);
  });

  test('never warms for a host that does not project changes', () => {
    const { calls, rerender } = warmer({ session: sessions[0]!, frame: null, projects: false });
    rerender({ session: sessions[0]!, frame: frames[0]!, projects: false });
    expect(calls.filter((call) => call === 'schedule')).toEqual([]);
    rerender({ session: sessions[0]!, frame: frames[0]!, projects: true });
    expect(calls.filter((call) => call === 'schedule')).toHaveLength(1);
  });
});
