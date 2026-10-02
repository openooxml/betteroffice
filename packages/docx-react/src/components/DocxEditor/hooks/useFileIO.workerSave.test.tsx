import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, yrsToDocument, type YrsSession } from '@betteroffice/docx/yrs';
import type { Document } from '@betteroffice/docx/types/document';
import type { PagedEditorRef } from '../PagedEditor';
import { awaitWorkerOpenReplica, deferWorkerOpenReplica } from '../internals/workerOpenReplica';
import {
  awaitWorkerOpenSaves,
  peekWorkerOpenSave,
  registerWorkerOpenSave,
  takeWorkerOpenSave,
  workerOpenSaveNeedsBase,
  type WorkerOpenSave,
} from '../internals/workerOpenSave';
import { useFileIO } from './useFileIO';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');

const ROOT = resolve(import.meta.dir, '../../../../../..');
const bytes = new Uint8Array(
  readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures/page-fragments/pages.docx'))
);
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm')))
  )
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function workerOpened(save: WorkerOpenSave) {
  const worker = await createYrsSession();
  const session = await createYrsSession();
  sessions.push(worker, session);
  worker.openDocx(bytes, true);
  const hydrated = mock(() => {});
  let compatibilityBase: Document | null = null;
  const savedBase = mock(async (saved: ArrayBuffer): Promise<Document> => {
    const base = worker.materializeDocx();
    if (!base) throw new Error('the worker has no package');
    return { ...base, originalBuffer: saved };
  });
  const replica = deferWorkerOpenReplica(
    session,
    async () => {
      const update = worker.encodeState();
      const saves = awaitWorkerOpenSaves(session);
      if (saves) await saves;
      const saved = peekWorkerOpenSave(session);
      const base = saved ? await savedBase(saved) : null;
      return () => {
        hydrated();
        session.openDocx(bytes, false);
        session.loadState(update);
        compatibilityBase = base;
      };
    },
    () => {
      if (workerOpenSaveNeedsBase(session)) throw new Error('The saved resident document could not be hydrated');
      session.openDocx(bytes, true);
    },
    () => {},
    { active: () => true, request: () => replica.start() }
  );
  registerWorkerOpenSave(session, save);
  const flushes: Array<boolean | undefined> = [];
  const editor = {
    getYrsSession: () => session,
    flushPendingInput: async (awaitReplica?: boolean) => {
      flushes.push(awaitReplica);
      if (awaitReplica !== false) await awaitWorkerOpenReplica(session);
    },
    getDocument: () => {
      const base = compatibilityBase ?? session.materializeDocx();
      takeWorkerOpenSave(session);
      return base ? yrsToDocument(session, base) : null;
    },
  } satisfies Partial<PagedEditorRef>;
  const pagedEditorRef = { current: editor as unknown as PagedEditorRef | null };
  const saved: ArrayBuffer[] = [];
  const errors: Error[] = [];
  const { result } = renderHook(() =>
    useFileIO({
      pagedEditorRef,
      resolveImage: (() => null) as never,
      comments: [],
      documentName: undefined,
      onSave: (buffer) => saved.push(buffer),
      onOpen: undefined,
      onError: (error) => errors.push(error),
      onPrint: undefined,
      onDocumentNameChange: undefined,
      loadBuffer: async () => {},
      focusActiveEditor: () => {},
    })
  );
  return { session, replica, hydrated, savedBase, flushes, saved, errors, pagedEditorRef, save: result.current.handleSave };
}

test('saves in the worker without loading the pending replica', async () => {
  const bytesOut = new Uint8Array([1, 2, 3]).buffer;
  const saver = mock<WorkerOpenSave>(async () => ({ bytes: bytesOut, full: false }));
  const opened = await workerOpened(saver);
  expect(await opened.save()).toBe(bytesOut);
  expect(saver).toHaveBeenCalledTimes(1);
  expect(opened.saved).toEqual([bytesOut]);
  expect(opened.flushes).toEqual([false]);
  expect(opened.replica.started).toBe(false);
  expect(opened.hydrated).not.toHaveBeenCalled();
  expect(opened.savedBase).not.toHaveBeenCalled();
  expect(takeWorkerOpenSave(opened.session)).toBe(bytesOut);
});

test('fetches the last saved base once when the replica later hydrates', async () => {
  const first = bytes.slice().buffer;
  const second = bytes.slice().buffer;
  let saves = 0;
  const opened = await workerOpened(async () => ({
    bytes: saves++ === 0 ? first : second,
    full: false,
  }));
  expect(await opened.save()).toBe(first);
  expect(await opened.save()).toBe(second);
  expect(opened.savedBase).not.toHaveBeenCalled();
  await awaitWorkerOpenReplica(opened.session);
  expect(opened.savedBase).toHaveBeenCalledTimes(1);
  expect(opened.savedBase).toHaveBeenCalledWith(second);
  expect(opened.pagedEditorRef.current!.getDocument()?.originalBuffer).toBe(second);
  expect(await opened.save()).toBeInstanceOf(ArrayBuffer);
  expect(saves).toBe(2);
  expect(opened.savedBase).toHaveBeenCalledTimes(1);
});

test('falls back to the replica when the worker cannot save', async () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const opened = await workerOpened(async () => {
      throw new Error('Resident engine worker trapped');
    });
    const buffer = await opened.save();
    expect(buffer && new Uint8Array(buffer).subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
    expect(opened.hydrated).toHaveBeenCalledTimes(1);
    expect(opened.saved).toEqual([buffer!]);
    expect(opened.errors).toEqual([]);
    expect(warn).toHaveBeenCalledTimes(1);
  } finally {
    warn.mockRestore();
  }
});

test('falls back to the replica when no worker save is admitted', async () => {
  const opened = await workerOpened(() => null);
  const buffer = await opened.save();
  expect(buffer && new Uint8Array(buffer).subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
  expect(opened.hydrated).toHaveBeenCalledTimes(1);
});

test('a replica already loading saves on the main thread', async () => {
  const saver = mock<WorkerOpenSave>(async () => ({ bytes: new ArrayBuffer(1), full: false }));
  const opened = await workerOpened(saver);
  opened.replica.start();
  const buffer = await opened.save();
  expect(buffer && new Uint8Array(buffer).subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
  expect(saver).not.toHaveBeenCalled();
});

test('a document replaced while the worker saves fails the save', async () => {
  const other = await createYrsSession();
  sessions.push(other);
  let opened!: Awaited<ReturnType<typeof workerOpened>>;
  opened = await workerOpened(async () => {
    const editor = opened.pagedEditorRef.current!;
    opened.pagedEditorRef.current = { ...editor, getYrsSession: () => other } as PagedEditorRef;
    return { bytes: new ArrayBuffer(1), full: true };
  });
  expect(await opened.save()).toBeNull();
  expect(opened.errors.map((error) => error.message)).toEqual(['The document changed while saving']);
  expect(opened.saved).toEqual([]);
});
