import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '@betteroffice/docx/docx';
import type { Document } from '@betteroffice/docx/types/document';
import { isMacPlatform } from '../../../commands/descriptors';
import type { PagedEditorRef } from '../PagedEditor';
import { useFileIO } from './useFileIO';
import { useKeyboardShortcuts } from './useKeyboardShortcuts';
import { useDocxCommandBinding, type DocxCommandInputs } from './useDocxCommands';
import type { PagedEditorCommandBridge } from './usePagedEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
let fixture: Document;
beforeAll(async () => {
  fixture = await parseDocx(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx'))),
    { preloadFonts: false }
  );
});
const MOD = isMacPlatform() ? { metaKey: true } : { ctrlKey: true };

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function commandInputs(
  document: Document,
  session: object,
  pagedEditorRef: { current: PagedEditorRef },
  save: DocxCommandInputs['save']
): DocxCommandInputs {
  const bridge = {
    session: () => session,
    rootStory: () => 'body',
    hasPendingInput: () => false,
    hasSelection: () => false,
    toolbarSelection: () => null,
    selectedImage: () => null,
    subscribe: () => () => {},
    runAfterPendingInput: () => {
      throw new Error('Save must not wait in the input queue');
    },
  } as unknown as PagedEditorCommandBridge;
  const noop = () => {};
  return {
    pagedEditorRef,
    bridgeRef: { current: bridge },
    isLoading: false,
    parseError: null,
    document,
    session: session as never,
    readOnly: false,
    mode: 'editing',
    modeControlled: false,
    onModeChange: undefined,
    setEditingMode: noop,
    sidebarOpen: false,
    sidebarControlled: false,
    sidebarHasSetter: false,
    setShowCommentsSidebar: noop,
    setExpandedSidebarItem: noop,
    zoom: 1,
    setZoom: noop,
    showFileOpen: false,
    showHelpMenu: false,
    partEditing: false,
    fontFamilies: undefined,
    documentFonts: [],
    theme: null,
    i18n: undefined,
    isDark: false,
    displayListQueries: null,
    getCachedStyleResolver: () => ({}) as never,
    hyperlinkDialog: {} as never,
    findReplace: {} as never,
    save,
    reservePrint: () => ({ prepare: async () => {}, print: () => true, cancel: noop }),
    renderedDisplayList: () => Promise.reject(new Error('Not rendered')),
    openDocument: noop,
    pickImage: noop,
    tableAction: () => false,
    openImageProperties: noop,
    openPageSetup: noop,
    openWatermark: noop,
    refreshTrackedChanges: noop,
  };
}

function setup(
  options: {
    onSaveRequest?: () => boolean | void | Promise<boolean | void>;
    flush?: () => Promise<void>;
    focused?: boolean;
  } = {}
) {
  const events: string[] = [];
  const errors: Error[] = [];
  const saved: ArrayBuffer[] = [];
  const document = structuredClone(fixture);
  const session = {
    canUndo: () => false,
    canRedo: () => false,
    onUpdate: () => () => {},
    selection: () => null,
  };
  const editor = {
    getYrsSession: () => session,
    isFocused: () => options.focused ?? true,
    focus: () => {},
    flushPendingInput: async () => {
      events.push('flush');
      await options.flush?.();
    },
    getDocument: () => {
      events.push('snapshot');
      return document;
    },
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor };
  const hook = renderHook(() => {
    const io = useFileIO({
      pagedEditorRef,
      resolveImage: () => null,
      comments: document.package.document.comments ?? [],
      documentName: 'saved',
      onSave: (buffer) => {
        events.push('saved');
        saved.push(buffer);
      },
      onSaveRequest: options.onSaveRequest,
      downloadOnSave: false,
      onError: (error) => errors.push(error),
      onOpen: undefined,
      onPrint: undefined,
      onDocumentNameChange: undefined,
      loadBuffer: async () => {},
      focusActiveEditor: () => {},
    });
    const { controller } = useDocxCommandBinding(
      commandInputs(document, session, pagedEditorRef, io.handleDownloadDocument)
    );
    useKeyboardShortcuts({
      commands: controller,
      pagedEditorRef,
      disableFindReplaceShortcuts: true,
      tableSelection: { state: { tableIndex: null } } as never,
    });
    return { io, controller };
  });
  return { hook, events, errors, saved, editor, pagedEditorRef };
}

test('a host owns Save before serialization and can save explicitly without reentry', async () => {
  let requests = 0;
  const state = setup({
    onSaveRequest: async () => {
      requests += 1;
      expect(state.events).toEqual([]);
      await state.hook.result.current.io.handleSave();
    },
  });
  await act(async () => {
    await state.hook.result.current.io.handleDownloadDocument();
  });
  expect(requests).toBe(1);
  expect(state.events).toEqual(['flush', 'snapshot', 'saved']);
  expect(state.saved).toHaveLength(1);
  expect(state.saved[0].byteLength).toBeGreaterThan(0);
  expect(state.errors).toEqual([]);
});

test('an awaited request can continue the built-in export exactly once', async () => {
  let release!: () => void;
  let requests = 0;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const state = setup({
    onSaveRequest: async () => {
      requests += 1;
      await gate;
      return true;
    },
  });
  const first = state.hook.result.current.io.handleDownloadDocument();
  const second = state.hook.result.current.io.handleDownloadDocument();
  expect(first).toBe(second);
  await Promise.resolve();
  expect(state.events).toEqual([]);
  await act(async () => {
    release();
    await first;
  });
  expect(requests).toBe(1);
  expect(state.events).toEqual(['flush', 'snapshot', 'saved']);
  expect(state.errors).toEqual([]);
});

test('cancellation and callback failures prevent export', async () => {
  const cancelled = setup({ onSaveRequest: () => false });
  await cancelled.hook.result.current.io.handleDownloadDocument();
  expect(cancelled.events).toEqual([]);
  const failure = new Error('revision mismatch');
  const failed = setup({
    onSaveRequest: () => {
      throw failure;
    },
  });
  await failed.hook.result.current.io.handleDownloadDocument();
  expect(failed.events).toEqual([]);
  expect(failed.errors).toEqual([failure]);
});

test('built-in export waits for input and aborts if the document changes', async () => {
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const state = setup({ flush: () => gate });
  const saved = state.hook.result.current.io.handleSave();
  expect(state.events).toEqual(['flush']);
  state.pagedEditorRef.current = { ...state.editor };
  release();
  expect(await saved).toBeNull();
  expect(state.events).toEqual(['flush']);
  expect(state.errors[0].message).toContain('document changed');
});

test('failed input flush prevents serialization', async () => {
  const state = setup({
    flush: async () => {
      throw new Error('input failed');
    },
  });
  await state.hook.result.current.io.handleDownloadDocument();
  expect(state.events).toEqual(['flush']);
  expect(state.errors[0].message).toBe('input failed');
});

test('the save shortcut invokes only the focused editor and ignores repeat events', async () => {
  let inactiveRequests = 0;
  let activeRequests = 0;
  setup({
    focused: false,
    onSaveRequest: () => {
      inactiveRequests += 1;
    },
  });
  setup({
    onSaveRequest: () => {
      activeRequests += 1;
    },
  });
  const event = new KeyboardEvent('keydown', {
    key: 's',
    ...MOD,
    bubbles: true,
    cancelable: true,
  });
  await act(async () => {
    document.dispatchEvent(event);
  });
  expect(event.defaultPrevented).toBe(true);
  expect(activeRequests).toBe(1);
  expect(inactiveRequests).toBe(0);
  await act(async () => {
    document.dispatchEvent(
      new KeyboardEvent('keydown', {
        key: 's',
        ...MOD,
        repeat: true,
        bubbles: true,
        cancelable: true,
      })
    );
  });
  expect(activeRequests).toBe(1);
});

test('the save command keeps interception, sharing and failures', async () => {
  const cancelled = setup({ onSaveRequest: () => false });
  expect(await cancelled.hook.result.current.controller.store.execute('save', null)).toEqual({
    ok: true,
    status: 'requested',
  });
  expect(cancelled.events).toEqual([]);

  let requests = 0;
  const continued = setup({
    onSaveRequest: async () => {
      requests += 1;
      return true;
    },
  });
  const store = continued.hook.result.current.controller.store;
  let results: unknown[] = [];
  await act(async () => {
    results = await Promise.all([store.execute('save', null), store.execute('save', null)]);
  });
  expect(results).toEqual([
    { ok: true, status: 'executed' },
    { ok: true, status: 'executed' },
  ]);
  expect(requests).toBe(1);
  expect(continued.events).toEqual(['flush', 'snapshot', 'saved']);

  const failing = setup({
    flush: async () => {
      throw new Error('input failed');
    },
  });
  const failure = await failing.hook.result.current.controller.store.execute('save', null);
  expect(failure.ok ? null : failure.failure.code).toBe('command-failed');
  expect(failing.errors[0].message).toBe('input failed');
});
