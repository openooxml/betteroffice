import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { DisplayList } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { DocxCommandAdmissionError } from '../../../commands/createDocxCommandStore';
import type { DocxCommandResult } from '../../../commands/types';
import type { PagedEditorRef } from '../PagedEditor';
import { currentYrsToolbarSelection } from '../yrsToolbar';
import {
  useDocxCommandBinding,
  type DocxCommandInputs,
  type DocxImageInsert,
} from './useDocxCommands';
import { useFindReplaceBridge } from './useFindReplaceBridge';
import type { DocxPrintJob } from './useFileIO';
import type { PagedEditorCommandBridge } from './usePagedEditorRefApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const quiet = console.error;
const sessions: YrsSession[] = [];

beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  );
});
afterEach(() => {
  cleanup();
  console.error = quiet;
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function newSession(text = 'Hello world') {
  const session = await createYrsSession({ clientId: 900 + sessions.length });
  sessions.push(session);
  const { paraId } = session.createStory('body', text);
  session.beginUndoCapture();
  session.setSelection({ story: 'body', paraId, offset: 0 }, { story: 'body', paraId, offset: 5 });
  return { session, paraId };
}

const code = (result: DocxCommandResult) => (result.ok ? result.status : result.failure.code);

/** A binding over a real session, with a hand-driven input queue and editor bridge. */
function mount(initial: YrsSession, overrides: Partial<DocxCommandInputs> = {}) {
  let session = initial;
  let hold: Promise<void> = Promise.resolve();
  const opened: string[] = [];
  const bridge = {
    session: () => session,
    rootStory: () => 'body',
    hasPendingInput: () => false,
    hasSelection: () => session.selection() !== null,
    toolbarSelection: () => {
      const paragraphs = session.paragraphs('body');
      const map = createYrsInputPositionMap(
        'body',
        paragraphs.map((paragraph) => ({ paraId: paragraph.paraId, length: paragraph.text.length }))
      );
      return currentYrsToolbarSelection(session, map);
    },
    selectedImage: () => null,
    imagePosition: () => null,
    subscribe: () => () => {},
    format: () => true,
    command: () => true,
    history: () => false,
    select: () => true,
    runAfterPendingInput<T>(operation: () => T | Promise<T>): Promise<T> {
      const admitted = session;
      return hold.then(() => {
        if (session !== admitted) throw new DocxCommandAdmissionError('document-replaced');
        return operation();
      });
    },
  } as unknown as PagedEditorCommandBridge & Record<string, unknown>;
  const editor = {
    getYrsSession: () => session,
    syncYrsInputState: () => true,
    yrsLocToDisplayPosition: (loc: { paraId: string; offset: number }) => {
      const index = session.paragraphs('body').findIndex((p) => p.paraId === loc.paraId);
      return index < 0 ? null : index * 1000 + loc.offset + 1;
    },
    scrollToPosition: () => {},
    focus: () => {},
  } as unknown as PagedEditorRef;
  const noop = () => {};
  const inputs: DocxCommandInputs = {
    pagedEditorRef: { current: editor },
    bridgeRef: { current: bridge },
    isLoading: false,
    parseError: null,
    document: { package: {} } as never,
    get session() {
      return session;
    },
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
    getCachedStyleResolver: () => ({} as never),
    hyperlinkDialog: {
      openInsert: () => opened.push('link'),
      openEdit: () => opened.push('link'),
    } as never,
    findReplace: {
      openFind: () => opened.push('find'),
      openReplace: () => opened.push('replace'),
      setMatches: noop,
      goToMatch: noop,
    } as never,
    save: async () => 'saved',
    reservePrint: () => ({ finish: async () => true, cancel: noop }),
    renderedDisplayList: () => Promise.reject(new Error('Not rendered')),
    openDocument: noop,
    pickImage: noop,
    tableAction: () => false,
    openImageProperties: noop,
    openPageSetup: noop,
    openWatermark: noop,
    refreshTrackedChanges: noop,
    ...overrides,
  };
  const hook = renderHook(() => {
    const commands = useDocxCommandBinding(inputs);
    const find = useFindReplaceBridge({
      pagedEditorRef: inputs.pagedEditorRef,
      findReplace: inputs.findReplace,
      complete: (write) => commands.complete('replace', write),
    });
    return { commands, find };
  });
  return {
    hook,
    bridge,
    opened,
    get store() {
      return hook.result.current.commands.controller.store;
    },
    replaceDocument(next: YrsSession) {
      session = next;
    },
    /** Holds accepted input until the returned release runs. */
    holdInput(): () => void {
      let release!: () => void;
      hold = new Promise<void>((resolve) => {
        release = resolve;
      });
      return release;
    },
  };
}

describe('editor command binding', () => {
  test('engine refusals fail the command instead of reporting a no-op', async () => {
    const { session } = await newSession();
    console.error = () => {};
    const editor = mount(session);
    editor.bridge.format = () => {
      throw new Error('refused');
    };
    editor.bridge.command = () => {
      throw new Error('refused');
    };
    expect(editor.store.getState('bold').enabled).toBe(true);
    expect(code(await editor.store.execute('bold', null))).toBe('command-failed');
    expect(code(await editor.store.execute('insertPageBreak', null))).toBe('command-failed');
  });

  test('a link dialog completes only against the document and selection it opened with', async () => {
    const { session, paraId } = await newSession();
    const editor = mount(session);
    let writes = 0;
    const write = (): DocxCommandResult => {
      writes += 1;
      return { ok: true, status: 'executed' };
    };
    const complete = () => editor.hook.result.current.commands.complete('link', write);

    expect(code(await editor.store.execute('insertLink', null))).toBe('opened');
    expect(editor.opened).toEqual(['link']);
    const release = editor.holdInput();
    const pending = complete();
    session.insertText({ story: 'body', paraId, offset: 0 }, '>> ');
    release();
    expect(code(await pending)).toBe('executed');
    expect(writes).toBe(1);

    session.setSelection(
      { story: 'body', paraId, offset: 9 },
      { story: 'body', paraId, offset: 11 }
    );
    expect(code(await complete())).toBe('target-changed');

    await editor.store.execute('insertLink', null);
    const { session: replacement } = await newSession('Another document');
    editor.replaceDocument(replacement);
    expect(code(await complete())).toBe('document-replaced');
    expect(writes).toBe(1);
  });

  test('each image picker keeps the document and selection it opened with through decoding', async () => {
    const { session, paraId } = await newSession();
    const picks: DocxImageInsert[] = [];
    const editor = mount(session, { pickImage: (insert) => picks.push(insert) });
    const inserted: unknown[] = [];
    editor.bridge.command = (command: unknown) => inserted.push(command) > 0;
    const image = { src: 'data:image/png;base64,', rId: 'rId_img' };

    expect(code(await editor.store.execute('insertImage', null))).toBe('opened');
    session.setSelection({ story: 'body', paraId, offset: 9 });
    await editor.store.execute('insertImage', null);
    expect(code(await picks[0](image))).toBe('target-changed');
    expect(code(await picks[1](image))).toBe('executed');

    await editor.store.execute('insertImage', null);
    const { session: replacement } = await newSession('Another document');
    editor.replaceDocument(replacement);
    await editor.store.execute('insertImage', null);
    expect(picks).toHaveLength(4);
    expect(code(await picks[2](image))).toBe('document-replaced');
    expect(code(await picks[3](image))).toBe('executed');
    expect(inserted).toHaveLength(2);
  });

  test('replace finds its target again after preceding input and refuses a moved one', async () => {
    const { session, paraId } = await newSession('foo bar foo');
    const editor = mount(session);
    await editor.store.execute('find', null);
    const { find } = editor.hook.result.current;
    act(() => {
      find.handleFind('foo', { matchCase: false, matchWholeWord: false });
    });
    const release = editor.holdInput();
    let replaced: Promise<boolean> | undefined;
    act(() => {
      replaced = find.handleReplace('baz');
    });
    session.insertText({ story: 'body', paraId, offset: 0 }, 'X');
    release();
    expect(await replaced!).toBe(true);
    expect(session.paragraphs('body')[0].text).toBe('Xbaz bar foo');

    session.setSelection(
      { story: 'body', paraId, offset: 5 },
      { story: 'body', paraId, offset: 8 }
    );
    expect(await find.handleReplace('qux')).toBe(false);
    expect(session.paragraphs('body')[0].text).toBe('Xbaz bar foo');

    expect(
      await find.handleReplaceAll('foo', 'qux', { matchCase: false, matchWholeWord: false })
    ).toBe(1);
    expect(session.paragraphs('body')[0].text).toBe('Xbaz bar qux');
  });

  test('print reserves its window at once and prints only after input and rendering settle', async () => {
    const { session } = await newSession();
    const events: string[] = [];
    let rendered!: (displayList: DisplayList) => void;
    const displayList = { pages: [] } as unknown as DisplayList;
    const job: DocxPrintJob = {
      finish: async (list) => {
        events.push(list === displayList ? 'printed' : 'wrong list');
        return true;
      },
      cancel: () => events.push('cancelled'),
    };
    const editor = mount(session, {
      reservePrint: () => {
        events.push('reserved');
        return job;
      },
      renderedDisplayList: () => {
        events.push('rendering');
        return new Promise<DisplayList>((resolve) => {
          rendered = resolve;
        });
      },
    });
    const release = editor.holdInput();
    const printing = editor.store.execute('print', null);
    expect(events).toEqual(['reserved']);
    release();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(events).toEqual(['reserved', 'rendering']);
    rendered(displayList);
    expect(code(await printing)).toBe('executed');
    expect(events).toEqual(['reserved', 'rendering', 'printed']);
  });

  test('print reports rendering and input failures and closes its window', async () => {
    const { session } = await newSession();
    console.error = () => {};
    const events: string[] = [];
    const editor = mount(session, {
      reservePrint: () => ({ finish: async () => true, cancel: () => events.push('cancelled') }),
      renderedDisplayList: () => Promise.reject(new Error('The document did not finish rendering')),
    });
    expect(code(await editor.store.execute('print', null))).toBe('command-failed');
    expect(events).toEqual(['cancelled']);

    editor.bridge.runAfterPendingInput = () =>
      Promise.reject(new DocxCommandAdmissionError('input-failed'));
    expect(code(await editor.store.execute('print', null))).toBe('input-failed');
    expect(events).toEqual(['cancelled', 'cancelled']);
  });

  test('accepts and rejects header and footer revisions by id', async () => {
    const { session } = await newSession();
    const header = session.createStory('hf:rId7', 'Header');
    const footer = session.createStory('hf:rId8', 'Footer');
    const author = { name: 'Reviewer', date: '2026-01-01T00:00:00Z' };
    session.insertText({ story: 'hf:rId7', paraId: header.paraId, offset: 6 }, ' new', author);
    session.insertText({ story: 'hf:rId8', paraId: footer.paraId, offset: 6 }, ' old', author);
    const [inHeader, inFooter] = session.listRevisions();
    const editor = mount(session);
    expect(editor.store.getState('reviewNext').enabled).toBe(false);
    expect(editor.store.getState('reviewAccept', { revisionId: inHeader.revisionId }).enabled).toBe(
      true
    );

    expect(
      code(await editor.store.execute('reviewAccept', { revisionId: inHeader.revisionId }))
    ).toBe('executed');
    expect(
      code(await editor.store.execute('reviewReject', { revisionId: inFooter.revisionId }))
    ).toBe('executed');
    expect(session.listRevisions()).toEqual([]);
    expect(session.paragraphs('hf:rId7')[0].text).toBe('Header new');
    expect(session.paragraphs('hf:rId8')[0].text).toBe('Footer');
    const gone = await editor.store.execute('reviewAccept', { revisionId: inHeader.revisionId });
    expect(code(gone)).toBe('revision-not-found');
  });
});
