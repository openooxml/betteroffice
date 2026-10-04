import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { DisplayList } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { DocxCommandAdmissionError } from '../../../commands/createDocxCommandStore';
import type { DocxCommandResult } from '../../../commands/types';
import type { PagedEditorRef } from '../PagedEditor';
import { createYrsPositionProjection } from '../internals/yrsPositionProjection';
import { awaitWorkerOpenReplica, deferWorkerOpenReplica, workerOpenReplicaPending } from '../internals/workerOpenReplica';
import { currentYrsToolbarSelection } from '../yrsToolbar';
import {
  useDocxCommandBinding,
  type DocxCommandInputs,
  type DocxImageInsert,
} from './useDocxCommands';
import { useFindReplaceBridge } from './useFindReplaceBridge';
import type { DocxPrintJob } from './useFileIO';
import { usePagedEditorCommandBridge, type PagedEditorCommandBridge } from './usePagedEditorRefApi';

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
  mock.restore();
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
    isWorkerViewer: () => false,
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
      state: { isOpen: true },
    } as never,
    save: async () => 'saved',
    reservePrint: () => ({ prepare: async () => {}, print: () => true, cancel: noop }),
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
  for (const readOnly of [true, false]) {
    test(`worker viewer mutations refuse ${readOnly ? 'read-only' : 'viewing-mode'} without input admission`, async () => {
      const { session } = await newSession();
      const hydrate = mock(async () => () => {});
      const request = mock(() => {});
      const fallback = mock(() => {});
      const replica = deferWorkerOpenReplica(session, hydrate, fallback, () => {}, { active: () => true, request });
      const editor = mount(session, {
        readOnly, mode: 'viewing', experimentalWorkerOpen: true,
        pagedEditorRef: { current: { isWorkerViewer: () => true } as PagedEditorRef },
      });
      const admit = spyOn(editor.bridge, 'runAfterPendingInput').mockImplementation(async (operation) => {
        await awaitWorkerOpenReplica(session);
        return operation();
      });
      const accept = spyOn(session, 'acceptChange');
      const reject = spyOn(session, 'rejectChange');
      const listing = spyOn(session, 'listRevisions');
      for (const id of ['bold', 'insertPageBreak', 'reviewAccept', 'reviewReject'] as const) {
        const result = editor.store.execute(id, id === 'reviewAccept' || id === 'reviewReject' ? { revisionId: 'revision' } : null);
        expect(admit).not.toHaveBeenCalled();
        expect(code(await result)).toBe(readOnly ? 'read-only' : 'viewing-mode');
      }
      expect(code(await editor.store.execute('reviewReject', null))).toBe(readOnly ? 'read-only' : 'viewing-mode');
      expect(accept).not.toHaveBeenCalled();
      expect(reject).not.toHaveBeenCalled();
      expect(listing).not.toHaveBeenCalled();
      expect(admit).not.toHaveBeenCalled();
      expect(replica.started).toBe(false);
      expect(request).not.toHaveBeenCalled();
      expect(hydrate).not.toHaveBeenCalled();
      expect(fallback).not.toHaveBeenCalled();
    });
  }

  test('viewer revision navigation refuses without reading or admitting the edit peer', async () => {
    const { session } = await newSession();
    const request = mock(() => {});
    const hydrate = mock(async () => () => {});
    const replica = deferWorkerOpenReplica(session, hydrate, () => {}, () => {}, { active: () => true, request });
    const editor = mount(session, {
      mode: 'viewing', experimentalWorkerOpen: true, viewerSession: true,
      pagedEditorRef: { current: { isWorkerViewer: () => false } as PagedEditorRef },
    });
    const admit = spyOn(editor.bridge, 'runAfterPendingInput').mockImplementation(() => new Promise<never>(() => {}));
    const listing = spyOn(session, 'listRevisions');
    for (const id of ['reviewNext', 'reviewPrevious'] as const) {
      expect(editor.store.getState(id)).toMatchObject({ enabled: false, disabledReason: { code: 'no-revisions' } });
      const result = editor.store.execute(id, null);
      expect(admit).not.toHaveBeenCalled();
      expect(code(await result)).toBe('no-revisions');
    }
    expect(listing).not.toHaveBeenCalled();
    expect(replica.started).toBe(false);
    expect(request).not.toHaveBeenCalled();
    expect(hydrate).not.toHaveBeenCalled();
  });

  test('viewer session revision decisions refuse after worker read routing ends', async () => {
    const { session } = await newSession();
    const editor = mount(session, {
      mode: 'viewing', viewerSession: true,
      pagedEditorRef: { current: { isWorkerViewer: () => false } as PagedEditorRef },
    });
    const admit = spyOn(editor.bridge, 'runAfterPendingInput').mockImplementation(() => new Promise<never>(() => {}));
    const listing = spyOn(session, 'listRevisions');
    for (const id of ['reviewAccept', 'reviewReject'] as const) {
      const result = editor.store.execute(id, { revisionId: 'revision' });
      expect(admit).not.toHaveBeenCalled();
      expect(code(await result)).toBe('viewing-mode');
    }
    expect(listing).not.toHaveBeenCalled();
  });

  test('editor revision navigation remains ordered and selects document revisions', async () => {
    const { session, paraId } = await newSession();
    session.insertText({ story: 'body', paraId, offset: 5 }, ' new', { name: 'Reviewer', date: '2026-01-01T00:00:00Z' });
    const [revision] = session.listRevisions();
    const editor = mount(session, {
      pagedEditorRef: { current: { isWorkerViewer: () => false, getSelectionRange: () => null } as unknown as PagedEditorRef },
    });
    const admit = spyOn(editor.bridge, 'runAfterPendingInput');
    const select = spyOn(editor.bridge, 'select');
    for (const id of ['reviewNext', 'reviewPrevious'] as const) {
      expect(code(await editor.store.execute(id, null))).toBe('executed');
      expect(select).toHaveBeenLastCalledWith(
        { story: revision.range.story, ...revision.range.start },
        { story: revision.range.story, ...revision.range.end }
      );
    }
    expect(admit).toHaveBeenCalledTimes(2);
  });

  test('editor revision decisions remain ordered after pending input', async () => {
    const { session } = await newSession();
    const editor = mount(session);
    const admit = spyOn(editor.bridge, 'runAfterPendingInput');
    const release = editor.holdInput();
    let settled = false;
    const accept = editor.store.execute('reviewAccept', { revisionId: 'revision' });
    const reject = editor.store.execute('reviewReject', null).then((result) => { settled = true; return result; });
    expect(admit).toHaveBeenCalledTimes(2);
    await Promise.resolve();
    expect(settled).toBe(false);
    release();
    expect(code(await accept)).toBe('revision-not-found');
    expect(code(await reject)).toBe('revision-required');
  });

  async function workerViewer(selectedText: Promise<string>) {
    const session = await createYrsSession({ clientId: 950 + sessions.length });
    sessions.push(session);
    const hydrate = mock(async () => () => {});
    const request = mock(() => {});
    const replica = deferWorkerOpenReplica(session, hydrate, () => {}, () => {}, {
      active: () => true, request,
    });
    const readSelectedText = mock(() => selectedText);
    const openFind = mock(() => {});
    const openReplace = mock(() => {});
    const editor = mount(session, {
      experimentalWorkerOpen: true,
      mode: 'viewing',
      pagedEditorRef: { current: {
        isWorkerViewer: () => true,
        getYrsSession: () => session,
        readSelectedText,
      } as unknown as PagedEditorRef },
      findReplace: { openFind, openReplace, state: { isOpen: true } } as never,
    });
    editor.bridge.toolbarSelection = () => null;
    editor.bridge.hasSelection = () => false;
    const admit = mock(async () => { throw new Error('unexpected replica admission'); });
    editor.bridge.runAfterPendingInput = admit;
    return { ...editor, session, replica, hydrate, request, readSelectedText, openFind, openReplace, admit };
  }

  test('worker viewer find opens with worker selected text while the replica stays pending', async () => {
    const editor = await workerViewer(Promise.resolve('Selected text'));
    try {
      const pending = editor.store.execute('find', null);
      expect(pending).toBeInstanceOf(Promise);
      expect(code(await pending)).toBe('opened');
      expect(editor.openFind).toHaveBeenCalledWith('Selected text');
      expect(editor.readSelectedText).toHaveBeenCalledTimes(1);
      expect(editor.admit).not.toHaveBeenCalled();
      expect(editor.request).not.toHaveBeenCalled();
      expect(editor.hydrate).not.toHaveBeenCalled();
      expect(workerOpenReplicaPending(editor.session)).toBe(true);
    } finally {
      editor.replica.cancel();
    }
  });

  test('worker viewer replace is refused without requesting the replica', async () => {
    const editor = await workerViewer(Promise.resolve('Selected text'));
    try {
      expect(code(await editor.store.execute('replace', null))).toBe('viewing-mode');
      expect(editor.openReplace).not.toHaveBeenCalled();
      expect(editor.readSelectedText).not.toHaveBeenCalled();
      expect(editor.admit).not.toHaveBeenCalled();
      expect(editor.request).not.toHaveBeenCalled();
      expect(editor.hydrate).not.toHaveBeenCalled();
      expect(workerOpenReplicaPending(editor.session)).toBe(true);
    } finally {
      editor.replica.cancel();
    }
  });

  test('worker viewer find bounds its selected-text read to 500 ms', async () => {
    const editor = await workerViewer(new Promise(() => {}));
    try {
      expect(code(await editor.store.execute('find', null))).toBe('opened');
      expect(editor.openFind).toHaveBeenCalledWith('');
      expect(editor.admit).not.toHaveBeenCalled();
      expect(editor.request).not.toHaveBeenCalled();
      expect(editor.hydrate).not.toHaveBeenCalled();
    } finally {
      editor.replica.cancel();
    }
  });

  test('worker viewer find opens empty when its selected-text read rejects', async () => {
    const selectedText = Promise.reject(new Error('read failed'));
    void selectedText.catch(() => {});
    const editor = await workerViewer(selectedText);
    try {
      expect(code(await editor.store.execute('find', null))).toBe('opened');
      expect(editor.openFind).toHaveBeenCalledWith('');
      expect(editor.request).not.toHaveBeenCalled();
      expect(editor.hydrate).not.toHaveBeenCalled();
    } finally {
      editor.replica.cancel();
    }
  });

  test('editor find and replace prefill from the synchronous Yrs selection', async () => {
    const { session } = await newSession();
    const openFind = mock(() => {});
    const openReplace = mock(() => {});
    const editor = mount(session, { findReplace: { openFind, openReplace, state: { isOpen: true } } as never });
    expect(code(await editor.store.execute('find', null))).toBe('opened');
    expect(openFind).toHaveBeenCalledWith('Hello');
    expect(code(await editor.store.execute('replace', null))).toBe('opened');
    expect(openReplace).toHaveBeenCalledWith('Hello');
  });

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

  for (const onDemand of [true, false]) {
    test(`viewer print skips peer admission with ${onDemand ? 'an on-demand' : 'a regular'} pending replica`, async () => {
      const { session } = await newSession();
      const hydrate = mock(async () => () => {});
      const request = mock(() => {});
      const replica = deferWorkerOpenReplica(session, hydrate, () => {}, () => {}, onDemand ? { active: () => true, request } : undefined);
      const displayList = { pages: [] } as unknown as DisplayList;
      const prepare = mock(async (_list: DisplayList) => {});
      const print = mock(() => true);
      const cancel = mock(() => {});
      const editor = mount(session, {
        mode: 'viewing', viewerSession: true, experimentalWorkerOpen: true,
        reservePrint: () => ({ prepare, print, cancel }),
        renderedDisplayList: async () => displayList,
      });
      const admit = spyOn(editor.bridge, 'runAfterPendingInput').mockImplementation(() => new Promise<never>(() => {}));
      const printing = editor.store.execute('print', null);
      expect(admit).not.toHaveBeenCalled();
      expect(code(await printing)).toBe('executed');
      expect(prepare).toHaveBeenCalledWith(displayList);
      expect(print).toHaveBeenCalledTimes(1);
      expect(cancel).not.toHaveBeenCalled();
      expect(replica.started).toBe(false);
      expect(request).not.toHaveBeenCalled();
      expect(hydrate).not.toHaveBeenCalled();
    });
  }

  test('print reserves its window at once and prints only after input and rendering settle', async () => {
    const { session } = await newSession();
    const events: string[] = [];
    let rendered!: (displayList: DisplayList) => void;
    const displayList = { pages: [] } as unknown as DisplayList;
    const job: DocxPrintJob = {
      prepare: async (list) => {
        events.push(list === displayList ? 'prepared' : 'wrong list');
      },
      print: () => {
        events.push('printed');
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
    expect(events).toEqual(['reserved', 'rendering', 'prepared', 'printed']);
  });

  test('print closes its window instead of printing a document replaced while its pages render', async () => {
    const { session } = await newSession();
    const { session: replacement } = await newSession('Another document');
    const events: string[] = [];
    let prepared!: () => void;
    const editor = mount(session, {
      reservePrint: () => ({
        prepare: () => {
          events.push('preparing');
          return new Promise<void>((resolve) => {
            prepared = resolve;
          });
        },
        print: () => {
          events.push('printed');
          return true;
        },
        cancel: () => events.push('cancelled'),
      }),
      renderedDisplayList: async () => ({ pages: [] }) as unknown as DisplayList,
    });
    const printing = editor.store.execute('print', null);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(events).toEqual(['preparing']);
    editor.replaceDocument(replacement);
    prepared();
    expect(code(await printing)).toBe('document-replaced');
    expect(events).toEqual(['preparing', 'cancelled']);
  });

  test('print reports rendering and input failures and closes its window', async () => {
    const { session } = await newSession();
    console.error = () => {};
    const events: string[] = [];
    const editor = mount(session, {
      reservePrint: () => ({
        prepare: async () => {},
        print: () => true,
        cancel: () => events.push('cancelled'),
      }),
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

describe('editor command bridge', () => {
  test('an image handle follows its own image among images sharing a relationship id', async () => {
    const { session, paraId } = await newSession('ABCDE');
    const image = { src: 'data:image/png;base64,', rId: 'rIdShared' };
    session.insertImage({ story: 'body', paraId, offset: 2 }, image);
    session.insertImage({ story: 'body', paraId, offset: 1 }, image);
    const projection = () => createYrsPositionProjection(session, 'body');
    const toLoc = (position: number) => {
      const target = projection()!.targetAt(position);
      const map = createYrsInputPositionMap(target.story, session.paragraphSpans(target.story));
      return displayPositionToYrsLoc(map, target.displayPosition);
    };
    const bridgeRef: { current: PagedEditorCommandBridge | null } = { current: null };
    renderHook(() =>
      usePagedEditorCommandBridge({
        bridgeRef,
        yrsInputRef: { current: null },
        session,
        rootStory: 'body',
        inputPositionMap: () => null,
        latestSelectionRef: { current: null },
        listenersRef: { current: new Set() },
        getPositionProjection: projection,
        displayPositionToLoc: toLoc,
        format: () => false,
        command: () => false,
        syncYrsInputState: () => true,
        yrsLocToDisplayPosition: (loc) => projection()?.positionForLoc(loc) ?? null,
        scrollToPositionImpl: () => {},
      })
    );
    const bridge = bridgeRef.current!;
    const second = bridge.imageHandle(4)!;
    session.insertText({ story: 'body', paraId, offset: 0 }, '>> ');
    const pos = bridge.imagePosition(second)!;
    expect(pos).toBe(7);

    session.setImageGeometryAt(toLoc(pos)!, {
      widthEmu: 1,
      heightEmu: 1,
      other: { alt: 'second' },
    });
    const alts = session
      .storySegments('body')
      .flatMap((segment) => (segment.kind === 'embed' ? [segment.payload.alt ?? null] : []));
    expect(alts).toEqual([null, 'second']);

    const removed = { story: 'body', start: { paraId, offset: 6 }, end: { paraId, offset: 7 } };
    session.deleteRange(removed);
    expect(bridge.imagePosition(second)).toBeNull();
  });
});

test('viewer sidebar commands stay available without document reads', async () => {
  const { session } = await newSession();
  const reads = [
    spyOn(session, 'selection'), spyOn(session, 'paragraphs'),
    spyOn(session, 'listRevisions'), spyOn(session, 'version'),
    spyOn(session, 'canUndo'), spyOn(session, 'canRedo'),
  ];
  let opened = false;
  try {
    const editor = mount(session, {
      viewerSession: true, readOnly: true, mode: 'viewing',
      setShowCommentsSidebar: (next) => { opened = typeof next === 'function' ? next(opened) : next; },
    });
    expect(editor.store.getState('commentsSidebar').enabled).toBe(true);
    await act(async () => {
      expect(code(await editor.store.execute('commentsSidebar', null))).toBe('executed');
    });
    expect(opened).toBe(true);
    for (const read of reads) expect(read).not.toHaveBeenCalled();
  } finally {
    for (const read of reads) read.mockRestore();
  }
});
