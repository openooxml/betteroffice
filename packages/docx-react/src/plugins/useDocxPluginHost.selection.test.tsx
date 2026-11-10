import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { createYrsInputPositionMap, createYrsSession, withYrsSelectionCause, type YrsSession } from '@betteroffice/docx/yrs';
import { selectionClock } from '../../../../e2e/browser/plugin-host/selection-clock';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import { currentYrsToolbarSelection } from '../components/DocxEditor/yrsToolbar';
import { yrsSelectionState, useSelectionTracker } from '../components/DocxEditor/hooks/useSelectionTracker';
import { stampRevisionPreviewKey, stampSourceVersion } from '../components/DocxEditor/internals/layoutProvenance';
import type { ViewerSelectionChange } from '../components/DocxEditor/internals/viewerSelectionController';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { SelectionState } from '../components/DocxEditor/types';
import { defineDocxPlugin } from './defineDocxPlugin';
import type { DocxPluginContext, DocxPluginError, DocxPluginSelection } from './types';
import { useDocxPluginHost } from './useDocxPluginHost';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const sessions: YrsSession[] = [];
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function setup(viewerSelection: boolean, initialRange: { from: number; to: number } | null = { from: 5, to: 9 }, story = 'hf:stale', prepare?: (session: YrsSession) => void) {
  let shellReads = 0;
  let range = initialRange;
  const session = await createYrsSession();
  sessions.push(session);
  const [paraId] = session.loadStories([{ storyId: story, paragraphs: [{ text: 'Uniform text for selection' }] }])[story]!;
  const selectRange = (next: { from: number; to: number }) => session.setSelection(
    { story, paraId: paraId!, offset: next.from }, { story, paraId: paraId!, offset: next.to }
  );
  selectRange(initialRange ?? { from: 0, to: 0 });
  prepare?.(session);
  const readSelection = session.selection;
  session.selection = () => { shellReads += 1; return readSelection(); };
  const queries = {
    displayList: { pages: [] },
    sourceState: () => ({ status: 'ready' }),
    pageCount: () => 0,
  } as unknown as DisplayListQueries;
  stampSourceVersion(queries, session.version());
  stampRevisionPreviewKey(queries, '');
  const editor = {
    getYrsSession: () => session,
    getSelectionRange: () => range,
    getLayout: () => null,
  } as unknown as PagedEditorRef;
  const events: DocxPluginSelection[] = [];
  const settledSelections: DocxPluginSelection[] = [];
  const propSelections: Array<SelectionState | null> = [];
  const subscribedSelections: Array<SelectionState | null> = [];
  const subscribers = { current: new Set<(state: SelectionState | null) => void>([
    (selection) => subscribedSelections.push(selection),
  ]) };
  const plugin = defineDocxPlugin({
    id: 'test.selection',
    createState: () => null,
    onEvent(_context, event) {
      if (event.type === 'selection-change') events.push(event.selection);
      if (event.type === 'selection-settle') settledSelections.push(event.selection);
    },
  });
  const errors: DocxPluginError[] = [];
  const options = {
    onPluginError: (error: DocxPluginError) => errors.push(error),
    plugins: [plugin], pagedEditorRef: { current: editor },
    writeModeRef: { current: 'viewing' as const }, mode: 'viewing' as const,
    readOnly: true, viewerSelection, commands: createDocxCommandController(),
    session, loadGeneration: 0, queries, layoutError: null, zoom: 1,
    canvasHostRef: { current: null }, overlayTarget: null,
    selectionChangeSubscribersRef: subscribers,
    i18n: undefined, onRenderedDomContextReady: undefined,
  };
  const borderSpecRef = { current: { style: 'single', size: 4, color: { rgb: '000000' } } };
  const { result, rerender, unmount } = renderHook(() => {
    const host = useDocxPluginHost(options);
    const tracker = useSelectionTracker({
      borderSpecRef, theme: null, setFloatingCommentBtn: () => {},
      applySelectionDelta: () => {}, recomputeFloatingCommentBtn: () => {},
      onSelectionChange: (selection) => propSelections.push(selection),
      selectionChangeSubscribersRef: subscribers,
    });
    return { host, tracker };
  });
  await waitFor(() => expect(result.current.host.activations[0]?.context.snapshot.layout).toBeTruthy());
  const formatting = () => {
    const currentStory = session.selection()!.head.story;
    return yrsSelectionState(currentYrsToolbarSelection(session, createYrsInputPositionMap(currentStory,
      session.paragraphs(currentStory).map(paragraph => ({ paraId: paragraph.paraId, length: paragraph.text.length }))))!);
  };
  return { result, rerender, unmount, options, errors, subscribers, formatting, events, settledSelections, propSelections, subscribedSelections, session,
    selectRange,
    setDisplayRange: (next: typeof range) => { range = next; },
    setRange: (next: typeof range) => { range = next; if (next) selectRange(next); },
    shellReads: () => shellReads };
}

test('a paragraph removed at flush time clears selection without blaming the plugin', async () => {
  const view = await setup(false, { from: 1, to: 3 }, 'body');
  const locate = view.session.locateParagraph;
  const removed = view.session.selection()!.head.paraId;
  try {
    view.session.locateParagraph = (story, paraId) => {
      if (paraId === removed) throw new Error('Paragraph no longer exists');
      return locate(story, paraId);
    };
    await act(async () => {
      const context = view.result.current.host.activations[0]!.context;
      view.result.current.host.publishSelection();
      await context.run(() => {});
    });
    expect(view.errors).toEqual([]);
    expect(view.result.current.host.activations).toHaveLength(1);
    expect(view.result.current.host.activations[0]!.context.snapshot.selection).toEqual({ formatting: null, displayRange: null });
    expect(view.events.at(-1)).toEqual({ formatting: null, displayRange: null });
  } finally { view.session.locateParagraph = locate; }
});

for (const mappedBeforeDeadline of [true, false]) {
  test(`editor selection settles with its current display range (mapped before deadline: ${mappedBeforeDeadline})`, async () => {
    const clock = selectionClock();
    try {
      const view = await setup(false, null, 'body');
      const range = { from: 5, to: 9 };
      await act(async () => {
        view.selectRange(range);
        if (mappedBeforeDeadline) view.setRange(range);
        view.result.current.host.publishSelection();
      });
      await act(async () => { await clock.tick(349); });
      expect(view.settledSelections).toEqual([]);
      await act(async () => { await clock.tick(1); });
      if (!mappedBeforeDeadline) {
        expect(view.settledSelections).toEqual([]);
        expect(view.result.current.host.activations[0]!.context.snapshot.selectionSettled).toBe(false);
        await act(async () => {
          view.setRange(range);
          view.result.current.host.publishSelection();
        });
      }
      expect(view.settledSelections).toEqual([{
        formatting: view.formatting(),
        displayRange: { story: 'body', ...range,
          layoutId: view.result.current.host.activations[0]!.context.snapshot.layout!.id },
      }]);
      expect(view.result.current.host.activations[0]!.context.snapshot.selectionSettled).toBe(true);
      await act(async () => view.result.current.host.publishSelection());
      await act(async () => { await clock.tick(400); });
      expect(view.settledSelections).toHaveLength(1);
    } finally { cleanup(); clock.restore(); }
  });
}

for (const [kind, story] of [
  ['header', 'hf:rIdHeader'], ['footer', 'hf:rIdFooter'], ['footnote', 'fn:1'], ['endnote', 'en:1'],
] as const) {
  test(`${kind} caret movement settles once without a display range`, async () => {
    const clock = selectionClock();
    try {
      const view = await setup(false, null, story);
      await act(async () => { await clock.tick(350); });
      view.settledSelections.length = 0;
      await act(async () => view.selectRange({ from: 1, to: 1 }));
      expect(view.result.current.host.activations[0]!.context.snapshot.selectionSettled).toBe(false);
      expect(view.result.current.host.activations[0]!.context.snapshot.selection.displayRange).toBeNull();
      await act(async () => { await clock.tick(349); });
      expect(view.settledSelections).toEqual([]);
      await act(async () => { await clock.tick(1); });
      expect(view.session.selection()!.head).toMatchObject({ story, offset: 1 });
      expect(view.settledSelections).toEqual([{ formatting: view.formatting(), displayRange: null }]);
      expect(view.result.current.host.activations[0]!.context.snapshot.selectionSettled).toBe(true);
      await act(async () => view.result.current.host.publishSelection());
      await act(async () => { await clock.tick(700); });
      expect(view.settledSelections).toHaveLength(1);
      expect(view.result.current.host.activations[0]!.context.snapshot.selection.displayRange).toBeNull();
    } finally { cleanup(); clock.restore(); }
  });
}

test('a clear followed by a seed reseed settles the published empty selection once', async () => {
  const clock = selectionClock();
  try {
    const view = await setup(false, { from: 1, to: 3 }, 'body');
    const formatting: SelectionState = {
      hasSelection: true, isMultiParagraph: false, textFormatting: { bold: true },
      paragraphFormatting: {}, styleId: null, startParagraphIndex: 0, endParagraphIndex: 0,
    };
    await act(async () => view.result.current.tracker.handleSelectionChange(formatting));
    await act(async () => { await clock.tick(350); });
    view.events.length = 0;
    view.settledSelections.length = 0;
    const readSelection = view.session.selection;
    try {
      view.session.selection = () => null;
      await act(async () => {
        withYrsSelectionCause(view.session, 'seed', () => view.selectRange({ from: 0, to: 0 }));
        view.result.current.host.publishSelection();
        view.result.current.host.publishSelection();
      });
    } finally { view.session.selection = readSelection; }
    await act(async () => {
      withYrsSelectionCause(view.session, 'seed', () => view.setRange({ from: 0, to: 0 }));
      view.result.current.host.publishSelection();
      view.result.current.tracker.handleSelectionChange({ ...formatting, textFormatting: { italic: true } });
    });
    expect(view.session.selection()!.head.offset).toBe(0);
    expect(view.events).toEqual([{ formatting: null, displayRange: null }]);
    await act(async () => { await clock.tick(349); });
    expect(view.settledSelections).toEqual([]);
    await act(async () => { await clock.tick(1); });
    expect(view.settledSelections).toEqual([{ formatting: null, displayRange: null }]);
    expect(view.result.current.host.activations[0]!.context.snapshot.selection).toEqual({ formatting: null, displayRange: null });
    expect(view.result.current.host.activations[0]!.context.snapshot.selectionSettled).toBe(true);
    await act(async () => view.result.current.host.publishSelection());
    await act(async () => { await clock.tick(700); });
    expect(view.events).toHaveLength(1);
    expect(view.settledSelections).toHaveLength(1);
  } finally { cleanup(); clock.restore(); }
});

test('viewer provenance survives subscriber and layout publication without reading shell selection', async () => {
  const view = await setup(true);
  const selection: ViewerSelectionChange = {
    head: 12,
    displayRange: { story: 'body', from: 1, to: 12, layoutId: 'presented-layout' },
    isMultiParagraph: false,
  };
  const publish = (next: ViewerSelectionChange) => act(async () => {
    view.result.current.host.publishViewerSelection(next);
    view.result.current.tracker.handleViewerSelectionChange(next);
  });
  await publish(selection);
  expect(view.events).toEqual([{ formatting: null, displayRange: selection.displayRange }]);
  await publish(selection);
  view.rerender();
  await act(async () => view.result.current.host.publishSelection());
  expect(view.events).toHaveLength(1);
  expect(view.propSelections).toHaveLength(1);
  const beforeSettle = view.propSelections.length;
  await publish({ ...selection, isMultiParagraph: true });
  expect(view.propSelections).toHaveLength(beforeSettle + 1);
  expect(view.propSelections.at(-1)?.isMultiParagraph).toBe(true);
  expect(view.events).toHaveLength(1);
  await publish({ head: null, displayRange: null, isMultiParagraph: false });
  expect(view.propSelections.at(-1)).toEqual({
    hasSelection: false, isMultiParagraph: false, textFormatting: {}, paragraphFormatting: {},
    styleId: null, startParagraphIndex: -1, endParagraphIndex: -1,
  });
  expect(view.events.at(-1)).toEqual({ formatting: null, displayRange: null });
  expect(view.subscribedSelections).toEqual(view.propSelections);
  expect(view.result.current.host.activations[0]!.context.snapshot.selection).toEqual(view.events.at(-1)!);
  expect(view.shellReads()).toBe(0);
});

test('an editor session reads live formatting and does not notify for null tracker state', async () => {
  const view = await setup(false, { from: 5, to: 9 }, 'hf:stale', session => {
    const live = session.selection()!;
    const range = { story: live.head.story, start: live.anchor, end: live.head };
    session.formatRange(range, { bold: true });
    session.applyParagraphStyle(range, 'Heading1');
    session.setParagraphAttrs(range, { alignment: 'center' });
  });
  const selection = view.formatting();
  expect(selection).toMatchObject({ hasSelection: true, isMultiParagraph: false,
    textFormatting: { bold: true }, paragraphFormatting: { alignment: 'center' },
    styleId: 'Heading1', startParagraphIndex: 0, endParagraphIndex: 0 });
  await act(async () => view.result.current.tracker.handleSelectionChange(selection));
  expect(view.propSelections).toEqual([selection]);
  expect(view.subscribedSelections).toEqual([selection]);
  expect(view.propSelections[0]).toBe(selection);
  expect(view.events.at(-1)).toEqual({
    formatting: selection,
    displayRange: { story: 'hf:stale', from: 5, to: 9,
      layoutId: view.result.current.host.activations[0]!.context.snapshot.layout!.id },
  });
  const before = view.events.length;
  await act(async () => view.result.current.tracker.handleSelectionChange(null));
  expect(view.propSelections).toEqual([selection]);
  expect(view.subscribedSelections).toEqual([selection]);
  expect(view.events).toHaveLength(before);
});


test('the move entering a header pulls complete formatting and mapping once after all callers mark dirty', async () => {
  let headerId = '';
  const view = await setup(false, { from: 1, to: 1 }, 'body', session => {
    headerId = session.loadStories([{ storyId: 'hf:header', paragraphs: [{ text: 'Header formatting' }] }])['hf:header']![0]!;
    const range = { story: 'hf:header', start: { paraId: headerId, offset: 0 }, end: { paraId: headerId, offset: 17 } };
    session.formatRange(range, { bold: true, italic: true, fontSize: 18, color: { rgb: '123456' } });
    session.setParagraphAttrs(range, { alignment: 'center', spaceAfter: 120 });
  });
  view.events.length = 0;
  const staleFormatting = view.formatting();
  await act(async () => {
    view.session.setSelection({ story: 'hf:header', paraId: headerId, offset: 2 });
    view.result.current.host.publishSelection();
    view.result.current.tracker.handleSelectionChange(staleFormatting);
    view.setDisplayRange({ from: 3, to: 3 });
    view.result.current.host.publishSelection();
  });
  expect(view.events).toEqual([{
    formatting: view.formatting(),
    displayRange: { story: 'hf:header', from: 3, to: 3,
      layoutId: view.result.current.host.activations[0]!.context.snapshot.layout!.id },
  }]);
  expect(view.events[0]!.formatting).toMatchObject({
    textFormatting: { bold: true, italic: true, fontSize: 36, color: { rgb: '123456' } },
    paragraphFormatting: { alignment: 'center', spaceAfter: 120 },
    startParagraphIndex: 0, endParagraphIndex: 0,
  });
  expect(view.events[0]!.formatting).not.toEqual(staleFormatting);
  await act(async () => view.result.current.tracker.handleSelectionChange(view.formatting()));
  expect(view.events).toHaveLength(1);
});

test('a remote update and local caret move in one turn emit one local change and one settle', async () => {
  const clock = selectionClock();
  try {
    const view = await setup(false, { from: 1, to: 1 }, 'body');
    await act(async () => { await clock.tick(350); });
    view.events.length = 0;
    view.settledSelections.length = 0;
    const causes: Array<'remote' | 'seed' | undefined> = [];
    const host = view.result.current.host.host;
    const publish = host.selectionChanged;
    host.selectionChanged = (selection, identity, publication) => {
      causes.push(publication?.cause);
      return publish(selection, identity, publication);
    };
    const peer = await createYrsSession();
    try {
      peer.loadState(view.session.encodeState());
      const point = view.session.selection()!.head;
      peer.insertText({ ...point, offset: 0 }, 'remote ');
      await act(async () => {
        view.session.applyUpdate(peer.encodeStateAsUpdate(view.session.encodeStateVector()));
        stampSourceVersion(view.options.queries, view.session.version());
        view.selectRange({ from: 3, to: 3 });
        view.setDisplayRange({ from: 4, to: 4 });
        view.result.current.host.publishSelection();
      });
      expect(causes).toEqual([undefined]);
      expect(view.events).toEqual([{
        formatting: view.formatting(), displayRange: { story: 'body', from: 4, to: 4,
          layoutId: view.events[0]!.displayRange!.layoutId },
      }]);
      await act(async () => { await clock.tick(349); });
      expect(view.settledSelections).toEqual([]);
      await act(async () => { await clock.tick(1); });
      expect(view.settledSelections).toEqual(view.events);
      await act(async () => { await clock.tick(700); });
      expect(view.settledSelections).toHaveLength(1);
    } finally { peer.destroy(); }
  } finally { cleanup(); clock.restore(); }
});

test('setting a host selection and invoking a plugin synchronously flushes the new selection', async () => {
  const view = await setup(false, { from: 1, to: 1 }, 'body');
  let snapshot: DocxPluginSelection | undefined;
  let invoked = false;
  await act(async () => {
    view.selectRange({ from: 7, to: 7 });
    view.setDisplayRange({ from: 8, to: 8 });
    view.result.current.host.host.guard('test.selection', 'command', context => {
      invoked = true;
      snapshot = context.snapshot.selection;
    }, undefined);
    expect(invoked).toBe(true);
    expect(snapshot).toEqual({ formatting: view.formatting(),
      displayRange: { story: 'body', from: 8, to: 8, layoutId: snapshot!.displayRange!.layoutId } });
  });
  expect(view.events.at(-1)).toEqual(snapshot!);
});

for (const story of ['hf:header', 'hf:footer']) test(`a synchronous caret flush preserves a current ${story} layout`, async () => {
  const view = await setup(false, { from: 1, to: 1 }, story);
  const host = view.result.current.host.host;
  const retained = host.activations()[0]!.context;
  const layout = retained.snapshot.layout;
  expect(layout).not.toBeNull();
  expect(layout!.version).toBe(view.session.version());
  let current: DocxPluginContext<unknown> | undefined;
  view.events.length = 0;
  await act(async () => {
    view.selectRange({ from: 2, to: 2 });
    view.setDisplayRange({ from: 3, to: 3 });
    host.guard('test.selection', 'command', context => { current = context; }, undefined);
    expect(current!.snapshot.layout).toEqual(layout);
    expect(current!.snapshot.selection).toEqual({ formatting: view.formatting(),
      displayRange: { story, from: 3, to: 3, layoutId: layout!.id } });
  });
  expect(view.session.version()).toBe(layout!.version);
  expect(retained.snapshot.selection.displayRange).toMatchObject({ from: 1, to: 1 });
  expect(view.events).toEqual([current!.snapshot.selection]);
  expect(host.activations()[0]!.context.snapshot.layout).toEqual(layout);
});
