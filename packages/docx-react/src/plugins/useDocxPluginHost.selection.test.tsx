import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import { useSelectionTracker } from '../components/DocxEditor/hooks/useSelectionTracker';
import { stampRevisionPreviewKey, stampSourceVersion } from '../components/DocxEditor/internals/layoutProvenance';
import type { ViewerSelectionChange } from '../components/DocxEditor/internals/viewerSelectionController';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { SelectionState } from '../components/DocxEditor/types';
import { defineDocxPlugin } from './defineDocxPlugin';
import type { DocxPluginSelection } from './types';
import { useDocxPluginHost } from './useDocxPluginHost';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function setup(viewerSelection: boolean) {
  let shellReads = 0;
  const session = {
    version: () => 'v1',
    getProposals: () => ({ version: 'v1', previewVersion: 0, proposals: [] }),
    onUpdate: () => () => {},
    onProposalChange: () => () => {},
    selection: () => { shellReads += 1; return { head: { story: 'header:stale' } }; },
  } as unknown as YrsSession;
  const queries = {
    displayList: { pages: [] },
    sourceState: () => ({ status: 'ready' }),
    pageCount: () => 0,
  } as unknown as DisplayListQueries;
  stampSourceVersion(queries, 'v1');
  stampRevisionPreviewKey(queries, '');
  const editor = {
    getYrsSession: () => session,
    getSelectionRange: () => ({ from: 5, to: 9 }),
    getLayout: () => null,
  } as unknown as PagedEditorRef;
  const events: DocxPluginSelection[] = [];
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
    },
  });
  const options = {
    plugins: [plugin], pagedEditorRef: { current: editor },
    writeModeRef: { current: 'viewing' as const }, mode: 'viewing' as const,
    readOnly: true, viewerSelection, commands: createDocxCommandController(),
    session, loadGeneration: 0, queries, layoutError: null, zoom: 1,
    canvasHostRef: { current: null }, overlayTarget: null,
    selectionChangeSubscribersRef: subscribers,
    i18n: undefined, onRenderedDomContextReady: undefined,
  };
  const borderSpecRef = { current: { style: 'single', size: 4, color: { rgb: '000000' } } };
  const { result, rerender } = renderHook(() => {
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
  return { result, rerender, events, propSelections, subscribedSelections,
    shellReads: () => shellReads };
}

test('viewer provenance survives subscriber and layout publication without reading shell selection', async () => {
  const view = await setup(true);
  const selection: ViewerSelectionChange = {
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
  await publish({ displayRange: null, isMultiParagraph: false });
  expect(view.propSelections.at(-1)).toEqual({
    hasSelection: false, isMultiParagraph: false, textFormatting: {}, paragraphFormatting: {},
    styleId: null, startParagraphIndex: -1, endParagraphIndex: -1,
  });
  expect(view.events.at(-1)).toEqual({ formatting: null, displayRange: null });
  expect(view.subscribedSelections).toEqual(view.propSelections);
  expect(view.result.current.host.activations[0]!.context.snapshot.selection).toEqual(view.events.at(-1)!);
  expect(view.shellReads()).toBe(0);
});

test('an editor session preserves formatting payloads and does not notify for null tracker state', async () => {
  const view = await setup(false);
  const selection: SelectionState = {
    hasSelection: true, isMultiParagraph: false,
    textFormatting: { bold: true }, paragraphFormatting: { alignment: 'center' },
    styleId: 'Heading1', startParagraphIndex: 2, endParagraphIndex: 2,
  };
  await act(async () => view.result.current.tracker.handleSelectionChange(selection));
  expect(view.propSelections).toEqual([selection]);
  expect(view.subscribedSelections).toEqual([selection]);
  expect(view.propSelections[0]).toBe(selection);
  expect(view.events.at(-1)).toEqual({
    formatting: selection,
    displayRange: { story: 'header:stale', from: 5, to: 9,
      layoutId: view.result.current.host.activations[0]!.context.snapshot.layout!.id },
  });
  const before = view.events.length;
  await act(async () => view.result.current.tracker.handleSelectionChange(null));
  expect(view.propSelections).toEqual([selection]);
  expect(view.subscribedSelections).toEqual([selection]);
  expect(view.events).toHaveLength(before);
});
