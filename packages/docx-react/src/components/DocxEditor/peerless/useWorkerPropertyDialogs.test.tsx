import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useWorkerPropertyDialogs } from './useWorkerPropertyDialogs';
import type { WorkerShellController } from './workerProperties';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, renderHook } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

function fixture() {
  const selection = { anchor: { story: 'body', paraId: 'p', offset: 1 }, focus: { story: 'body', paraId: 'p', offset: 4 } };
  const state = { status: 'ready', view: { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1, throughSeq: 1, workerSelection: selection, geometry: null } };
  const applyQuery = mock(async (_query: unknown, _intent: unknown) => ({ ok: true, status: 'executed' }));
  const readAt = mock(async () => ({ selectedText: 'Old' }));
  const submit = mock(async (_intent: unknown) => ({ ok: true, status: 'executed' }));
  const range = { story: 'body', start: { paraId: 'p', offset: 0 }, end: { paraId: 'p', offset: 5 } };
  const controller = { subscribe: () => () => {}, getSnapshot: () => state, applyQuery, readAt,
    submit,
    environment: () => ({ readOnly: false, mode: 'editing' }), ref: { focus() {}, flushPendingInput: async () => {} },
    query: async (request: { target?: { kind: string } }) => ({ version: 'captured-native', selection,
      value: request.target?.kind === 'object' ? { kind: 'object', handle: 'physical-second', properties: { width: 80, height: 40, alt: 'Old' } } :
        { kind: 'hyperlink', range, hyperlink: { href: 'https://example.com/old', tooltip: null } } }),
  } as unknown as WorkerShellController;
  const hook = renderHook(() => useWorkerPropertyDialogs(controller, {}));
  return { hook, applyQuery, selection, readAt, controller, range, submit };
}

test('caret link opens expanded, editable, and submits display text', async () => {
  const { hook, applyQuery, selection, readAt, range, submit } = fixture();
  selection.focus.offset = selection.anchor.offset;
  const expanded = { anchor: { story: range.story, ...range.start }, focus: { story: range.story, ...range.end } };
  await act(async () => { await hook.result.current.openLink(); });
  const dialog = hook.result.current.element.props.children[2];
  expect(dialog.props.displayTextReadOnly).toBeFalsy();
  expect(readAt).toHaveBeenCalledWith({ kind: 'rawSelectionInfo', selection: expanded }, 'captured-native');
  await act(async () => { await dialog.props.onSubmit({ url: 'https://example.com/new', tooltip: 'New', displayText: 'Changed' }); });
  expect(applyQuery.mock.calls[0]![1]).toEqual({ kind: 'setHyperlink', payload: { selection: expanded,
    hyperlink: { href: 'https://example.com/new', tooltip: 'New' }, displayText: 'Changed', expectVersion: 'captured-native' } });
  const end = { story: 'body', paraId: 'p', offset: 7 };
  expect(submit).toHaveBeenCalledWith({ kind: 'select', payload: { selection: { anchor: end, focus: end } } });
});

test('link dialog keeps edited values across worker view updates', async () => {
  const { hook } = fixture();
  await act(async () => { await hook.result.current.openLink(); });
  const view = render(hook.result.current.element);
  const text = view.getByLabelText('Display Text', { exact: true }) as HTMLInputElement;
  const url = view.getByLabelText('URL', { exact: true }) as HTMLInputElement;
  fireEvent.change(text, { target: { value: 'Changed' } });
  fireEvent.change(url, { target: { value: 'https://example.com/changed' } });
  hook.rerender();
  view.rerender(hook.result.current.element);
  expect(text.value).toBe('Changed');
  expect(url.value).toBe('https://example.com/changed');
});

test('bare caret link inserts display text and existing link removal uses its expanded range', async () => {
  const { hook, applyQuery, selection, controller, range } = fixture();
  selection.focus.offset = selection.anchor.offset;
  await act(async () => { await hook.result.current.openLink(); });
  await act(async () => { await hook.result.current.element.props.children[2].props.onRemove(); });
  expect(applyQuery.mock.calls[0]![1]).toEqual({ kind: 'setHyperlink', payload: {
    selection: { anchor: { story: range.story, ...range.start }, focus: { story: range.story, ...range.end } },
    hyperlink: null, expectVersion: 'captured-native',
  } });
  controller.query = (async () => ({ version: 'captured-native', selection, value: { kind: 'hyperlink',
    range: { story: 'body', start: selection.anchor, end: selection.focus }, hyperlink: null } })) as never;
  await act(async () => { await hook.result.current.openLink(); });
  await act(async () => { await hook.result.current.element.props.children[2].props.onSubmit({ url: 'https://example.com/new', displayText: 'New' }); });
  expect(applyQuery.mock.calls[1]![1]).toEqual({ kind: 'setHyperlink', payload: { selection,
    hyperlink: { href: 'https://example.com/new' }, displayText: 'New', expectVersion: 'captured-native' } });
});

test('popup link edits and removal query the expanded current caret through the barrier', async () => {
  const { hook, applyQuery, selection, range } = fixture();
  selection.focus.offset = selection.anchor.offset;
  const expanded = { anchor: { story: range.story, ...range.start }, focus: { story: range.story, ...range.end } };
  await act(async () => { expect(await hook.result.current.applyPopupLink('Popup', 'https://example.com/popup', 'Tip')).toBe(true); });
  expect(applyQuery.mock.calls[0]![1]).toEqual({ kind: 'setHyperlink', payload: { selection: expanded,
    hyperlink: { href: 'https://example.com/popup', tooltip: 'Tip' }, displayText: 'Popup', expectVersion: 'captured-native' } });
  await act(async () => { expect(await hook.result.current.removeLinkAtCaret()).toBe(true); });
  expect(applyQuery.mock.calls[1]![1]).toEqual({ kind: 'setHyperlink', payload: { selection: expanded,
    hyperlink: null, expectVersion: 'captured-native' } });
});

test('hyperlink edits apply in suggesting and refuse viewing or read-only', async () => {
  const { hook, controller, applyQuery } = fixture();
  for (const mode of ['suggesting', 'viewing', 'editing']) {
    controller.environment = (() => ({ mode, readOnly: mode === 'editing' })) as never;
    await act(async () => { await hook.result.current.openLink(); });
    await act(async () => { await hook.result.current.element.props.children[2].props.onSubmit({ url: 'https://example.com/new', displayText: 'New' }); });
  }
  expect(applyQuery).toHaveBeenCalledTimes(1);
});

test('popup link mutations retain the hybrid body-story restriction', async () => {
  const { hook, controller, selection, applyQuery } = fixture();
  selection.anchor.story = selection.focus.story = 'hf:header';
  controller.query = (async () => ({ version: 'captured-native', selection, value: { kind: 'hyperlink',
    range: { story: 'hf:header', start: selection.anchor, end: selection.focus }, hyperlink: { href: 'https://example.com/header' } } })) as never;
  await act(async () => { expect(await hook.result.current.applyPopupLink('Header', 'https://example.com/new')).toBe(false); });
  await act(async () => { expect(await hook.result.current.removeLinkAtCaret()).toBe(false); });
  expect(applyQuery).not.toHaveBeenCalled();
});

test('split cell dialog opens with the span defaults and applies the chosen numbers', async () => {
  const { hook } = fixture();
  const apply = mock(async (_rows: number, _columns: number) => ({ ok: true, status: 'executed' }));
  await act(async () => {
    (hook.result.current as unknown as { openSplitCell(defaults: unknown, apply: unknown): void })
      .openSplitCell({ initialRows: 2, initialCols: 2, minRows: 2, minCols: 1 }, apply);
  });
  const find = (node: unknown): { props: Record<string, unknown> } | undefined => {
    if (!node || typeof node !== 'object') return undefined;
    const props = (node as { props?: Record<string, unknown> }).props;
    if (props && 'initialRows' in props) return node as { props: Record<string, unknown> };
    const children = props?.children;
    for (const child of Array.isArray(children) ? children : [children]) {
      const found = find(child);
      if (found) return found;
    }
    return undefined;
  };
  const dialog = find(hook.result.current.element);
  expect(dialog?.props).toMatchObject({ isOpen: true, initialRows: 2, initialCols: 2, minRows: 2, minCols: 1 });
  await act(async () => { await (dialog!.props.onApply as (rows: number, columns: number) => void)(2, 3); });
  expect(apply).toHaveBeenCalledWith(2, 3);
});

test('cancelling the split cell dialog submits no split command', async () => {
  const { hook } = fixture();
  const apply = mock(async (_rows: number, _columns: number) => ({ ok: true, status: 'executed' }));
  await act(async () => {
    (hook.result.current as unknown as { openSplitCell(defaults: unknown, apply: unknown): void })
      .openSplitCell({ initialRows: 2, initialCols: 2, minRows: 2, minCols: 1 }, apply);
  });
  const find = (node: unknown): { props: Record<string, unknown> } | undefined => {
    if (!node || typeof node !== 'object') return undefined;
    const props = (node as { props?: Record<string, unknown> }).props;
    if (props && 'initialRows' in props) return node as { props: Record<string, unknown> };
    const children = props?.children;
    for (const child of Array.isArray(children) ? children : [children]) {
      const found = find(child);
      if (found) return found;
    }
    return undefined;
  };
  const dialog = find(hook.result.current.element);
  expect(dialog?.props).toMatchObject({ isOpen: true, initialRows: 2, initialCols: 2 });
  await act(async () => { (dialog!.props.onClose as () => void)(); });
  expect(apply).not.toHaveBeenCalled();
  expect(find(hook.result.current.element)).toBeUndefined();
});

test('image properties retain the opened physical handle after an adjacent image is selected', async () => {
  const { hook, applyQuery, selection, controller } = fixture();
  selection.focus.offset = 2;
  const query = mock(async () => ({ version: 'captured-native', selection,
    value: { kind: 'object', handle: 'physical-first', properties: { width: 80, height: 40, alt: 'First' } } }));
  controller.query = query as never;
  await act(async () => { await Reflect.get(hook.result.current, 'openImage')(); });
  const dialog = hook.result.current.element.props.children.find((child: { props?: { currentData?: unknown } } | null) => child?.props?.currentData);
  expect(dialog.props.currentData.alt).toBe('First');
  selection.anchor.offset = 2;
  selection.focus.offset = 3;
  await act(async () => { await dialog.props.onApply({ alt: 'Edited first' }); });
  expect(applyQuery.mock.calls[0]![1]).toMatchObject({ kind: 'setObjectProperties', payload: {
    handle: 'physical-first', patch: { alt: 'Edited first' }, expectVersion: 'captured-native',
  } });
});

test('image position floats an inline image with the wrap property the worker accepts', async () => {
  const { hook, applyQuery, controller, selection } = fixture();
  controller.query = (async () => ({ version: 'captured-native', selection,
    value: { kind: 'object', handle: 'physical-first', properties: { width: 80, height: 40, wrapType: 'inline' } } })) as never;
  await act(async () => { await hook.result.current.openImagePosition(); });
  const dialog = hook.result.current.element.props.children.find((child: { props?: { currentData?: unknown } } | null) => child?.props?.currentData);
  const placement = { horizontal: { relativeTo: 'column', align: 'left' }, vertical: { relativeTo: 'paragraph', align: 'top' },
    distTop: 0, distBottom: 0, distLeft: 0, distRight: 0 };
  await act(async () => { await dialog.props.onApply(placement); });
  expect(applyQuery.mock.calls[0]![1]).toEqual({ kind: 'setObjectProperties', payload: { handle: 'physical-first', expectVersion: 'captured-native',
    patch: { position: { horizontal: placement.horizontal, vertical: placement.vertical }, distTop: 0, distBottom: 0, distLeft: 0, distRight: 0, wrap: 'square' } } });
});
