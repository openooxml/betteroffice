import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { evaluateDocxCommand } from '../../../commands/evaluate';
import { useWorkerContextMenu } from './useWorkerContextMenu';
import {
  buildImageTransformPatch,
  buildImageWrapPatch,
  imageWrapTargetFromAttrs,
  useWorkerPropertyDialogs,
} from './useWorkerPropertyDialogs';
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

test('link dialog keeps edited values when an earlier open resolves late', async () => {
  const { hook, controller } = fixture();
  let resolveText!: (value: { selectedText: string }) => void;
  const pendingText = new Promise<{ selectedText: string }>(resolve => { resolveText = resolve; });
  controller.readAt = mock(() => pendingText) as never;
  let earlier!: ReturnType<typeof hook.result.current.openLink>;
  await act(async () => { earlier = hook.result.current.openLink(); });
  controller.readAt = mock(async () => ({ selectedText: 'Old' })) as never;
  await act(async () => { await hook.result.current.openLink(); });
  const opened = hook.result.current.element.props.children[2];
  const view = render(hook.result.current.element);
  const text = view.getByLabelText('Display Text', { exact: true }) as HTMLInputElement;
  const url = view.getByLabelText('URL', { exact: true }) as HTMLInputElement;
  fireEvent.change(text, { target: { value: 'Changed' } });
  fireEvent.change(url, { target: { value: 'https://example.com/changed' } });
  await act(async () => { resolveText({ selectedText: 'Old' }); await earlier; });
  view.rerender(hook.result.current.element);
  expect(text.value).toBe('Changed');
  expect(url.value).toBe('https://example.com/changed');
  expect(hook.result.current.element.props.children[2].props.initialData).toBe(opened.props.initialData);
  await act(async () => { await hook.result.current.openLink(); });
  view.rerender(hook.result.current.element);
  expect(text.value).toBe('Changed');
  expect(url.value).toBe('https://example.com/changed');
  expect(hook.result.current.element.props.children[2].props.initialData).toBe(opened.props.initialData);
});

test('changing the worker session invalidates unfinished link opens', async () => {
  const { hook, controller } = fixture();
  let resolveText!: (value: { selectedText: string }) => void;
  const pendingText = new Promise<{ selectedText: string }>(resolve => { resolveText = resolve; });
  controller.readAt = mock(() => pendingText) as never;
  let pending!: ReturnType<typeof hook.result.current.openLink>;
  await act(async () => { pending = hook.result.current.openLink(); });
  controller.getSnapshot().view!.identity.epoch++;
  hook.rerender();
  await act(async () => { resolveText({ selectedText: 'Old' }); await pending; });
  expect(hook.result.current.element.props.children[2]).toBeNull();
  controller.readAt = mock(async () => ({ selectedText: 'Old' })) as never;
  await act(async () => { await hook.result.current.openLink(); });
  expect(hook.result.current.element.props.children[2].props.initialData.displayText).toBe('Old');
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

test('computes imageState.wrap and submits setObjectProperties patches for applyImageWrap and applyImageTransform', async () => {
  expect(imageWrapTargetFromAttrs({ wrapType: 'inline' })).toBe('inline');
  expect(imageWrapTargetFromAttrs({ wrapType: 'square', displayMode: 'float', cssFloat: 'left' })).toBe('squareLeft');
  expect(imageWrapTargetFromAttrs({ wrapType: 'square', displayMode: 'float', cssFloat: 'right' })).toBe('squareRight');
  expect(imageWrapTargetFromAttrs({ wrapType: 'behind', displayMode: 'float', cssFloat: 'none' })).toBe('behind');
  expect(buildImageTransformPatch({}, 'rotateCW')).toEqual({ transform: 'rotate(90deg)' });
  expect(buildImageTransformPatch({ transform: 'rotate(90deg)' }, 'flipH')).toEqual({ transform: 'rotate(90deg) scaleX(-1)' });
  expect(buildImageTransformPatch({ transform: 'rotate(90deg)' }, 'rotateCCW')).toEqual({ transform: null });

  const { applyQuery, controller, selection } = fixture();
  let currentProps: Record<string, unknown> = { width: 80, height: 40, wrapType: 'square', displayMode: 'float', cssFloat: 'left' };
  controller.query = (async () => ({
    version: 'captured-native',
    selection,
    value: { kind: 'object', handle: 'physical-first', properties: currentProps },
  })) as never;
  const hook = renderHook(() => useWorkerPropertyDialogs(controller, {}));
  await act(async () => {});
  expect(hook.result.current.imageState).toEqual({ wrap: 'squareLeft' });

  await act(async () => { await hook.result.current.applyImageWrap('behind'); });
  expect(applyQuery.mock.calls[0]![1]).toEqual({
    kind: 'setObjectProperties',
    payload: {
      handle: 'physical-first',
      expectVersion: 'captured-native',
      patch: buildImageWrapPatch(currentProps, 'behind'),
    },
  });

  currentProps = { ...currentProps, transform: 'rotate(90deg)' };
  await act(async () => { await hook.result.current.applyImageTransform('flipH'); });
  expect(applyQuery.mock.calls[1]![1]).toEqual({
    kind: 'setObjectProperties',
    payload: {
      handle: 'physical-first',
      expectVersion: 'captured-native',
      patch: { transform: 'rotate(90deg) scaleX(-1)' },
    },
  });

  await act(async () => { await hook.result.current.applyImageWrap('inline', 'context-menu-handle'); });
  expect(applyQuery.mock.calls[2]![1]).toEqual({
    kind: 'setObjectProperties',
    payload: {
      handle: 'context-menu-handle',
      expectVersion: 'captured-native',
      patch: buildImageWrapPatch(currentProps, 'inline'),
    },
  });
});

test('applyImageWrap writes margin alignment without a horizontal offset for both side presets', async () => {
  const { hook, applyQuery, controller, selection } = fixture();
  let currentProps: Record<string, unknown> = { width: 80, height: 40, wrapType: 'inline' };
  controller.query = (async () => ({ version: 'captured-native', selection,
    value: { kind: 'object', handle: 'physical-first', properties: currentProps } })) as never;
  for (const [target, side, wrapText] of [
    ['squareRight', 'right', 'left'],
    ['squareLeft', 'left', 'right'],
  ] as const) {
    await act(async () => { expect(await hook.result.current.applyImageWrap(target)).toBe(true); });
    const patch = {
      wrap: 'square', wrapType: 'square', displayMode: 'float', cssFloat: side, wrapText,
      position: { horizontal: { relativeTo: 'margin', align: side }, vertical: { relativeTo: 'paragraph', posOffset: 0 } },
    };
    expect(applyQuery.mock.calls.at(-1)![1]).toEqual({ kind: 'setObjectProperties', payload: {
      handle: 'physical-first', expectVersion: 'captured-native', patch,
    } });
    currentProps = { ...currentProps, ...patch };
  }
});

test('buildImageWrapPatch omits wrap distance keys for non-inline targets so existing distances survive', () => {
  const currentWithDistances = {
    wrapType: 'inline',
    distTop: 12,
    distBottom: 24,
    distLeft: 36,
    distRight: 48,
  };
  for (const target of ['squareLeft', 'behind'] as const) {
    const patch = buildImageWrapPatch(currentWithDistances, target);
    expect('distTop' in patch).toBe(false);
    expect('distBottom' in patch).toBe(false);
    expect('distLeft' in patch).toBe(false);
    expect('distRight' in patch).toBe(false);
  }
  const inlinePatch = buildImageWrapPatch(currentWithDistances, 'inline');
  expect(inlinePatch).toMatchObject({
    distTop: null,
    distBottom: null,
    distLeft: null,
    distRight: null,
  });
});

test('image context-menu wrap action refuses in header, footer, note stories and in suggesting mode without sending a worker mutation', async () => {
  const { hook: dialogsHook, applyQuery, controller, selection, submit } = fixture();
  let currentStory = 'hf:header';
  let currentMode: 'editing' | 'suggesting' = 'editing';
  const translate = (key: string) => key;

  const createViewSnapshot = () => ({
    identity: { session: 's', incarnation: 1, epoch: 1 },
    version: 1,
    throughSeq: 1,
    workerSelection: {
      anchor: { story: currentStory, paraId: 'p', offset: 0 },
      focus: { story: currentStory, paraId: 'p', offset: 1 },
    },
    geometry: {
      frameEpoch: 2,
      layoutRevision: 3,
      documentVersion: 'native',
      pages: [
        {
          page: { pageId: 'p', pageIndex: 0 },
          bounds: { x: 0, y: 0, width: 600, height: 800 },
        },
      ],
      objects: {
        version: 'native',
        frameEpoch: 2,
        layoutRevision: 3,
        objects: [
          {
            handle: 'image-handle-1',
            kind: 'image' as const,
            story: currentStory,
            rects: [{ pageIndex: 0, x: 10, y: 10, width: 100, height: 80 }],
          },
        ],
      },
    },
  });

  const getEnvironment = () => ({
    status: 'ready' as const,
    readOnly: false,
    mode: currentMode,
    bodyStory: currentStory === 'body',
    image: { wrap: null },
    translate,
  });

  let snapshot = { status: 'ready', view: createViewSnapshot() };
  const getViewSnapshot = () => snapshot.view;
  controller.getSnapshot = (() => snapshot) as never;
  controller.environment = getEnvironment as never;
  const binding = { viewStore: { getSnapshot: getViewSnapshot } };

  const onWrapImage = mock(
    async (
      handle: string,
      wrap: 'inline' | 'squareLeft' | 'squareRight' | 'topAndBottom' | 'behind' | 'inFront',
    ) => dialogsHook.result.current.applyImageWrap(wrap, handle),
  );
  const errors: Error[] = [];
  const onError: NonNullable<Parameters<typeof useWorkerContextMenu>[0]['onError']> = (err) => {
    errors.push(err);
  };
  const menuHook = renderHook(() =>
    useWorkerContextMenu({
      binding: () => binding,
      environment: getEnvironment,
      commands: { getState: () => ({ enabled: false }) },
      clipboard: async () => {},
      submit,
      onWrapImage,
      onError,
    } as unknown as Parameters<typeof useWorkerContextMenu>[0]),
  );

  const pageEl = document.createElement('div');
  pageEl.className = 'paged-editor__page';
  pageEl.dataset.pageIndex = '0';
  pageEl.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: 600,
      bottom: 800,
      width: 600,
      height: 800,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    }) as DOMRect;
  document.body.appendChild(pageEl);

  const readContextMenu = () =>
    (
      menuHook.result.current.element as unknown as {
        props: {
          children: Array<{
            props: {
              items: Array<{ action: string; disabled?: boolean }>;
              onAction: (action: string) => void;
            };
          } | null>;
        };
      }
    ).props.children[0]!;

  for (const story of ['hf:header', 'hf:footer', 'fn:1', 'en:1'] as const) {
    currentStory = story;
    currentMode = 'editing';
    selection.anchor.story = story;
    selection.focus.story = story;
    snapshot = { status: 'ready', view: createViewSnapshot() };
    errors.length = 0;

    act(() => {
      menuHook.result.current.open({
        preventDefault() {},
        stopPropagation() {},
        clientX: 30,
        clientY: 30,
        target: pageEl,
      } as never);
    });

    const contextMenu = readContextMenu();
    const wrapItems = contextMenu.props.items.filter((item) => item.action.startsWith('wrap:'));
    expect(wrapItems.length).toBeGreaterThan(0);
    for (const item of wrapItems) {
      expect(item.disabled).toBe(true);
    }

    const expectedGate = evaluateDocxCommand(
      'imageWrap',
      { wrap: 'squareLeft' },
      {
        status: 'ready',
        readOnly: false,
        mode: 'editing',
        bodyStory: false,
        image: { wrap: null },
        translate,
      } as never,
    );
    expect(expectedGate).toMatchObject({
      enabled: false,
      disabledReason: { code: 'unsupported-story' },
    });

    await act(async () => {
      contextMenu.props.onAction('wrap:squareLeft');
    });

    expect(errors.map((e) => e.message)).toEqual([
      expectedGate.enabled ? '' : expectedGate.disabledReason.message,
    ]);
    expect(onWrapImage).not.toHaveBeenCalled();
    expect(applyQuery).not.toHaveBeenCalled();
  }

  currentStory = 'body';
  currentMode = 'suggesting';
  selection.anchor.story = 'body';
  selection.focus.story = 'body';
  snapshot = { status: 'ready', view: createViewSnapshot() };
  errors.length = 0;

  act(() => {
    menuHook.result.current.open({
      preventDefault() {},
      stopPropagation() {},
      clientX: 30,
      clientY: 30,
      target: pageEl,
    } as never);
  });

  const suggestingMenu = readContextMenu();
  const suggestingWrapItems = suggestingMenu.props.items.filter((item) =>
    item.action.startsWith('wrap:'),
  );
  expect(suggestingWrapItems.length).toBeGreaterThan(0);
  for (const item of suggestingWrapItems) {
    expect(item.disabled).toBe(true);
  }

  const expectedSuggestingGate = evaluateDocxCommand(
    'imageWrap',
    { wrap: 'squareLeft' },
    {
      status: 'ready',
      readOnly: false,
      mode: 'suggesting',
      bodyStory: true,
      image: { wrap: null },
      translate,
    } as never,
  );
  expect(expectedSuggestingGate).toMatchObject({
    enabled: false,
    disabledReason: { code: 'suggesting-unsupported' },
  });

  await act(async () => {
    suggestingMenu.props.onAction('wrap:squareLeft');
  });

  expect(errors.map((e) => e.message)).toEqual([
    expectedSuggestingGate.enabled ? '' : expectedSuggestingGate.disabledReason.message,
  ]);
  expect(onWrapImage).not.toHaveBeenCalled();
  expect(applyQuery).not.toHaveBeenCalled();
  pageEl.remove();
});

