import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { DOCX_COMMAND_DESCRIPTORS } from '../../commands/descriptors';
import { definitionProblem, pluginDefinition } from '../defineDocxPlugin';
import { pluginSelectionKeys } from '../selectionIdentity';
import type {
  DocxAnchorGeometryResult,
  DocxPluginContext,
  DocxPluginDefinition,
  DocxPluginGeometry,
  DocxPluginSelection,
  DocxPluginSnapshot,
} from '../types';
import { DOCX_MINI_TOOLBAR_ID, docxMiniToolbar } from '../../mini-toolbar';
import { INITIAL_MINI_TOOLBAR_STATE, type MiniToolbarState } from './model';

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

type Context = DocxPluginContext<MiniToolbarState>;
type CommandStates = Record<string, { enabled: boolean; active?: boolean | 'mixed'; disabledReason?: { code: string; message: string } }>;

const REFUSED = { enabled: false, disabledReason: { code: 'permission-denied', message: 'not granted' } };
type FocusAnswer = { ok: true } | { ok: false; failure: { code: string; message: string } };
const NO_INPUT_REFUSAL: FocusAnswer = { ok: false, failure: { code: 'unsupported-policy', message: 'The document input cannot take focus now' } };

function definitionOf(plugin = docxMiniToolbar()): DocxPluginDefinition<MiniToolbarState> {
  return pluginDefinition(plugin) as unknown as DocxPluginDefinition<MiniToolbarState>;
}

function snapshot(overrides: Partial<DocxPluginSnapshot> = {}): DocxPluginSnapshot {
  const value = {
    focused: true,
    inputFocused: true,
    locale: 'en',
    generation: 'document-1',
    version: '1',
    mode: 'editing',
    readOnly: false,
    grant: {},
    selection: { formatting: { hasSelection: true }, displayRange: { story: 'body', from: 4, to: 9, layoutId: 'layout-1' } },
    selectionSettled: true,
    layout: { id: 'layout-1', version: '1', previewVersion: 0, zoom: 1, pageCount: 1 },
    ...overrides,
  } as unknown as DocxPluginSnapshot;
  pluginSelectionKeys.set(value, 'selection-1');
  return value;
}

const PAGE = { x: 100, y: 0, width: 700, height: 1000 };

function geometry(rect = { x: 300, y: 300, width: 80, height: 20 }): DocxPluginGeometry {
  const result: DocxAnchorGeometryResult = {
    ok: true, version: '1', previewVersion: 0, layoutId: 'layout-1',
    rects: [{ ...rect, pageIndex: 0 }],
    anchor: { x: rect.x + rect.width, y: rect.y, width: 0, height: rect.height, pageIndex: 0 },
    pageRect: PAGE,
  };
  return {
    viewport: { x: 0, y: 0, width: 1000, height: 800 },
    getAnchorGeometry: () => result,
    getPageRect: () => ({ ...PAGE, pageIndex: 0, built: true }),
  } as unknown as DocxPluginGeometry;
}

function context(options: { states?: CommandStates; state?: Partial<MiniToolbarState>; snapshot?: Partial<DocxPluginSnapshot>; focus?: () => FocusAnswer } = {}) {
  const calls = { executed: [] as string[], focused: 0, focusedFrom: [] as (Element | null)[], states: [] as MiniToolbarState[] };
  let current: MiniToolbarState = { ...INITIAL_MINI_TOOLBAR_STATE, ...options.state };
  const value = {
    pluginId: DOCX_MINI_TOOLBAR_ID,
    snapshot: snapshot(options.snapshot),
    get state() { return current; },
    setState(next: MiniToolbarState | ((previous: MiniToolbarState) => MiniToolbarState)) {
      current = typeof next === 'function' ? next(current) : next;
      calls.states.push(current);
      return true;
    },
    commands: {
      getState: (id: string) => options.states?.[id] ?? REFUSED,
      getDescriptor: (id: keyof typeof DOCX_COMMAND_DESCRIPTORS) => DOCX_COMMAND_DESCRIPTORS[id],
      subscribe: () => () => {},
      execute: async (id: string) => {
        calls.executed.push(id);
        return { ok: true, status: 'executed' };
      },
    },
    focusDocument: () => {
      calls.focused += 1;
      calls.focusedFrom.push(document.activeElement);
      if (options.focus) return options.focus();
      return { ok: true };
    },
    run: async (action: (next: Context) => unknown) => { await action(value as unknown as Context); },
  };
  return { context: value as unknown as Context, calls };
}

function renderOverlay(plugin: ReturnType<typeof docxMiniToolbar>, ctx: Context, geo = geometry()) {
  const Overlay = definitionOf(plugin).overlay!;
  const view = render(<Overlay context={ctx} geometry={geo} />);
  return Object.assign(view, { update: (next: Context) => view.rerender(<Overlay context={next} geometry={geo} />) });
}

function pressListeners(target: Document) {
  const add = target.addEventListener;
  const remove = target.removeEventListener;
  const counts = { added: 0, removed: 0 };
  const isPress = (type: string, options?: boolean | EventListenerOptions) =>
    type === 'pointerdown' && (options === true || (typeof options === 'object' && options.capture === true));
  target.addEventListener = function (this: Document, type: string, listener: EventListenerOrEventListenerObject, options?: boolean | AddEventListenerOptions) {
    if (isPress(type, options)) counts.added += 1;
    return add.call(this, type, listener, options);
  } as typeof target.addEventListener;
  target.removeEventListener = function (this: Document, type: string, listener: EventListenerOrEventListenerObject, options?: boolean | EventListenerOptions) {
    if (isPress(type, options)) counts.removed += 1;
    return remove.call(this, type, listener, options);
  } as typeof target.removeEventListener;
  return {
    counts,
    attached: () => counts.added - counts.removed,
    restore() {
      target.addEventListener = add;
      target.removeEventListener = remove;
    },
  };
}

const GRANTED: CommandStates = {
  bold: { enabled: true, active: true },
  italic: { enabled: true, active: 'mixed' },
  underline: { enabled: true, active: false },
  strikethrough: { enabled: true, active: false },
  clearFormatting: { enabled: true },
  pasteAsPlainText: { enabled: true },
};

describe('docxMiniToolbar', () => {
  test('is a valid plugin with an overlay and one focus command bound to Alt+F10', () => {
    const plugin = docxMiniToolbar();
    expect(plugin.id).toBe(DOCX_MINI_TOOLBAR_ID);
    const definition = definitionOf(plugin);
    expect(definitionProblem(definition as unknown as DocxPluginDefinition<unknown>)).toBeNull();
    expect(definition.overlay).toBeDefined();
    expect(definition.commands?.map(({ id, mutatesDocument, shortcuts }) => ({ id, mutatesDocument, shortcuts }))).toEqual([
      { id: 'focus', mutatesDocument: false, shortcuts: ['Alt+F10'] },
    ]);
    expect(definition.toolbar).toBeUndefined();
  });

  test('refuses unknown item ids when it is created', () => {
    expect(() => docxMiniToolbar({ items: ['bold', 'insertLink' as never] })).toThrow(TypeError);
  });

  test('enables its focus command only while the toolbar is shown for an eligible selection', () => {
    const [focus] = definitionOf().commands!;
    const hidden = context();
    const shown = context({ state: { shown: true } });
    const blurred = context({ state: { shown: true }, snapshot: { inputFocused: false } });
    const typing = context({ state: { shown: true }, snapshot: { selectionSettled: false } });
    const dismissed = context({ state: { shown: true, dismissed: '["body",4,9]' } });
    expect(focus!.getState!(shown.context)).toEqual({ enabled: true });
    for (const { context: ctx } of [hidden, blurred, typing, dismissed]) {
      const state = focus!.getState!(ctx);
      expect(state.enabled).toBe(false);
      expect(state.disabledReason?.code).toBe('not-shown');
    }
  });

  test('the focus command asks the toolbar for focus and leaves it shown', async () => {
    const [focus] = definitionOf().commands!;
    const requested = context({ state: { shown: true } });
    expect(await focus!.execute(requested.context)).toEqual({ ok: true, status: 'requested' });
    expect(requested.context.state).toEqual({ ...INITIAL_MINI_TOOLBAR_STATE, shown: true, focusRequest: INITIAL_MINI_TOOLBAR_STATE.focusRequest + 1 });
    expect(requested.calls.focused).toBe(0);
  });

  test('a new selection clears the dismissal', async () => {
    const { context: ctx } = context({ state: { dismissed: '["body",4,9]' } });
    await definitionOf().onEvent!(ctx, {
      type: 'selection-change', generation: 'document-1', version: '1',
      selection: { formatting: null, displayRange: { story: 'body', from: 10, to: 14, layoutId: 'layout-1' } },
    });
    expect(ctx.state.dismissed).toBeNull();
  });
});

describe('the overlay', () => {
  test('without a grant shows every button disabled and does not throw', () => {
    const { context: ctx } = context();
    const view = renderOverlay(docxMiniToolbar(), ctx);
    const toolbar = view.getByRole('toolbar', { name: 'Selection formatting' });
    const buttons = toolbar.querySelectorAll('button');
    expect([...buttons].map((button) => button.getAttribute('aria-label'))).toEqual([
      'Bold', 'Italic', 'Underline', 'Strikethrough', 'Clear formatting', 'Paste as Plain Text',
    ]);
    expect([...buttons].every((button) => button.disabled)).toBe(true);
  });

  test('shows the host items in the host order with pressed states', () => {
    const { context: ctx } = context({ states: GRANTED });
    const view = renderOverlay(docxMiniToolbar({ items: ['italic', 'bold', 'pasteAsPlainText'] }), ctx);
    const buttons = [...view.getByRole('toolbar').querySelectorAll('button')];
    expect(buttons.map((button) => button.getAttribute('aria-label'))).toEqual(['Italic', 'Bold', 'Paste as Plain Text']);
    expect(buttons.map((button) => button.getAttribute('aria-pressed'))).toEqual(['mixed', 'true', null]);
    expect(buttons.every((button) => !button.disabled)).toBe(true);
  });

  test('is placed above the selection and inside the page, clear of the comment button', () => {
    const { context: ctx } = context({ states: GRANTED });
    const view = renderOverlay(docxMiniToolbar(), ctx, geometry({ x: 740, y: 300, width: 50, height: 20 }));
    const toolbar = view.getByRole('toolbar') as HTMLElement;
    const left = parseFloat(toolbar.style.left);
    const top = parseFloat(toolbar.style.top);
    const width = parseFloat(toolbar.style.width);
    const height = parseFloat(toolbar.style.height);
    expect(top + height).toBeLessThanOrEqual(300);
    expect(left + width).toBeLessThanOrEqual(PAGE.x + PAGE.width);
    expect(toolbar.style.pointerEvents).toBe('auto');
  });

  test('renders nothing for a selection that is not eligible', () => {
    for (const overrides of [{ selectionSettled: false }, { mode: 'viewing' as const }, { readOnly: true }, { inputFocused: false }]) {
      const { context: ctx } = context({ states: GRANTED, snapshot: overrides });
      const view = renderOverlay(docxMiniToolbar(), ctx);
      expect(view.queryByRole('toolbar')).toBeNull();
      cleanup();
    }
  });

  test('a click runs the command without taking focus from the document', () => {
    const { context: ctx, calls } = context({ states: GRANTED });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    const bold = view.getByRole('button', { name: 'Bold' });
    expect(fireEvent.mouseDown(bold)).toBe(false);
    fireEvent.click(bold);
    expect(calls.executed).toEqual(['bold']);
  });

  test('Tab in the toolbar returns focus to the document and keeps the toolbar', () => {
    const { context: ctx, calls } = context({ states: GRANTED, state: { shown: true } });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    const bold = view.getByRole('button', { name: 'Bold' });
    act(() => bold.focus());
    expect(fireEvent.keyDown(bold, { key: 'Tab' })).toBe(false);
    expect(calls.focused).toBe(1);
    expect(ctx.state.shown).toBe(true);
    expect(ctx.state.dismissed).toBeNull();
  });

  test('Escape in the toolbar returns focus to the document and dismisses the toolbar for this selection', () => {
    const { context: ctx, calls } = context({ states: GRANTED, state: { shown: true } });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    const bold = view.getByRole('button', { name: 'Bold' });
    act(() => bold.focus());
    expect(fireEvent.keyDown(bold, { key: 'Escape' })).toBe(false);
    expect(calls.focused).toBe(1);
    expect(ctx.state.shown).toBe(false);
    expect(ctx.state.dismissed).toBe('["body",4,9]');
    expect(definitionOf().commands![0]!.getState!(ctx).enabled).toBe(false);
  });

  test('arrows rove between enabled buttons and Escape and Tab return focus to the document', () => {
    const states = { ...GRANTED, italic: REFUSED };
    const { context: ctx, calls } = context({ states });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    const buttons = [...view.getByRole('toolbar').querySelectorAll('button')];
    expect(buttons.map((button) => button.tabIndex)).toEqual([-1, -1, -1, -1, -1, -1]);
    act(() => buttons[0]!.focus());
    fireEvent.keyDown(buttons[0]!, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(buttons[2]!);
    fireEvent.keyDown(buttons[2]!, { key: 'End' });
    expect(document.activeElement).toBe(buttons[5]!);
    fireEvent.keyDown(buttons[5]!, { key: 'Home' });
    expect(document.activeElement).toBe(buttons[0]!);
    fireEvent.keyDown(buttons[0]!, { key: 'Escape' });
    fireEvent.keyDown(buttons[0]!, { key: 'Tab' });
    expect(calls.focused).toBe(2);
  });
});

describe('focus and pointer presses', () => {
  const next = (previous: Context, overrides: Partial<DocxPluginSnapshot> = {}) =>
    context({ states: GRANTED, state: previous.state, snapshot: overrides });
  const moved = { formatting: { hasSelection: true }, displayRange: { story: 'body', from: 10, to: 14, layoutId: 'layout-1' } } as unknown as DocxPluginSelection;

  test('is hidden while neither the document input nor the toolbar has focus', () => {
    const { context: ctx } = context({ states: GRANTED, snapshot: { inputFocused: false } });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    expect(view.queryByRole('toolbar')).toBeNull();
    expect(ctx.state.shown).toBe(false);
  });

  test('shows while the document input has focus', () => {
    const { context: ctx } = context({ states: GRANTED });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    expect(view.getByRole('toolbar').getAttribute('data-docx-mini-toolbar')).toBe('shown');
    expect(ctx.state.shown).toBe(true);
  });

  test('stays when focus moves from the document input into the toolbar', () => {
    const { context: ctx } = context({ states: GRANTED });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    const bold = view.getByRole('button', { name: 'Bold' });
    act(() => bold.focus());
    expect(ctx.state.toolbarFocused).toBe(true);
    view.update(next(ctx, { inputFocused: false }).context);
    expect(view.getByRole('toolbar').getAttribute('data-docx-mini-toolbar')).toBe('shown');
    expect(document.activeElement).toBe(bold);
  });

  test('stays while a toolbar button holds focus before the toolbar has recorded it', () => {
    const { context: ctx } = context({ states: GRANTED });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    const bold = view.getByRole('button', { name: 'Bold' });
    act(() => bold.focus());
    const early = context({ states: GRANTED, state: { shown: true, toolbarFocused: false }, snapshot: { inputFocused: false } });
    view.update(early.context);
    expect(view.queryByRole('toolbar')).not.toBeNull();
    expect(document.activeElement).toBe(bold);
    act(() => bold.blur());
    view.update(context({ states: GRANTED, state: { shown: true, toolbarFocused: false }, snapshot: { inputFocused: false } }).context);
    expect(view.queryByRole('toolbar')).toBeNull();
  });

  test('a press outside the toolbar dismisses it for this selection until a new selection settles', async () => {
    const plugin = docxMiniToolbar();
    const { context: ctx } = context({ states: GRANTED });
    const view = renderOverlay(plugin, ctx);
    expect(view.queryByRole('toolbar')).not.toBeNull();
    fireEvent.pointerDown(document.body);
    expect(ctx.state.dismissed).toBe('["body",4,9]');
    expect(ctx.state.shown).toBe(false);
    view.update(next(ctx).context);
    expect(view.queryByRole('toolbar')).toBeNull();

    const changing = next(ctx, { selection: moved, selectionSettled: false }).context;
    await definitionOf(plugin).onEvent!(changing, { type: 'selection-change', generation: 'document-1', version: '1', selection: moved });
    expect(changing.state.dismissed).toBeNull();
    view.update(changing);
    expect(view.queryByRole('toolbar')).toBeNull();
    view.update(next(changing, { selection: moved }).context);
    expect(view.queryByRole('toolbar')).not.toBeNull();
  });

  test('a press inside the toolbar keeps it', () => {
    const { context: ctx } = context({ states: GRANTED });
    const view = renderOverlay(docxMiniToolbar(), ctx);
    fireEvent.pointerDown(view.getByRole('button', { name: 'Italic' }));
    fireEvent.pointerDown(view.getByRole('toolbar'));
    expect(ctx.state.dismissed).toBeNull();
    expect(ctx.state.shown).toBe(true);
    view.update(next(ctx).context);
    expect(view.queryByRole('toolbar')).not.toBeNull();
  });

  test('listens for presses on its document with one capture listener, only while shown', () => {
    const spy = pressListeners(document);
    try {
      const hidden = context({ states: GRANTED, snapshot: { inputFocused: false } }).context;
      const view = renderOverlay(docxMiniToolbar(), hidden);
      expect(view.queryByRole('toolbar')).toBeNull();
      expect(spy.counts).toEqual({ added: 0, removed: 0 });

      const shown = next(hidden).context;
      view.update(shown);
      expect(view.queryByRole('toolbar')).not.toBeNull();
      expect(spy.counts).toEqual({ added: 1, removed: 0 });
      const still = next(shown, { mode: 'suggesting' }).context;
      view.update(still);
      expect(view.queryByRole('toolbar')).not.toBeNull();
      expect(spy.counts).toEqual({ added: 1, removed: 0 });

      view.update(next(still, { inputFocused: false }).context);
      expect(view.queryByRole('toolbar')).toBeNull();
      expect(spy.counts).toEqual({ added: 1, removed: 1 });

      const again = next(still).context;
      view.update(again);
      expect(spy.attached()).toBe(1);
      fireEvent.pointerDown(document.body);
      view.update(next(again).context);
      expect(view.queryByRole('toolbar')).toBeNull();
      expect(spy.counts).toEqual({ added: 2, removed: 2 });
    } finally {
      spy.restore();
    }
  });
});

describe('focus hand-back', () => {
  const DIALOG_REFUSAL: FocusAnswer = { ok: false, failure: { code: 'unsupported-policy', message: 'An editor dialog is open' } };
  const collapsed = { formatting: { hasSelection: false }, displayRange: { story: 'body', from: 9, to: 9, layoutId: 'layout-1' } } as unknown as DocxPluginSelection;
  const scrolled = { ...geometry(), viewport: { x: 0, y: 2000, width: 1000, height: 800 } } as DocxPluginGeometry;

  function documentInput() {
    const input = document.createElement('textarea');
    document.body.append(input);
    return input;
  }

  function focusedToolbar(focus?: () => FocusAnswer) {
    const plugin = docxMiniToolbar();
    const Overlay = definitionOf(plugin).overlay!;
    const first = context({ states: GRANTED, state: { shown: true }, focus });
    const view = render(<Overlay context={first.context} geometry={geometry()} />);
    const bold = view.getByRole('button', { name: 'Bold' });
    act(() => bold.focus());
    const after = (snapshotOverrides: Partial<DocxPluginSnapshot>, geo = geometry()) => {
      const next = context({ states: GRANTED, state: first.context.state, snapshot: { inputFocused: false, ...snapshotOverrides }, focus });
      view.rerender(<Overlay context={next.context} geometry={geo} />);
      return next.calls;
    };
    return { view, bold, calls: first.calls, after };
  }

  test('every button stays out of the tab order', () => {
    for (const states of [GRANTED, { ...GRANTED, bold: REFUSED }, {}]) {
      const { context: ctx } = context({ states });
      const view = renderOverlay(docxMiniToolbar(), ctx);
      const buttons = [...view.getByRole('toolbar').querySelectorAll('button')];
      expect(buttons.map((button) => button.tabIndex)).toEqual(buttons.map(() => -1));
      cleanup();
    }
  });

  for (const [cause, overrides, geo] of [
    ['the selection collapses', { selection: collapsed }, undefined],
    ['the mode changes to viewing', { mode: 'viewing' as const }, undefined],
    ['the selection scrolls out of view', {}, scrolled],
  ] as const) {
    test(`when ${cause} with a button focused, focus returns to the document before the toolbar goes`, () => {
      const input = documentInput();
      try {
        const { view, bold, after } = focusedToolbar(() => {
          input.focus();
          return { ok: true };
        });
        const calls = after(overrides, geo);
        expect(calls.focused).toBe(1);
        expect(calls.focusedFrom[0]).toBe(bold);
        expect(document.activeElement).toBe(input);
        expect(view.queryByRole('toolbar')).toBeNull();
      } finally {
        input.remove();
      }
    });
  }

  test('unmounting with a button focused returns focus to the document first', () => {
    const { view, bold, calls } = focusedToolbar();
    view.unmount();
    expect(calls.focused).toBe(1);
    expect(calls.focusedFrom[0]).toBe(bold);
  });

  test('a refused focusDocument leaves Tab and Escape to the browser and keeps the toolbar', () => {
    const { view, bold, calls } = focusedToolbar(() => DIALOG_REFUSAL);
    expect(fireEvent.keyDown(bold, { key: 'Tab' })).toBe(true);
    expect(fireEvent.keyDown(bold, { key: 'Escape' })).toBe(true);
    expect(calls.focused).toBe(2);
    expect(view.getByRole('toolbar')).not.toBeNull();
    expect(document.activeElement).toBe(bold);
    expect(calls.states.some((state) => state.dismissed !== null)).toBe(false);
  });

  test('in the in-thread editor, which refuses focusDocument, the toolbar stays under its focus when it would hide', () => {
    const { view, bold, after } = focusedToolbar(() => NO_INPUT_REFUSAL);
    const calls = after({ selection: collapsed });
    expect(calls.focused).toBe(1);
    expect(view.getByRole('toolbar')).not.toBeNull();
    expect(document.activeElement).toBe(bold);
    expect(fireEvent.keyDown(bold, { key: 'Tab' })).toBe(true);
  });
});
