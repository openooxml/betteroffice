import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { useRef } from 'react';
import { createT, en } from '@betteroffice/xlsx-i18n';
import type { ReactNode } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  EditorToolbar,
  Toolbar,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  XlsxCommandProvider,
  XlsxToolbar,
  useEditorToolbar,
  type FormattingAction,
  type ToolbarProps,
} from '../../index';
import {
  createXlsxCommandController,
  type XlsxCommandController,
} from '../../commands/createXlsxCommandStore';
import { isMacPlatform } from '../../commands/descriptors';
import { PLAIN_FORMATTING, testBinding, testEnvironment } from '../../commands/testing';
import { useCommandShortcuts } from '../../commands/useCommandShortcuts';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

function screen() {
  return within(document.body);
}

const originalRect = HTMLElement.prototype.getBoundingClientRect;

beforeAll(() => {
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    const width = this.getAttribute('role') === 'toolbar' ? 4000 : 0;
    return { width, height: 28, top: 0, left: 0, right: width, bottom: 28, x: 0, y: 0, toJSON() {} } as DOMRect;
  };
});

afterEach(cleanup);
afterAll(async () => {
  HTMLElement.prototype.getBoundingClientRect = originalRect;
  if (ownsDom) await GlobalRegistrator.unregister();
});

const MOD = isMacPlatform() ? { metaKey: true } : { ctrlKey: true };

function editor(overrides = {}) {
  const harness = testBinding({ translate: createT(en), ...overrides });
  const controller = createXlsxCommandController();
  controller.attach(harness.binding);
  return { harness, controller };
}

function reasonOf(element: Element): string | null {
  const id = element.getAttribute('aria-describedby');
  return id ? (document.getElementById(id)?.textContent ?? null) : null;
}

function precedes(a: Element, b: Element): boolean {
  return (a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0;
}

function CompactToolbar({ onHostAction }: { onHostAction(): void }) {
  return (
    <EditorToolbar mode="commands">
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Formatting">
          <ToolbarCommandSelect id="numberFormat" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <ToolbarButton title="Host action" onClick={onHostAction}>
          Host action
        </ToolbarButton>
      </EditorToolbar.Toolbar>
    </EditorToolbar>
  );
}

function Shortcuts({
  controller,
  children,
}: {
  controller: XlsxCommandController;
  children: ReactNode;
}) {
  const root = useRef<HTMLDivElement>(null);
  useCommandShortcuts(controller, root);
  return <div ref={root}>{children}</div>;
}

describe('host-composed xlsx toolbar', () => {
  test('renders built-in controls in the host order next to a host action', async () => {
    const { harness, controller } = editor();
    let hostActions = 0;
    render(
      <XlsxCommandProvider commands={controller.store}>
        <CompactToolbar onHostAction={() => (hostActions += 1)} />
      </XlsxCommandProvider>
    );
    const toolbar = screen().getByRole('toolbar');
    const formats = within(toolbar).getByRole('button', { name: 'More number formats' });
    const bold = within(toolbar).getByRole('button', { name: 'Bold' });
    const undo = within(toolbar).getByRole('button', { name: 'Undo' });
    const host = within(toolbar).getByRole('button', { name: 'Host action' });
    expect(precedes(formats, bold)).toBe(true);
    expect(precedes(bold, undo)).toBe(true);
    expect(precedes(undo, host)).toBe(true);
    expect(within(toolbar).queryByRole('button', { name: 'Italic' })).toBeNull();

    fireEvent.click(bold);
    fireEvent.click(host);
    fireEvent.click(formats);
    fireEvent.click(screen().getByRole('menuitemradio', { name: 'Percent' }));
    await act(async () => {});
    expect(hostActions).toBe(1);
    expect(harness.calls).toEqual([
      { id: 'bold', args: null, ordered: true },
      { id: 'numberFormat', args: { value: 'percent' }, ordered: true },
    ]);
  });

  test('reflects mixed marks and states why a control is disabled', async () => {
    const { harness, controller } = editor({ canUndo: false });
    render(
      <XlsxCommandProvider commands={controller.store}>
        <CompactToolbar onHostAction={() => {}} />
      </XlsxCommandProvider>
    );
    const bold = screen().getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    const undo = screen().getByRole('button', { name: 'Undo' });
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(reasonOf(undo)).toBe('Nothing to undo.');
    fireEvent.click(undo);

    const selection = testEnvironment().selection as Exclude<
      ReturnType<typeof testEnvironment>['selection'],
      'chart' | null
    >;
    await act(async () => {
      harness.update({ selection: { ...selection, formatting: { ...PLAIN_FORMATTING, bold: undefined } } });
      controller.refresh();
    });
    expect(bold.getAttribute('aria-pressed')).toBe('mixed');
    expect(harness.calls).toEqual([]);
  });

  test('stays unavailable until the host passes an editor', () => {
    render(
      <XlsxCommandProvider commands={null}>
        <CompactToolbar onHostAction={() => {}} />
      </XlsxCommandProvider>
    );
    const bold = screen().getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-disabled')).toBe('true');
    expect(reasonOf(bold)).toBe('The editor is not ready.');
  });

  test('renders the default rail without children and follows the editor locale outside it', async () => {
    const { harness, controller } = editor();
    harness.state.i18n = { _lang: 'de', toolbar: { bold: 'Fett' } };
    render(
      <XlsxCommandProvider commands={controller.store}>
        <EditorToolbar mode="commands" />
      </XlsxCommandProvider>
    );
    await act(async () => controller.refresh());
    const toolbar = screen().getByRole('toolbar');
    within(toolbar).getByRole('button', { name: 'Fett' });
    within(toolbar).getByRole('button', { name: 'Italic' });
    within(toolbar).getByRole('group', { name: 'Borders' });
  });

  test('projects command state onto the legacy context and routes its callbacks', async () => {
    const { harness, controller } = editor();
    let seen: ToolbarProps | null = null;
    function LegacyHost() {
      seen = useEditorToolbar();
      return null;
    }
    const selection = testEnvironment().selection as Exclude<
      ReturnType<typeof testEnvironment>['selection'],
      'chart' | null
    >;
    harness.update({
      selection: { ...selection, formatting: { ...PLAIN_FORMATTING, bold: undefined, italic: true } },
    });
    render(
      <XlsxCommandProvider commands={controller.store}>
        <EditorToolbar mode="commands">
          <LegacyHost />
        </EditorToolbar>
      </XlsxCommandProvider>
    );
    const props = seen as unknown as ToolbarProps;
    expect(props.currentFormatting?.bold).toBeUndefined();
    expect(props.currentFormatting?.italic).toBe(true);
    expect(props.selectionShape).toEqual({ rows: 3, columns: 2, canUnmerge: true });
    expect(props.canUndo).toBe(true);
    props.onFormat?.('currency');
    props.onFormat?.({ type: 'fontSize', value: 14 });
    props.onMerge?.('vertical');
    props.onZoomChange?.(1.5);
    await act(async () => {});
    expect(harness.calls.map(({ id, args }) => [id, args])).toEqual([
      ['zoom', { scale: 1.5 }],
      ['numberFormat', { value: 'currency' }],
      ['fontSize', { points: 14 }],
      ['merge', { value: 'vertical' }],
    ]);
  });

  test('rejects callbacks and state overrides in command mode', () => {
    const { controller } = editor();
    const originalError = console.error;
    console.error = () => {};
    try {
      expect(() =>
        render(
          <XlsxCommandProvider commands={controller.store}>
            <EditorToolbar mode="commands" onUndo={() => {}} />
          </XlsxCommandProvider>
        )
      ).toThrow(/remove onUndo/);
      expect(() =>
        render(
          <XlsxCommandProvider commands={controller.store}>
            <EditorToolbar mode="commands">
              <EditorToolbar.Toolbar canUndo />
            </EditorToolbar>
          </XlsxCommandProvider>
        )
      ).toThrow(/remove canUndo/);
    } finally {
      console.error = originalError;
    }
  });
});

describe('legacy prop-configured toolbar', () => {
  test('keeps its props, callbacks, context and appended children', () => {
    const actions: FormattingAction[] = [];
    const merges: string[] = [];
    let undos = 0;
    function Probe() {
      return <span data-testid="probe">{String(useEditorToolbar().canUndo)}</span>;
    }
    render(
      <EditorToolbar
        currentFormatting={{ bold: true, fontFamily: 'Georgia' }}
        selectionShape={{ rows: 1, columns: 3 }}
        onFormat={(action) => actions.push(action)}
        onMerge={(action) => merges.push(action)}
        onUndo={() => (undos += 1)}
        canUndo
      >
        <EditorToolbar.Toolbar>
          <button type="button">Appended</button>
        </EditorToolbar.Toolbar>
        <Probe />
      </EditorToolbar>
    );
    const toolbar = screen().getByRole('toolbar');
    const bold = within(toolbar).getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    expect(within(toolbar).getByRole('button', { name: 'Italic' }).getAttribute('aria-pressed')).toBe(
      'false'
    );
    expect(within(toolbar).getByRole('button', { name: 'Font family' }).textContent).toContain(
      'Georgia'
    );
    expect(screen().getByTestId('probe').textContent).toBe('true');
    const appended = within(toolbar).getByRole('button', { name: 'Appended' });
    expect(precedes(within(toolbar).getByRole('group', { name: 'Alignment' }), appended)).toBe(true);

    fireEvent.click(bold);
    fireEvent.click(within(toolbar).getByRole('button', { name: 'Format as currency' }));
    fireEvent.click(within(toolbar).getByRole('button', { name: 'Format as percent' }));
    fireEvent.click(within(toolbar).getByRole('button', { name: 'More number formats' }));
    fireEvent.click(screen().getByRole('menuitemradio', { name: 'Currency' }));
    fireEvent.click(within(toolbar).getByRole('button', { name: 'Increase decimal places' }));
    fireEvent.click(within(toolbar).getByRole('button', { name: 'Increase font size' }));
    fireEvent.click(within(toolbar).getByRole('button', { name: 'Undo' }));
    fireEvent.click(within(toolbar).getByTestId('xlsx-merge-all'));
    const redo = within(toolbar).getByRole('button', { name: 'Redo' });
    expect(redo.getAttribute('aria-disabled')).toBe('true');
    expect(reasonOf(redo)).toBe('Not available in this toolbar.');
    return act(async () => {}).then(() => {
      expect(actions).toEqual([
        'bold',
        'currency',
        'percent',
        { type: 'numberFormat', value: 'currency' },
        'increaseDecimal',
        { type: 'fontSize', value: 11 },
      ]);
      expect(undos).toBe(1);
      expect(merges).toEqual(['all']);
    });
  });

  test('works standalone, including the XlsxToolbar alias and explicit overrides', async () => {
    const actions: FormattingAction[] = [];
    render(
      <>
        <Toolbar onFormat={(action) => actions.push(action)} disabled />
        <XlsxToolbar onFormat={(action) => actions.push(action)} className="host" />
      </>
    );
    const [disabled, enabled] = screen().getAllByRole('toolbar');
    expect(enabled.className).toBe('host');
    const off = within(disabled).getByRole('button', { name: 'Italic' });
    expect(off.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(off);
    fireEvent.click(within(enabled).getByRole('button', { name: 'Italic' }));
    await act(async () => {});
    expect(actions).toEqual(['italic']);
  });
});

describe('xlsx font size control', () => {
  test('keeps keyboard focus on commit and returns to the grid only after a pointer choice', async () => {
    const { harness, controller } = editor();
    render(
      <XlsxCommandProvider commands={controller.store}>
        <EditorToolbar mode="commands">
          <EditorToolbar.Toolbar>
            <ToolbarCommandSelect id="fontSize" />
            <ToolbarCommandButton id="bold" />
          </EditorToolbar.Toolbar>
        </EditorToolbar>
      </XlsxCommandProvider>
    );
    const size = screen().getByRole('combobox', { name: 'Font size' }) as HTMLInputElement;
    act(() => size.focus());
    fireEvent.change(size, { target: { value: '14' } });
    fireEvent.keyDown(size, { key: 'Enter' });
    expect(document.activeElement).toBe(size);

    fireEvent.change(size, { target: { value: '15' } });
    const bold = screen().getByRole('button', { name: 'Bold' });
    act(() => bold.focus());
    await act(async () => {
      await new Promise((resolve) => requestAnimationFrame(() => resolve(null)));
      await new Promise((resolve) => requestAnimationFrame(() => resolve(null)));
    });
    expect(document.activeElement).toBe(bold);
    expect(harness.state.focused).toBe(0);

    act(() => size.focus());
    fireEvent.click(screen().getByRole('option', { name: '18' }));
    await act(async () => {
      await new Promise((resolve) => requestAnimationFrame(() => resolve(null)));
    });
    expect(harness.state.focused).toBe(1);
    expect(harness.calls.map(({ args }) => args)).toEqual([
      { points: 14 },
      { points: 15 },
      { points: 18 },
    ]);
  });
});

describe('xlsx toolbar dropdown semantics', () => {
  test('menus take arrow keys and Escape returns to the button', () => {
    const picked: string[] = [];
    render(
      <ToolbarDropdown title="Formats" trigger="123">
        {(close) => (
          <>
            <ToolbarMenuItem label="Number" selected onClick={() => picked.push('number')} close={close} />
            <ToolbarMenuItem label="Percent" selected={false} onClick={() => picked.push('percent')} close={close} />
            <ToolbarMenuItem label="Clear" disabled description="Nothing to clear" close={close} />
          </>
        )}
      </ToolbarDropdown>
    );
    const trigger = screen().getByRole('button', { name: 'Formats' });
    act(() => trigger.focus());
    fireEvent.keyDown(trigger, { key: 'ArrowDown' });
    const menu = screen().getByRole('menu', { name: 'Formats' });
    const items = within(menu).getAllByRole('menuitemradio');
    expect(document.activeElement).toBe(items[0]);
    expect(items[0].getAttribute('aria-checked')).toBe('true');
    fireEvent.keyDown(menu, { key: 'ArrowDown' });
    expect(document.activeElement).toBe(items[1]);
    fireEvent.keyDown(menu, { key: 'End' });
    const clear = within(menu).getByRole('menuitem', { name: 'Clear' });
    expect(document.activeElement).toBe(clear);
    expect(reasonOf(clear)).toBe('Nothing to clear');
    fireEvent.keyDown(menu, { key: 'Escape' });
    expect(screen().queryByRole('menu')).toBeNull();
    expect(document.activeElement).toBe(trigger);

    fireEvent.keyDown(trigger, { key: 'Enter' });
    fireEvent.click(within(screen().getByRole('menu')).getByRole('menuitemradio', { name: 'Percent' }));
    expect(picked).toEqual(['percent']);
    expect(screen().queryByRole('menu')).toBeNull();
    expect(document.activeElement).toBe(trigger);
  });

  test('opens above a trigger near the bottom edge and stays on screen', () => {
    const height = window.innerHeight;
    Object.defineProperty(window, 'innerHeight', { value: 300, configurable: true });
    const rect = HTMLElement.prototype.getBoundingClientRect;
    const scroll = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'scrollHeight');
    HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
      if (this.getAttribute('aria-label') === 'Formats' && this.tagName === 'BUTTON') {
        return { width: 40, height: 28, top: 260, left: 10, right: 50, bottom: 288, x: 10, y: 260, toJSON() {} } as DOMRect;
      }
      return rect.call(this);
    };
    Object.defineProperty(HTMLElement.prototype, 'scrollHeight', {
      configurable: true,
      get(this: HTMLElement) {
        return this.getAttribute('role') === 'menu' ? 400 : 0;
      },
    });
    try {
      render(
        <ToolbarDropdown title="Formats" trigger="123">
          {(close) => <ToolbarMenuItem label="Number" close={close} />}
        </ToolbarDropdown>
      );
      fireEvent.click(screen().getByRole('button', { name: 'Formats' }));
      const menu = screen().getByRole('menu', { name: 'Formats' });
      const top = parseFloat(menu.style.top);
      const maxHeight = parseFloat(menu.style.maxHeight);
      expect(top).toBeGreaterThanOrEqual(8);
      expect(top + maxHeight).toBeLessThanOrEqual(256);
      expect(maxHeight).toBe(248);
    } finally {
      HTMLElement.prototype.getBoundingClientRect = rect;
      if (scroll) Object.defineProperty(HTMLElement.prototype, 'scrollHeight', scroll);
      else Reflect.deleteProperty(HTMLElement.prototype, 'scrollHeight');
      Object.defineProperty(window, 'innerHeight', { value: height, configurable: true });
    }
  });

  test('arbitrary content opens as a dialog, not a menu', () => {
    render(
      <ToolbarDropdown title="Tools" trigger="Tools">
        {() => (
          <label>
            Width <input aria-label="Width" />
          </label>
        )}
      </ToolbarDropdown>
    );
    const trigger = screen().getByRole('button', { name: 'Tools' });
    act(() => trigger.focus());
    fireEvent.keyDown(trigger, { key: 'ArrowDown' });
    const dialog = screen().getByRole('dialog', { name: 'Tools' });
    expect(screen().queryByRole('menu')).toBeNull();
    expect(document.activeElement).toBe(within(dialog).getByLabelText('Width'));
    expect(trigger.getAttribute('aria-haspopup')).toBe('dialog');
    fireEvent.keyDown(dialog, { key: 'Escape' });
    expect(document.activeElement).toBe(trigger);
  });
});

describe('xlsx command shortcuts', () => {
  test('reach only the editor that owns the event, and fields keep their keys', async () => {
    const first = editor();
    const second = editor();
    render(
      <>
        <Shortcuts controller={first.controller}>
          <button type="button">first</button>
          <input aria-label="first field" />
        </Shortcuts>
        <Shortcuts controller={second.controller}>
          <button type="button">second</button>
        </Shortcuts>
      </>
    );
    fireEvent.keyDown(screen().getByRole('button', { name: 'first' }), { key: 'b', ...MOD });
    fireEvent.keyDown(screen().getByRole('button', { name: 'second' }), {
      key: 'z',
      shiftKey: true,
      ...MOD,
    });
    const field = screen().getByLabelText('first field');
    fireEvent.keyDown(field, { key: 'z', ...MOD });
    fireEvent.keyDown(field, { key: 's', ...MOD });
    fireEvent.keyDown(screen().getByRole('button', { name: 'first' }), {
      key: 'b',
      repeat: true,
      ...MOD,
    });
    await act(async () => {});
    expect(first.harness.calls.map((call) => call.id)).toEqual(['bold', 'save']);
    expect(second.harness.calls.map((call) => call.id)).toEqual(['redo']);
  });

  test('reach the editor from an external toolbar registered as its chrome', async () => {
    const { harness, controller } = editor();
    render(
      <>
        <Shortcuts controller={controller}>
          <span />
        </Shortcuts>
        <XlsxCommandProvider commands={controller.store}>
          <CompactToolbar onHostAction={() => {}} />
        </XlsxCommandProvider>
      </>
    );
    fireEvent.keyDown(screen().getByRole('button', { name: 'Host action' }), { key: 'i', ...MOD });
    await act(async () => {});
    expect(harness.calls.map((call) => call.id)).toEqual(['italic']);
  });
});
