import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { useRef } from 'react';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import { testBinding } from '../commands/testing';
import { DocxCommandProvider } from '../commands/DocxCommandProvider';
import { LocaleProvider } from '../i18n';
import {
  computeSelectionToolbarPosition,
  SelectionMiniToolbar,
  type SelectionMiniToolbarButton,
} from './SelectionMiniToolbar';
import { SelectionMiniToolbarHost } from './SelectionMiniToolbarHost';
import {
  resolveSelectionToolbarItems,
  useSelectionMiniToolbar,
} from './DocxEditor/hooks/useSelectionMiniToolbar';
import type { DocxSelectionToolbarItem } from './DocxEditor';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const anchor = { left: 100, top: 200, right: 180, bottom: 220, width: 80, height: 20 };
const viewport = { width: 1024, height: 768 };
const size = { width: 120, height: 36 };

function settle(ms = 40): Promise<void> {
  return act(async () => {
    await new Promise((resolve) => setTimeout(resolve, ms));
  });
}

test('positions above the selection by default', () => {
  expect(computeSelectionToolbarPosition(anchor, size, viewport)).toEqual({
    left: 80,
    top: 156,
    flipped: false,
  });
});

test('flips below the selection near the top edge', () => {
  const position = computeSelectionToolbarPosition({ ...anchor, top: 20, bottom: 40 }, size, viewport);
  expect(position.flipped).toBe(true);
  expect(position.top).toBe(48);
});

test('clamps horizontally inside the viewport margin', () => {
  expect(computeSelectionToolbarPosition({ ...anchor, left: 0, right: 20, width: 20 }, size, viewport).left).toBe(8);
  expect(
    computeSelectionToolbarPosition({ ...anchor, left: 1000, right: 1020, width: 20 }, size, viewport).left
  ).toBe(1024 - 120 - 8);
});

test('resolveSelectionToolbarItems defaults and hides when false', () => {
  expect(resolveSelectionToolbarItems(true)).toEqual(['clearFormatting', 'pastePlainText']);
  expect(resolveSelectionToolbarItems(undefined)).toEqual(['clearFormatting', 'pastePlainText']);
  expect(resolveSelectionToolbarItems(false)).toBeNull();
  expect(resolveSelectionToolbarItems({ items: ['bold'] })).toEqual(['bold']);
  expect(resolveSelectionToolbarItems({})).toEqual(['clearFormatting', 'pastePlainText']);
});

function stubButtons(): SelectionMiniToolbarButton[] {
  return [
    { id: 'clearFormatting', label: 'Clear formatting', icon: 'format_clear', onActivate: () => {} },
    { id: 'pastePlainText', label: 'Paste as Plain Text', shortcut: 'Ctrl+Shift+V', icon: 'content_paste', onActivate: () => {} },
  ];
}

test('renders a labelled toolbar with labelled buttons in order', () => {
  const view = render(
    <SelectionMiniToolbar anchorRect={anchor} buttons={stubButtons()} label="Selection toolbar" onDismiss={() => {}} />
  );
  const toolbar = view.getByRole('toolbar', { name: 'Selection toolbar' });
  const buttons = within(toolbar).getAllByRole('button');
  expect(buttons.map((button) => button.getAttribute('aria-label'))).toEqual([
    'Clear formatting',
    'Paste as Plain Text',
  ]);
});

test('coarse pointers get at least 44px targets', () => {
  const view = render(
    <SelectionMiniToolbar
      anchorRect={anchor}
      buttons={stubButtons()}
      label="Selection toolbar"
      coarsePointer
      onDismiss={() => {}}
    />
  );
  for (const button of view.getAllByRole('button')) {
    expect(button.className).toContain('min-w-[44px]');
    expect(button.className).toContain('min-h-[44px]');
  }
});

test('flips below when the anchor is at the top edge', () => {
  const view = render(
    <SelectionMiniToolbar
      anchorRect={{ ...anchor, top: 2, bottom: 22 }}
      buttons={stubButtons()}
      label="Selection toolbar"
      onDismiss={() => {}}
    />
  );
  expect(view.getByTestId('selection-mini-toolbar').getAttribute('data-flipped')).toBe('true');
});

test('roving focus moves with arrows, Home and End, and Escape dismisses to the editor', () => {
  let dismissed: boolean | null = null;
  const view = render(
    <SelectionMiniToolbar
      anchorRect={anchor}
      buttons={stubButtons()}
      label="Selection toolbar"
      onDismiss={(returnFocus) => {
        dismissed = returnFocus;
      }}
    />
  );
  const toolbar = view.getByRole('toolbar');
  const buttons = view.getAllByRole('button');
  expect(buttons[0]!.tabIndex).toBe(0);
  expect(buttons[1]!.tabIndex).toBe(-1);
  buttons[0]!.focus();
  fireEvent.keyDown(toolbar, { key: 'ArrowRight' });
  expect(document.activeElement).toBe(buttons[1]);
  fireEvent.keyDown(toolbar, { key: 'ArrowRight' });
  expect(document.activeElement).toBe(buttons[0]);
  fireEvent.keyDown(toolbar, { key: 'End' });
  expect(document.activeElement).toBe(buttons[1]);
  fireEvent.keyDown(toolbar, { key: 'Home' });
  expect(document.activeElement).toBe(buttons[0]);
  fireEvent.keyDown(toolbar, { key: 'Escape' });
  expect(dismissed as boolean | null).toBe(true);
});

function HookHarness({
  selectionKey,
  enabled = true,
  documentKey,
  withToolbar = false,
  portaled = false,
  toolbarOutside = false,
}: {
  selectionKey: string | null;
  enabled?: boolean;
  documentKey?: unknown;
  withToolbar?: boolean;
  portaled?: boolean;
  toolbarOutside?: boolean;
}) {
  const rootRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const mini = useSelectionMiniToolbar({
    enabled,
    rootRef,
    scrollRef,
    selectionKey,
    getAnchorRect: () => anchor,
    documentKey,
    isEditorTarget: portaled ? (target) => target instanceof Element && target.matches('[data-testid="portaled-input"]') : undefined,
    settleMs: 10,
  });
  return (
    <>
    <div ref={rootRef}>
      <div ref={scrollRef} data-testid="scroller">
        <textarea data-testid="editor" />
      </div>
      {!toolbarOutside && mini.visible && mini.anchorRect ? (
        withToolbar ? (
          <SelectionMiniToolbar
            ref={mini.toolbarRef}
            anchorRect={mini.anchorRect}
            buttons={stubButtons()}
            label="Selection toolbar"
            onDismiss={mini.dismiss}
          />
        ) : (
          <div data-testid="mini">shown</div>
        )
      ) : null}
    </div>
    <button type="button" data-testid="outside">
      Outside
    </button>
    <textarea data-testid="portaled-input" />
    {toolbarOutside && mini.visible && mini.anchorRect ? (
      <SelectionMiniToolbar
        ref={mini.toolbarRef}
        anchorRect={mini.anchorRect}
        buttons={stubButtons()}
        label="Selection toolbar"
        onDismiss={mini.dismiss}
      />
    ) : null}
    </>
  );
}

test('shows only after a settled selection, never while the pointer is down', async () => {
  const view = render(<HookHarness selectionKey="1:5" />);
  const editor = view.getByTestId('editor');
  fireEvent.pointerDown(editor);
  await settle();
  expect(view.queryByTestId('mini')).toBeNull();
  fireEvent.pointerUp(editor);
  await settle();
  expect(view.queryByTestId('mini')).not.toBeNull();
});

test('does not show between the clicks of a double click', async () => {
  const view = render(<HookHarness selectionKey="1:5" />);
  const editor = view.getByTestId('editor');
  fireEvent.pointerDown(editor);
  fireEvent.pointerUp(editor);
  fireEvent.pointerDown(editor);
  await settle();
  expect(view.queryByTestId('mini')).toBeNull();
  fireEvent.pointerUp(editor);
  await settle();
  expect(view.queryByTestId('mini')).not.toBeNull();
});

test('shows after a keyboard selection stops changing', async () => {
  const view = render(<HookHarness selectionKey="1:2" />);
  view.rerender(<HookHarness selectionKey="1:3" />);
  view.rerender(<HookHarness selectionKey="1:4" />);
  await settle();
  expect(view.queryByTestId('mini')).not.toBeNull();
});

test('hides on typing, scroll, Escape, collapse and document change', async () => {
  const view = render(<HookHarness selectionKey="1:5" documentKey="v1" />);
  const editor = view.getByTestId('editor');
  await settle();
  expect(view.queryByTestId('mini')).not.toBeNull();
  fireEvent.keyDown(editor, { key: 'x' });
  expect(view.queryByTestId('mini')).toBeNull();

  view.rerender(<HookHarness selectionKey="1:6" documentKey="v1" />);
  await settle();
  expect(view.queryByTestId('mini')).not.toBeNull();
  fireEvent.scroll(view.getByTestId('scroller'));
  expect(view.queryByTestId('mini')).toBeNull();

  view.rerender(<HookHarness selectionKey="1:7" documentKey="v1" />);
  await settle();
  expect(view.queryByTestId('mini')).not.toBeNull();
  fireEvent.keyDown(editor, { key: 'Escape' });
  expect(view.queryByTestId('mini')).toBeNull();

  view.rerender(<HookHarness selectionKey="1:8" documentKey="v1" />);
  await settle();
  expect(view.queryByTestId('mini')).not.toBeNull();
  view.rerender(<HookHarness selectionKey="1:8" documentKey="v2" />);
  expect(view.queryByTestId('mini')).toBeNull();

  view.rerender(<HookHarness selectionKey={null} documentKey="v2" />);
  expect(view.queryByTestId('mini')).toBeNull();
});

test('typing in an editor input outside the root hides the toolbar when the editor claims it', async () => {
  const view = render(<HookHarness selectionKey="1:5" />);
  await settle();
  fireEvent.keyDown(view.getByTestId('portaled-input'), { key: 'x' });
  expect(view.queryByTestId('mini')).not.toBeNull();
  view.unmount();

  const claimed = render(<HookHarness selectionKey="1:5" portaled />);
  await settle();
  expect(claimed.queryByTestId('mini')).not.toBeNull();
  fireEvent.keyDown(claimed.getByTestId('portaled-input'), { key: 'x' });
  expect(claimed.queryByTestId('mini')).toBeNull();
  claimed.rerender(<HookHarness selectionKey="1:6" portaled />);
  await settle();
  fireEvent.keyDown(claimed.getByTestId('portaled-input'), { key: 'Escape' });
  expect(claimed.queryByTestId('mini')).toBeNull();
});

test('stays when focus moves into the toolbar and hides when focus leaves the editor', async () => {
  const view = render(<HookHarness selectionKey="1:5" withToolbar />);
  const editor = view.getByTestId('editor');
  await settle();
  const toolbarButton = view.getAllByRole('button').find((button) => button.getAttribute('aria-label') === 'Clear formatting')!;
  fireEvent.focusOut(editor, { relatedTarget: toolbarButton });
  expect(view.queryByRole('toolbar')).not.toBeNull();
  fireEvent.focusOut(editor, { relatedTarget: view.getByTestId('outside') });
  expect(view.queryByRole('toolbar')).toBeNull();
});

test('focus moving from an editor input outside the root into a toolbar outside the root keeps it', async () => {
  const view = render(<HookHarness selectionKey="1:5" portaled toolbarOutside />);
  const input = view.getByTestId('portaled-input') as HTMLTextAreaElement;
  await settle();
  input.focus();
  fireEvent.keyDown(input, { key: 'F10', altKey: true });
  const focused = document.activeElement as HTMLElement;
  expect(focused.getAttribute('aria-label')).toBe('Clear formatting');
  fireEvent.focusOut(input, { relatedTarget: focused });
  expect(view.queryByRole('toolbar')).not.toBeNull();
  fireEvent.focusOut(input, { relatedTarget: view.getByTestId('outside') });
  expect(view.queryByRole('toolbar')).toBeNull();
});

test('tabbing out of a toolbar outside the root to elsewhere hides it', async () => {
  const view = render(<HookHarness selectionKey="1:5" portaled toolbarOutside />);
  const input = view.getByTestId('portaled-input') as HTMLTextAreaElement;
  await settle();
  input.focus();
  fireEvent.keyDown(input, { key: 'F10', altKey: true });
  const focused = document.activeElement as HTMLElement;
  fireEvent.focusOut(input, { relatedTarget: focused });
  expect(view.queryByRole('toolbar')).not.toBeNull();
  fireEvent.focusOut(focused, { relatedTarget: view.getByTestId('outside') });
  expect(view.queryByRole('toolbar')).toBeNull();
});

test('Alt+F10 moves focus into the shown toolbar', async () => {
  const view = render(<HookHarness selectionKey="1:5" withToolbar />);
  const editor = view.getByTestId('editor') as HTMLTextAreaElement;
  await settle();
  editor.focus();
  fireEvent.keyDown(editor, { key: 'F10', altKey: true });
  expect(document.activeElement?.getAttribute('aria-label')).toBe('Clear formatting');
});

function HostHarness({
  items,
  readOnly = false,
  pasteEnabled = true,
  onPastePlainText = () => {},
}: {
  items: readonly DocxSelectionToolbarItem[];
  readOnly?: boolean;
  pasteEnabled?: boolean;
  onPastePlainText?: () => void;
}) {
  const harness = testBinding(readOnly ? { readOnly: true } : {});
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  return (
    <LocaleProvider>
      <DocxCommandProvider commands={controller.store}>
        <SelectionMiniToolbarHost
          anchorRect={anchor}
          items={items}
          pasteEnabled={pasteEnabled}
          onPastePlainText={onPastePlainText}
          onDismiss={() => {}}
        />
      </DocxCommandProvider>
    </LocaleProvider>
  );
}

test('custom items render in order with the main toolbar labels', () => {
  const view = render(<HostHarness items={['bold', 'clearFormatting', 'pastePlainText']} />);
  const toolbar = view.getByRole('toolbar', { name: 'Selection toolbar' });
  expect(within(toolbar).getAllByRole('button').map((button) => button.getAttribute('aria-label'))).toEqual([
    'Bold',
    'Clear formatting',
    'Paste as Plain Text',
  ]);
});

test('a disabled command button is not shown', () => {
  const view = render(<HostHarness items={['clearFormatting', 'bold']} readOnly pasteEnabled={false} />);
  expect(view.queryByRole('toolbar')).toBeNull();
});

test('clicking a command runs it without moving focus from the editor surface', async () => {
  const harness = testBinding({});
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  const view = render(
    <LocaleProvider>
      <DocxCommandProvider commands={controller.store}>
        <SelectionMiniToolbarHost
          anchorRect={anchor}
          items={['bold']}
          pasteEnabled={false}
          onPastePlainText={() => {}}
          onDismiss={() => {}}
        />
      </DocxCommandProvider>
    </LocaleProvider>
  );
  fireEvent.click(view.getByRole('button', { name: 'Bold' }));
  await settle(10);
  expect(harness.calls.map((call) => call.id)).toEqual(['bold']);
});

test('Paste as plain text returns focus to the editor and pastes', async () => {
  const harness = testBinding({});
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  let pasted = 0;
  const view = render(
    <LocaleProvider>
      <DocxCommandProvider commands={controller.store}>
        <SelectionMiniToolbarHost
          anchorRect={anchor}
          items={['pastePlainText']}
          pasteEnabled
          onPastePlainText={() => { pasted += 1; }}
          onDismiss={() => {}}
        />
      </DocxCommandProvider>
    </LocaleProvider>
  );
  fireEvent.click(view.getByRole('button', { name: 'Paste as Plain Text' }));
  await settle(30);
  expect(pasted).toBe(1);
  expect(harness.state.focused).toBeGreaterThan(0);
});
