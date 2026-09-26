import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createDocxCommandController } from '../../commands/createDocxCommandStore';
import { isMacPlatform } from '../../commands/descriptors';
import { PLAIN_CONTEXT, testBinding } from '../../commands/testing';
import { useKeyboardShortcuts } from '../DocxEditor/hooks/useKeyboardShortcuts';
import type { DocxCommandController } from '../../commands/createDocxCommandStore';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');
// Loaded after the DOM exists: Radix picks its layout-effect hook at import time.
const {
  DocxCommandProvider,
  EditorToolbar,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarGroup,
  useDocxCommandState,
} = await import('../../index');

/** Queries the current document; `screen` stays bound to the first registered one. */
function screen() {
  return within(document.body);
}

const MOD = isMacPlatform() ? { metaKey: true } : { ctrlKey: true };

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function editor(overrides = {}) {
  const harness = testBinding(overrides);
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  return { harness, controller };
}

function CompactToolbar({ onHostAction }: { onHostAction(): void }) {
  return (
    <EditorToolbar>
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Formatting">
          <ToolbarCommandSelect id="paragraphStyle" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <EditorToolbar.Review />
        <ToolbarButton title="Host action" onClick={onHostAction}>
          Host action
        </ToolbarButton>
      </EditorToolbar.Toolbar>
    </EditorToolbar>
  );
}

function Shortcuts({ controller }: { controller: DocxCommandController }) {
  useKeyboardShortcuts({
    commands: controller,
    pagedEditorRef: { current: null },
    disableFindReplaceShortcuts: false,
    tableSelection: { state: { tableIndex: null } } as never,
  });
  return null;
}

function precedes(a: Element, b: Element): boolean {
  return (a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0;
}

function reasonOf(element: Element): string | null {
  const id = element.getAttribute('aria-describedby');
  return id ? (document.getElementById(id)?.textContent ?? null) : null;
}

describe('host-composed toolbar', () => {
  test('renders built-in controls in the host order next to a host action', async () => {
    const { harness, controller } = editor();
    let hostActions = 0;
    render(
      <DocxCommandProvider commands={controller.store}>
        <CompactToolbar onHostAction={() => (hostActions += 1)} />
      </DocxCommandProvider>
    );
    const toolbar = screen().getByRole('toolbar');
    expect(toolbar.closest('.oox-root')).not.toBeNull();
    const style = within(toolbar).getByLabelText('Select paragraph style');
    const bold = within(toolbar).getByRole('button', { name: 'Bold' });
    const undo = within(toolbar).getByRole('button', { name: 'Undo' });
    const review = within(toolbar).getByRole('group', { name: 'Review' });
    const host = within(toolbar).getByRole('button', { name: 'Host action' });
    expect(precedes(style, bold)).toBe(true);
    expect(precedes(bold, undo)).toBe(true);
    expect(precedes(undo, review)).toBe(true);
    expect(precedes(review, host)).toBe(true);
    expect(style.textContent).toContain('Normal text');
    for (const name of ['Accept change', 'Reject change', 'Previous change', 'Next change']) {
      within(review).getByRole('button', { name });
    }

    fireEvent.click(bold);
    fireEvent.click(host);
    await act(async () => {});
    expect(hostActions).toBe(1);
    expect(harness.calls).toEqual([{ id: 'bold', args: null, ordered: true }]);
  });

  test('reflects mixed marks, history and review availability with reasons', async () => {
    const { harness, controller } = editor({ currentRevisionId: null, canUndo: false });
    render(
      <DocxCommandProvider commands={controller.store}>
        <CompactToolbar onHostAction={() => {}} />
      </DocxCommandProvider>
    );
    const bold = screen().getByRole('button', { name: 'Bold' });
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    const undo = screen().getByRole('button', { name: 'Undo' });
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(reasonOf(undo)).toBe('commands.reasons.nothingToUndo');
    const accept = screen().getByRole('button', { name: 'Accept change' });
    expect(reasonOf(accept)).toBe('commands.reasons.revisionRequired');

    harness.update({
      canUndo: true,
      currentRevisionId: 'r1',
      selection: { context: { ...PLAIN_CONTEXT, bold: 'mixed' }, fontFamily: null, fontSize: null },
    });
    act(() => controller.refresh());
    expect(bold.getAttribute('aria-pressed')).toBe('mixed');
    expect(undo.hasAttribute('aria-disabled')).toBe(false);
    expect(accept.hasAttribute('aria-disabled')).toBe(false);

    fireEvent.click(undo);
    fireEvent.click(accept);
    await act(async () => {});
    expect(harness.calls.map((call) => call.id)).toEqual(['undo', 'reviewAccept']);
  });

  test('keyboard shortcuts and toolbar clicks run the same command', async () => {
    const { harness, controller } = editor();
    render(
      <DocxCommandProvider commands={controller.store}>
        <Shortcuts controller={controller} />
        <CompactToolbar onHostAction={() => {}} />
      </DocxCommandProvider>
    );
    const bold = screen().getByRole('button', { name: 'Bold' });
    fireEvent.click(bold);
    fireEvent.keyDown(bold, { key: 'b', ...MOD });
    await act(async () => {});
    expect(harness.calls).toEqual([
      { id: 'bold', args: null, ordered: true },
      { id: 'bold', args: null, ordered: true },
    ]);
  });

  test('keeps two editors isolated', async () => {
    const first = editor();
    const second = editor();
    render(
      <>
        <DocxCommandProvider commands={first.controller.store}>
          <Shortcuts controller={first.controller} />
          <div data-testid="first">
            <CompactToolbar onHostAction={() => {}} />
          </div>
        </DocxCommandProvider>
        <DocxCommandProvider commands={second.controller.store}>
          <Shortcuts controller={second.controller} />
          <div data-testid="second">
            <CompactToolbar onHostAction={() => {}} />
          </div>
        </DocxCommandProvider>
      </>
    );
    const secondBold = within(screen().getByTestId('second')).getByRole('button', { name: 'Bold' });
    fireEvent.click(within(screen().getByTestId('first')).getByRole('button', { name: 'Undo' }));
    fireEvent.keyDown(secondBold, { key: 'b', ...MOD });
    await act(async () => {});
    expect(first.harness.calls.map((call) => call.id)).toEqual(['undo']);
    expect(second.harness.calls.map((call) => call.id)).toEqual(['bold']);

    second.harness.update({ mode: 'viewing' });
    act(() => second.controller.refresh());
    expect(secondBold.getAttribute('aria-disabled')).toBe('true');
    expect(
      within(screen().getByTestId('first')).getByRole('button', { name: 'Bold' }).hasAttribute(
        'aria-disabled'
      )
    ).toBe(false);
  });

  test('a mounted control keeps its snapshot past the cache limit', () => {
    const { harness, controller } = editor();
    const seen: unknown[] = [];
    function Probe() {
      seen.push(useDocxCommandState('fontFamily', { family: 'Georgia' }));
      return null;
    }
    render(
      <DocxCommandProvider commands={controller.store}>
        <Probe />
      </DocxCommandProvider>
    );
    const rendered = seen.length;
    for (let round = 0; round < 3; round += 1) {
      for (let points = 1; points <= 1200; points += 1) {
        controller.store.getState('fontSize', { points: round * 2000 + points });
      }
      harness.update({ canUndo: !controller.store.getState('undo').enabled });
      act(() => controller.refresh());
    }
    expect(seen.length).toBe(rendered);
    expect(controller.store.getState('fontFamily', { family: 'Georgia' }) as unknown).toBe(seen[0]);
  });

  test('external chrome follows the editor locale and color mode', () => {
    const { harness, controller } = editor();
    render(
      <DocxCommandProvider commands={controller.store}>
        <CompactToolbar onHostAction={() => {}} />
      </DocxCommandProvider>
    );
    const root = screen().getByRole('toolbar').closest('.oox-root')!;
    expect(root.classList.contains('dark')).toBe(false);
    expect(screen().queryByRole('button', { name: 'Fett' })).toBeNull();

    harness.state.chrome = {
      i18n: { formattingBar: { bold: 'Fett' } } as never,
      isDark: true,
      theme: null,
    };
    act(() => controller.refresh());
    expect(root.classList.contains('dark')).toBe(true);
    expect(screen().getByRole('button', { name: 'Fett' })).toBeDefined();
  });

  test('shortcuts inside a selector popup reach only the editor that owns it', async () => {
    const first = editor();
    const second = editor();
    render(
      <>
        <DocxCommandProvider commands={first.controller.store}>
          <Shortcuts controller={first.controller} />
          <div data-testid="first">
            <CompactToolbar onHostAction={() => {}} />
          </div>
        </DocxCommandProvider>
        <DocxCommandProvider commands={second.controller.store}>
          <Shortcuts controller={second.controller} />
          <div data-testid="second">
            <CompactToolbar onHostAction={() => {}} />
          </div>
        </DocxCommandProvider>
      </>
    );
    const trigger = within(screen().getByTestId('second')).getByLabelText('Select paragraph style');
    await act(async () => {
      trigger.focus();
      fireEvent.keyDown(trigger, { key: 'Enter' });
    });
    const listbox = screen().getByRole('listbox');
    expect(screen().getByTestId('second').contains(listbox)).toBe(false);
    const saved = fireEvent.keyDown(listbox, { key: 's', ...MOD });
    await act(async () => {});
    expect(saved).toBe(false);
    expect(first.harness.calls).toEqual([]);
    expect(second.harness.calls.map((call) => call.id)).toEqual(['save']);
  });

  test('offers stable unavailable state before an editor is attached', () => {
    function Probe() {
      const state = useDocxCommandState('bold');
      return <output>{state.enabled ? 'enabled' : state.disabledReason.code}</output>;
    }
    render(
      <DocxCommandProvider commands={null}>
        <Probe />
      </DocxCommandProvider>
    );
    expect(screen().getByRole('status').textContent).toBe('editor-unavailable');
  });
});
