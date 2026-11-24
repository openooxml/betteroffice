import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import type { DocxCommandId } from '../commands/types';
import type { DocxCommandEnvironment } from '../commands/evaluate';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createDocxCommandController } from '../commands/createDocxCommandStore';
import { testBinding } from '../commands/testing';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');
const { DocxCommandProvider, EditorToolbar } = await import('../index');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const ALL_SUPPORTED = new Set<DocxCommandId>([
  'insertImage',
  'insertTable',
  'insertPageBreak',
  'insertSectionBreakNextPage',
  'insertSectionBreakContinuous',
  'insertTOC',
  'watermark',
]);

function mount(placement: 'menubar' | 'rail', environment: Partial<DocxCommandEnvironment> = {}) {
  const harness = testBinding(environment);
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  const view = render(
    <DocxCommandProvider commands={controller.store}>
      <EditorToolbar>
        {placement === 'menubar' ? (
          <EditorToolbar.TitleBar>
            <EditorToolbar.MenuBar />
          </EditorToolbar.TitleBar>
        ) : (
          <EditorToolbar.Toolbar>
            <EditorToolbar.InsertMenu />
          </EditorToolbar.Toolbar>
        )}
      </EditorToolbar>
    </DocxCommandProvider>
  );
  const trigger =
    placement === 'menubar'
      ? view.getByRole('menuitem', { name: 'Insert' })
      : within(view.getByRole('toolbar')).getByRole('button', {
          name: 'Insert',
        });
  return { harness, view, trigger };
}

function open(trigger: HTMLElement): HTMLElement {
  act(() => trigger.focus());
  fireEvent.keyDown(trigger, { key: 'Enter' });
  return document.getElementById(trigger.getAttribute('aria-controls')!)!;
}

function items(menu: HTMLElement) {
  return Array.from(menu.querySelectorAll<HTMLElement>('[role="menuitem"]'))
    .filter((item) => item.closest('[role="menu"]') === menu)
    .map((item) => ({
      label: Array.from(
        item.querySelectorAll('span:not([hidden])'),
        (span) => span.textContent
      ).join(''),
      disabled: item.getAttribute('aria-disabled') === 'true',
      reason: item.getAttribute('title'),
    }));
}

function openSubmenu(menu: HTMLElement, label: string): HTMLElement {
  const item = within(menu).getByRole('menuitem', {
    name: new RegExp(`^${label}`),
  });
  act(() => item.focus());
  fireEvent.keyDown(item, { key: 'ArrowRight' });
  return document.getElementById(item.getAttribute('aria-controls')!)!;
}

describe('InsertMenu', () => {
  test('renders the same entries in a custom rail as in the menu bar', () => {
    const menubar = mount('menubar', { supportedCommands: ALL_SUPPORTED });
    const expected = items(open(menubar.trigger));
    const breaks = items(openSubmenu(open(menubar.trigger), 'Break'));
    cleanup();

    const rail = mount('rail', { supportedCommands: ALL_SUPPORTED });
    expect(rail.trigger.getAttribute('role')).toBeNull();
    expect(items(open(rail.trigger))).toEqual(expected);
    expect(items(openSubmenu(open(rail.trigger), 'Break'))).toEqual(breaks);
    expect(expected.map((item) => item.label)).toEqual(['Image', 'Table', 'Break', 'Watermark']);
    expect(breaks.map((item) => item.label)).toEqual([
      'Page break',
      'Section break (next page)',
      'Section break (continuous)',
    ]);
  });

  test('inserts a table from the grid in a custom rail', async () => {
    const { harness, trigger } = mount('rail');
    const grid = openSubmenu(open(trigger), 'Table');
    fireEvent.click(grid.querySelector('[data-row="2"][data-column="3"]')!);
    await act(async () => {});
    expect(harness.calls.at(-1)).toEqual({
      id: 'insertTable',
      args: { rows: 2, columns: 3 },
      ordered: true,
    });
    expect(document.getElementById(trigger.getAttribute('aria-controls')!)).toBeNull();
  });

  test.each(['menubar', 'rail'] as const)(
    'hides commands the host leaves out and explains a disabled table in the %s',
    (placement) => {
      const supported = new Set(ALL_SUPPORTED);
      supported.delete('insertImage');
      supported.delete('watermark');
      const { trigger } = mount(placement, {
        supportedCommands: supported,
        mode: 'suggesting',
      });
      const entries = items(open(trigger));
      expect(entries.map((item) => item.label)).toEqual(['Table', 'Break']);
      expect(entries[0].disabled).toBe(true);
      expect(entries[0].reason).toBe('commands.reasons.suggestingUnsupported');
    }
  );
});
