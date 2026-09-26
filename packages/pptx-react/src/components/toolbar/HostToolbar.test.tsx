import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import type { ReactNode } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  EditorToolbar,
  PptxCommandProvider,
  ToolbarButton,
  ToolbarCommand,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  useEditorToolbar,
  usePptxCommand,
  type ToolbarProps,
} from '../../index';
import { createPptxCommandController } from '../../commands/createPptxCommandStore';
import { formatChord } from '../../commands/descriptors';
import { testBinding } from '../../commands/testing';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

function screen() {
  return within(document.body);
}

const originalRect = HTMLElement.prototype.getBoundingClientRect;

beforeAll(() => {
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    const rect = originalRect.call(this);
    return this.getAttribute('role') === 'toolbar' ? { ...rect.toJSON(), width: 4000 } : rect;
  };
});

afterEach(cleanup);
afterAll(async () => {
  HTMLElement.prototype.getBoundingClientRect = originalRect;
  if (ownsDom) await GlobalRegistrator.unregister();
});

function mount(ui: ReactNode, overrides: Parameters<typeof testBinding>[0] = {}) {
  const harness = testBinding(overrides);
  const controller = createPptxCommandController();
  controller.attach(harness.binding);
  const view = render(<PptxCommandProvider commands={controller.store}>{ui}</PptxCommandProvider>);
  return { harness, controller, view };
}

async function settle() {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

describe('host-composed PPTX toolbar', () => {
  test('arranges built-in controls and a host action from public exports only', async () => {
    let shared = 0;
    const { harness } = mount(
      <EditorToolbar mode="commands">
        <EditorToolbar.Toolbar>
          <ToolbarGroup label="Formatting">
            <ToolbarCommandSelect id="fontFamily" />
            <ToolbarCommandButton id="bold" />
          </ToolbarGroup>
          <ToolbarCommandButton id="undo" />
          <ToolbarCommandButton id="slideshow" />
          <ToolbarButton title="Share" onClick={() => (shared += 1)}>
            Share
          </ToolbarButton>
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    );

    const toolbar = screen().getByRole('toolbar', { name: 'Presentation formatting toolbar' });
    const names = within(toolbar)
      .getAllByRole('button')
      .map((button) => button.getAttribute('aria-label'));
    expect(names).toEqual(['Font family', 'Bold', 'Undo', 'Present', 'Share']);
    expect(screen().queryByTestId('pptx-save')).toBeNull();

    fireEvent.click(screen().getByTestId('pptx-bold'));
    fireEvent.click(screen().getByTestId('pptx-undo'));
    fireEvent.click(screen().getByRole('button', { name: 'Share' }));
    fireEvent.click(screen().getByTestId('pptx-font-family'));
    fireEvent.click(screen().getByRole('menuitemradio', { name: 'Georgia' }));
    await settle();

    expect(harness.calls).toEqual([
      { id: 'bold', args: null },
      { id: 'undo', args: null },
      { id: 'fontFamily', args: { family: 'Georgia' } },
    ]);
    expect(shared).toBe(1);
  });

  test('announces mixed marks, platform shortcuts and disabled reasons', () => {
    mount(
      <EditorToolbar mode="commands">
        <EditorToolbar.Toolbar>
          <ToolbarCommandButton id="bold" />
          <ToolbarCommandButton id="undo" />
        </EditorToolbar.Toolbar>
      </EditorToolbar>,
      {
        canUndo: false,
        text: {
          kind: 'range',
          bold: 'mixed',
          italic: false,
          underline: false,
          fontFamily: null,
          fontSize: null,
          color: null,
          alignment: null,
        },
      }
    );
    const bold = screen().getByTestId('pptx-bold');
    expect(bold.getAttribute('aria-pressed')).toBe('mixed');
    expect(bold.title).toBe(`Bold (${formatChord('Mod+B')})`);

    const undo = screen().getByTestId('pptx-undo');
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.hasAttribute('disabled')).toBe(false);
    const description = document.getElementById(undo.getAttribute('aria-describedby')!);
    expect(description?.textContent).toBe('There is nothing to undo.');
    expect(undo.title).toBe(`Undo (${formatChord('Mod+Z')}): There is nothing to undo.`);
  });

  test('renders the default rail when command-mode chrome has no children', async () => {
    const { harness } = mount(<EditorToolbar mode="commands" />);
    expect(screen().getByTestId('pptx-editor-toolbar')).toBeDefined();
    fireEvent.click(screen().getByTestId('pptx-save'));
    fireEvent.click(screen().getByTestId('pptx-tool-text-box'));
    await settle();
    expect(harness.calls).toEqual([
      { id: 'save', args: null },
      { id: 'tool', args: { value: 'textBox' } },
    ]);
  });

  test('refuses legacy state and callbacks in command mode', () => {
    const error = console.error;
    console.error = () => {};
    try {
      const props = { mode: 'commands', onFormat: () => {} } as unknown as Parameters<
        typeof EditorToolbar
      >[0];
      expect(() => mount(<EditorToolbar {...props} />)).toThrow(/onFormat/);
      expect(() =>
        mount(
          <EditorToolbar mode="commands">
            <EditorToolbar.Toolbar {...({ canUndo: true } as object)} />
          </EditorToolbar>
        )
      ).toThrow(/canUndo/);
    } finally {
      console.error = error;
    }
  });

  test('keeps legacy props, context overrides and appended host controls', () => {
    const saved: string[] = [];
    const seen: ToolbarProps[] = [];
    function Probe() {
      seen.push(useEditorToolbar());
      return <span>Probe</span>;
    }
    render(
      <EditorToolbar onSave={() => saved.push('context')} canUndo onUndo={() => saved.push('undo')}>
        <Probe />
        <EditorToolbar.Toolbar onSave={() => saved.push('explicit')}>
          <button type="button">Host extra</button>
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    );
    fireEvent.click(screen().getByTestId('pptx-save'));
    fireEvent.click(screen().getByTestId('pptx-undo'));
    expect(saved).toEqual(['explicit', 'undo']);
    expect(seen[seen.length - 1].canUndo).toBe(true);
    const buttons = Array.from(document.querySelectorAll('button'));
    const extra = buttons.findIndex((button) => button.textContent === 'Host extra');
    const arrange = buttons.indexOf(
      screen().getByTestId('pptx-shape-arrange') as HTMLButtonElement
    );
    expect(extra).toBeGreaterThan(arrange);
  });

  test('projects command state into the legacy context inside command mode', async () => {
    const seen: ToolbarProps[] = [];
    function Probe() {
      const props = useEditorToolbar();
      seen.push(props);
      return (
        <button type="button" onClick={() => props.onFormat?.('italic')}>
          Legacy italic
        </button>
      );
    }
    const { harness } = mount(
      <EditorToolbar mode="commands">
        <Probe />
      </EditorToolbar>,
      {
        text: {
          kind: 'range',
          bold: 'mixed',
          italic: true,
          underline: false,
          fontFamily: 'Arial',
          fontSize: null,
          color: '#112233',
          alignment: 'ctr',
        },
      }
    );
    await settle();
    const props = seen[seen.length - 1];
    expect(props.currentFormatting).toEqual({
      italic: true,
      underline: false,
      fontFamily: 'Arial',
      textColor: '#112233',
      align: 'ctr',
    });
    expect(props.textSelectionActive).toBe(true);
    expect(props.canUndo).toBe(true);
    fireEvent.click(screen().getByRole('button', { name: 'Legacy italic' }));
    await settle();
    expect(harness.calls).toEqual([{ id: 'italic', args: null }]);
  });

  test('follows the editor locale outside the editor and binds custom controls', async () => {
    function HostBold() {
      const bold = usePptxCommand('bold');
      return (
        <button
          type="button"
          aria-pressed={bold.state.active === true}
          onClick={() => void bold.execute()}
        >
          Host {bold.label}
        </button>
      );
    }
    const harness = testBinding();
    harness.state.chrome = { i18n: { toolbar: { bold: 'Fett' } } };
    const controller = createPptxCommandController();
    controller.attach(harness.binding);
    render(
      <PptxCommandProvider commands={controller.store}>
        <EditorToolbar mode="commands">
          <EditorToolbar.Toolbar>
            <ToolbarCommandButton id="bold" />
          </EditorToolbar.Toolbar>
          <HostBold />
        </EditorToolbar>
      </PptxCommandProvider>
    );
    expect(screen().getByTestId('pptx-bold').getAttribute('aria-label')).toBe('Fett');
    fireEvent.click(screen().getByRole('button', { name: 'Host Fett', pressed: false }));
    await settle();
    expect(harness.calls).toEqual([{ id: 'bold', args: null }]);
  });

  test('gives legacy dropdowns menu or dialog semantics by their content', () => {
    const picked: string[] = [];
    render(
      <>
        <ToolbarDropdown title="Layouts" trigger="L">
          {(close) => (
            <>
              <ToolbarMenuItem label="Title" onClick={() => picked.push('title')} close={close} />
              <ToolbarMenuItem
                label="Two columns"
                onClick={() => picked.push('two')}
                close={close}
              />
            </>
          )}
        </ToolbarDropdown>
        <ToolbarDropdown title="Options" trigger="O">
          {() => <input aria-label="Width" />}
        </ToolbarDropdown>
      </>
    );
    const layouts = screen().getByRole('button', { name: 'Layouts' });
    act(() => layouts.focus());
    fireEvent.keyDown(layouts, { key: 'ArrowDown' });
    const menu = screen().getByRole('menu', { name: 'Layouts' });
    expect(document.activeElement).toBe(within(menu).getByRole('menuitem', { name: 'Title' }));
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
    expect(document.activeElement).toBe(
      within(menu).getByRole('menuitem', { name: 'Two columns' })
    );
    fireEvent.click(document.activeElement!);
    expect(picked).toEqual(['two']);
    expect(document.activeElement).toBe(layouts);

    const options = screen().getByRole('button', { name: 'Options' });
    fireEvent.click(options);
    const dialog = screen().getByRole('dialog', { name: 'Options' });
    expect(options.getAttribute('aria-haspopup')).toBe('dialog');
    expect(document.activeElement).toBe(within(dialog).getByLabelText('Width'));
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    expect(screen().queryByRole('dialog', { name: 'Options' })).toBeNull();
    expect(document.activeElement).toBe(options);
  });

  test('requires the arguments a command cannot run without', async () => {
    // @ts-expect-error accepting needs the proposal to accept
    const missing = <ToolbarCommand id="proposalAccept" />;
    // @ts-expect-error an alignment button binds one alignment
    const unbound = <ToolbarCommandButton id="alignment" />;
    const { harness } = mount(
      <>
        {missing}
        {unbound}
        <ToolbarCommand id="proposalAccept" args={{ proposalId: 'p1' }} />
        <ToolbarCommand id="alignment" />
      </>
    );
    const buttons = screen().getAllByRole('button', { name: 'Accept proposal' });
    expect(buttons[0].getAttribute('aria-disabled')).toBe('true');
    expect(buttons[0].title).toBe('Accept proposal: The command received invalid arguments.');
    fireEvent.click(buttons[0]);
    fireEvent.click(buttons[1]);
    await settle();
    expect(harness.calls).toEqual([{ id: 'proposalAccept', args: { proposalId: 'p1' } }]);
    expect(screen().getAllByTestId('pptx-align-left')).toHaveLength(1);
  });

  test('a null provider reports the editor as unavailable', () => {
    render(
      <PptxCommandProvider commands={null}>
        <ToolbarCommandButton id="bold" />
      </PptxCommandProvider>
    );
    const bold = screen().getByTestId('pptx-bold');
    expect(bold.getAttribute('aria-disabled')).toBe('true');
    expect(bold.title).toContain('The editor is not ready.');
  });
});
