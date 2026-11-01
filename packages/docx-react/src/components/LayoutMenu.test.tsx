import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { useState } from 'react';
import type { SectionProperties } from '@betteroffice/docx/types/document';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import { DocxCommandProvider } from '../commands/DocxCommandProvider';
import { testBinding } from '../commands/testing';
import type { DocxCommandEnvironment } from '../commands/evaluate';
import { EditorChromeContext } from './EditorToolbarContext';
import { MenuBar } from './TitleBar';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function mount(
  initial: SectionProperties = {},
  unit: 'inch' | 'cm' = 'inch',
  environment: Partial<DocxCommandEnvironment> = {},
  readCurrent?: () => Promise<void>
) {
  const harness = testBinding(environment);
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  const applied: Partial<SectionProperties>[] = [];
  function Chrome() {
    const [properties, setProperties] = useState(initial);
    return <DocxCommandProvider commands={controller.store}>
      <EditorChromeContext.Provider value={{ showZoomControl: true, toolbarExtra: null, pageLayout: {
        currentProps: properties,
        unit,
        readCurrent,
        onApply: patch => { applied.push(patch); setProperties(current => ({ ...current, ...patch })); },
      } }}>
        <MenuBar />
      </EditorChromeContext.Provider>
    </DocxCommandProvider>;
  }
  const view = render(<Chrome />);
  const open = (name: string) => {
    const trigger = view.getByRole('menuitem', { name: 'Layout' });
    act(() => trigger.focus());
    fireEvent.keyDown(trigger, { key: 'Enter' });
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    fireEvent.keyDown(view.getByRole('menuitem', { name: new RegExp(`^${name}$`, 'i') }), { key: 'ArrowRight' });
  };
  return { ...view, applied, harness, controller, open };
}

describe('Layout menubar menu', () => {
  test('reads worker properties only when Layout opens and waits before enabling page choices', async () => {
    let reads = 0;
    let finish: (() => void) | undefined;
    const view = mount({}, 'inch', {}, () => {
      reads++;
      return new Promise<void>(resolve => { finish = resolve; });
    });
    expect(reads).toBe(0);
    const trigger = view.getByRole('menuitem', { name: 'Layout' });
    act(() => trigger.focus());
    fireEvent.keyDown(trigger, { key: 'Enter' });
    expect(reads).toBe(1);
    const margins = view.getByRole('menuitem', { name: /^margins$/i });
    expect(margins.getAttribute('aria-disabled')).toBeNull();
    fireEvent.keyDown(margins, { key: 'ArrowRight' });
    expect(view.getByRole('menuitemradio', { name: 'Normal' }).getAttribute('aria-disabled')).toBe('true');
    await act(async () => { finish!(); });
    expect(view.getByRole('menuitemradio', { name: 'Normal' }).getAttribute('aria-disabled')).toBeNull();
    expect(view.getByRole('menuitemradio', { name: 'Normal' }).getAttribute('aria-checked')).toBe('true');
    fireEvent.keyDown(view.getByRole('menuitemradio', { name: 'Normal' }), { key: 'Escape' });
    fireEvent.keyDown(margins, { key: 'Escape' });
    fireEvent.click(trigger);
    expect(reads).toBe(2);
    await act(async () => { finish!(); });
  });

  test('applies Narrow through page setup and updates the checked preset', () => {
    const view = mount();
    view.open('Margins');
    const normal = view.getByRole('menuitemradio', { name: 'Normal' });
    const narrow = view.getByRole('menuitemradio', { name: 'Narrow' });
    expect(normal.getAttribute('aria-checked')).toBe('true');
    expect(narrow.getAttribute('aria-checked')).toBe('false');
    expect(document.activeElement === normal).toBe(true);
    fireEvent.keyDown(normal, { key: 'ArrowDown' });
    expect(document.activeElement === narrow).toBe(true);
    fireEvent.click(narrow);
    expect(view.applied).toEqual([{ marginTop: 720, marginBottom: 720, marginLeft: 720, marginRight: 720 }]);
    expect(view.queryByRole('menu', { name: 'Layout' })).toBeNull();
    view.open('Margins');
    expect(view.getByRole('menuitemradio', { name: 'Narrow' }).getAttribute('aria-checked')).toBe('true');
    expect(view.getByRole('menuitemradio', { name: 'Normal' }).getAttribute('aria-checked')).toBe('false');
  });

  test('Landscape swaps the current dimensions and reflects the new orientation', () => {
    const view = mount({ pageWidth: 11906, pageHeight: 16838 });
    view.open('Orientation');
    expect(view.getByRole('menuitemradio', { name: 'Portrait' }).getAttribute('aria-checked')).toBe('true');
    fireEvent.click(view.getByRole('menuitemradio', { name: 'Landscape' }));
    expect(view.applied).toEqual([{ orientation: 'landscape', pageWidth: 16838, pageHeight: 11906 }]);
    view.open('Orientation');
    expect(view.getByRole('menuitemradio', { name: 'Landscape' }).getAttribute('aria-checked')).toBe('true');
    expect(view.getByRole('menuitemradio', { name: 'Portrait' }).getAttribute('aria-checked')).toBe('false');
  });

  test('paper sizes preserve landscape and show the ruler unit', () => {
    const view = mount({ pageWidth: 15840, pageHeight: 12240, orientation: 'landscape' }, 'cm');
    view.open('Size');
    const letter = view.getByRole('menuitemradio', { name: 'Letter' });
    expect(letter.getAttribute('aria-checked')).toBe('true');
    expect(letter.textContent).toContain('27.94 cm × 21.59 cm');
    fireEvent.click(view.getByRole('menuitemradio', { name: 'A4' }));
    expect(view.applied).toEqual([{ pageWidth: 16838, pageHeight: 11906, orientation: 'landscape' }]);
    view.open('Size');
    expect(view.getByRole('menuitemradio', { name: 'A4' }).getAttribute('aria-checked')).toBe('true');
  });

  test('custom margins use the existing page setup command', async () => {
    const view = mount({ marginLeft: 1000 }, 'cm');
    view.open('Margins');
    expect(view.getByRole('menuitemradio', { name: 'Normal' }).getAttribute('aria-checked')).toBe('false');
    expect(view.getByRole('menuitemradio', { name: 'Normal' }).textContent).toContain('Top 2.54 cm');
    await act(async () => { fireEvent.click(view.getByRole('menuitem', { name: 'Custom margins...' })); });
    expect(view.harness.calls.map(call => call.id)).toEqual(['pageSetup']);
    expect(view.applied).toEqual([]);
  });

  test('more paper sizes use the existing page setup command', async () => {
    const view = mount();
    view.open('Size');
    await act(async () => { fireEvent.click(view.getByRole('menuitem', { name: 'More paper sizes...' })); });
    expect(view.harness.calls.map(call => call.id)).toEqual(['pageSetup']);
  });

  test.each([{ readOnly: true }, { mode: 'viewing' as const }])('disables page items when page setup is disabled: %j', environment => {
    const view = mount({}, 'inch', environment);
    fireEvent.keyDown(view.getByRole('menuitem', { name: 'Layout' }), { key: 'Enter' });
    for (const name of ['Margins', 'Orientation', 'Size']) {
      const item = view.getByRole('menuitem', { name: new RegExp(`^${name}$`, 'i') });
      expect(item.getAttribute('aria-disabled')).toBe('true');
      fireEvent.keyDown(item, { key: 'Enter' });
      fireEvent.click(item);
      expect(view.queryByRole('menu', { name: new RegExp(`^${name}$`, 'i') })).toBeNull();
    }
    expect(view.applied).toEqual([]);
  });

  test('disables open choices when the page setup command becomes unavailable', () => {
    const view = mount();
    view.open('Margins');
    act(() => { view.harness.update({ readOnly: true }); view.controller.refresh(); });
    const narrow = view.getByRole('menuitemradio', { name: 'Narrow' });
    expect(narrow.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(narrow);
    expect(view.applied).toEqual([]);
  });

  test('indent choices execute the existing commands', async () => {
    const view = mount();
    for (const name of ['Increase Indent', 'Decrease Indent']) {
      view.open('Indent');
      await act(async () => { fireEvent.click(view.getByRole('menuitem', { name })); });
    }
    expect(view.harness.calls.map(call => call.id)).toEqual(['indent', 'outdent']);
  });
});
