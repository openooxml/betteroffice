import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { useState } from 'react';
import type { PasteFormattingOption } from '@betteroffice/docx/yrs';
import { PasteOptionsControl, menuPlacement } from './PasteOptionsControl';
import { usePasteOptions } from './usePasteOptions';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, screen } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function Harness(props: { enabled?: boolean; onChoose?: (option: PasteFormattingOption) => void }) {
  const [focused, setFocused] = useState(0);
  const paste = usePasteOptions({
    enabled: props.enabled ?? true,
    onChoose: props.onChoose ?? (() => {}),
    focusEditor: () => setFocused(count => count + 1),
  });
  return (
    <div>
      <div
        data-testid="surface"
        tabIndex={0}
        onKeyDown={(event) => { paste.handleKeyDown(event); }}
        onKeyUp={(event) => { paste.handleKeyUp(event); }}
      />
      <button type="button" onClick={() => paste.show({ left: 40, top: 60 })}>show</button>
      <button type="button" onClick={() => paste.dismiss()}>dismiss</button>
      <span data-testid="state">
        {paste.visible ? 'visible' : 'hidden'}:{paste.selected}:{focused}
      </span>
      {paste.visible && paste.anchor && (
        <PasteOptionsControl
          anchor={paste.anchor}
          menuOpen={paste.menuOpen}
          selected={paste.selected}
          onMenuOpenChange={paste.setMenuOpen}
          onChoose={paste.choose}
          onCloseMenu={paste.closeMenu}
        />
      )}
    </div>
  );
}

test('shows after a formatted paste and hides on dismiss', () => {
  render(<Harness />);
  expect(screen.getByTestId('state').textContent).toContain('hidden');
  fireEvent.click(screen.getByText('show'));
  expect(screen.getByTestId('state').textContent).toContain('visible:keep-source');
  expect(screen.getByRole('button', { name: 'Paste Options (Ctrl)' })).toBeTruthy();
  fireEvent.click(screen.getByText('dismiss'));
  expect(screen.getByTestId('state').textContent).toContain('hidden');
  expect(screen.queryByRole('button', { name: 'Paste Options (Ctrl)' })).toBeNull();
});

test('pasteOptions disabled never shows the control', () => {
  render(<Harness enabled={false} />);
  fireEvent.click(screen.getByText('show'));
  expect(screen.getByTestId('state').textContent).toContain('hidden');
  expect(screen.queryByRole('button', { name: 'Paste Options (Ctrl)' })).toBeNull();
});

test('menu exposes menuitemradio roles with aria-checked and chooses an option', () => {
  const chosen: PasteFormattingOption[] = [];
  render(<Harness onChoose={(option) => { chosen.push(option); }} />);
  fireEvent.click(screen.getByText('show'));
  const button = screen.getByRole('button', { name: 'Paste Options (Ctrl)' });
  expect(button.getAttribute('aria-haspopup')).toBe('menu');
  expect(button.getAttribute('aria-expanded')).toBe('false');
  fireEvent.click(button);
  expect(button.getAttribute('aria-expanded')).toBe('true');
  const keep = screen.getByRole('menuitemradio', { name: 'Keep Source Formatting' });
  const merge = screen.getByRole('menuitemradio', { name: 'Merge Formatting' });
  const textOnly = screen.getByRole('menuitemradio', { name: 'Keep Text Only' });
  expect(keep.getAttribute('aria-checked')).toBe('true');
  expect(merge.getAttribute('aria-checked')).toBe('false');
  expect(textOnly.getAttribute('aria-checked')).toBe('false');
  fireEvent.click(merge);
  expect(chosen).toEqual(['merge']);
  expect(screen.getByTestId('state').textContent).toContain('visible:merge:1');
  fireEvent.click(button);
  expect(screen.getByRole('menuitemradio', { name: 'Merge Formatting' }).getAttribute('aria-checked')).toBe('true');
  expect(screen.getByRole('menuitemradio', { name: 'Keep Source Formatting' }).getAttribute('aria-checked')).toBe('false');
});

test('Ctrl alone opens the menu; Ctrl with another key does not', () => {
  render(<Harness />);
  fireEvent.click(screen.getByText('show'));
  const surface = screen.getByTestId('surface');
  fireEvent.keyDown(surface, { key: 'Control' });
  fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
  fireEvent.keyUp(surface, { key: 'Control' });
  expect(screen.queryByRole('menu')).toBeNull();
  fireEvent.keyDown(surface, { key: 'Control' });
  fireEvent.keyUp(surface, { key: 'Control' });
  expect(screen.getByRole('menu')).toBeTruthy();
});

test('Escape closes the menu and returns focus to the editor', () => {
  render(<Harness />);
  fireEvent.click(screen.getByText('show'));
  fireEvent.click(screen.getByRole('button', { name: 'Paste Options (Ctrl)' }));
  expect(screen.getByRole('menu')).toBeTruthy();
  fireEvent.keyDown(screen.getByRole('menu'), { key: 'Escape' });
  expect(screen.queryByRole('menu')).toBeNull();
  expect(screen.getByTestId('state').textContent).toContain(':1');
});

test('arrow keys move between menu items and Enter chooses', () => {
  const chosen: PasteFormattingOption[] = [];
  render(<Harness onChoose={(option) => { chosen.push(option); }} />);
  fireEvent.click(screen.getByText('show'));
  fireEvent.click(screen.getByRole('button', { name: 'Paste Options (Ctrl)' }));
  const menu = screen.getByRole('menu');
  act(() => { screen.getByRole('menuitemradio', { name: 'Keep Source Formatting' }).focus(); });
  fireEvent.keyDown(menu, { key: 'ArrowDown' });
  fireEvent.keyDown(screen.getByRole('menuitemradio', { name: 'Merge Formatting' }), { key: 'Enter' });
  expect(chosen).toEqual(['merge']);
});

test('scroll keeps the control and resize hides it', () => {
  render(<Harness />);
  fireEvent.click(screen.getByText('show'));
  expect(screen.getByRole('button', { name: 'Paste Options (Ctrl)' })).toBeTruthy();
  fireEvent.scroll(window);
  expect(screen.getByRole('button', { name: 'Paste Options (Ctrl)' })).toBeTruthy();
  fireEvent(window, new Event('resize'));
  expect(screen.queryByRole('button', { name: 'Paste Options (Ctrl)' })).toBeNull();
});

test('the menu opens below the button and flips above it or shifts left at the viewport edges', () => {
  const viewport = { width: 1000, height: 800 };
  const menu = { width: 220, height: 110 };
  expect(menuPlacement({ left: 100, top: 100, bottom: 124 }, menu, viewport)).toEqual({ left: 0, top: 28 });
  expect(menuPlacement({ left: 100, top: 740, bottom: 764 }, menu, viewport)).toEqual({ left: 0, top: -114 });
  expect(menuPlacement({ left: 900, top: 100, bottom: 124 }, menu, viewport)).toEqual({ left: -128, top: 28 });
});
