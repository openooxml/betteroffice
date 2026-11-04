import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { ImageContextMenu } from './ImageContextMenu';
import { TextContextMenu } from './TextContextMenu';
import { HyperlinkPopup } from './ui/HyperlinkPopup';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function pressOutside() {
  document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, button: 0 }));
  document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }));
}

test('text context menu closes on a press outside right after it opens', () => {
  let closed = 0;
  render(<TextContextMenu isOpen position={{ x: 10, y: 10 }} hasSelection={false} isEditable
    onAction={() => {}} onClose={() => { closed += 1; }} />);
  pressOutside();
  expect(closed).toBe(1);
});

test('text context menu stays open on a press inside it', () => {
  let closed = 0;
  const view = render(<TextContextMenu isOpen position={{ x: 10, y: 10 }} hasSelection={false} isEditable
    onAction={() => {}} onClose={() => { closed += 1; }} />);
  view.getByRole('menu').dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, button: 0 }));
  expect(closed).toBe(0);
});

test('image context menu closes on a press outside right after it opens', () => {
  let closed = 0;
  render(<ImageContextMenu isOpen position={{ x: 10, y: 10 }} currentWrapType="inline"
    onApplyLayout={() => {}} onClose={() => { closed += 1; }} />);
  pressOutside();
  expect(closed).toBe(1);
});

test('link popup closes on a press outside right after it opens', () => {
  let closed = 0;
  render(<HyperlinkPopup data={{ href: 'https://example.com', displayText: 'Example', position: { left: 10, top: 10 } }}
    onNavigate={() => {}} onCopy={() => {}} onEdit={() => {}} onRemove={() => {}} onClose={() => { closed += 1; }} />);
  pressOutside();
  expect(closed).toBe(1);
});

test('text context menu takes focus, moves with arrows and returns focus on Escape', async () => {
  const { fireEvent } = await import('@testing-library/react');
  const origin = document.createElement('textarea');
  document.body.append(origin);
  origin.focus();
  let closed = 0;
  const actions: string[] = [];
  const view = render(<TextContextMenu isOpen position={{ x: 10, y: 10 }} hasSelection isEditable
    onAction={(action) => actions.push(action)} onClose={() => { closed += 1; }} />);
  const menu = view.getByRole('menu', { name: 'Text editing menu' });
  expect(document.activeElement).toBe(menu);
  fireEvent.keyDown(menu, { key: 'ArrowDown' });
  expect(document.activeElement?.getAttribute('role')).toBe('menuitem');
  expect(document.activeElement?.textContent).toContain('Cut');
  fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
  expect(closed).toBe(1);
  expect(actions).toEqual([]);
  expect(document.activeElement).toBe(origin);
  origin.remove();
});
