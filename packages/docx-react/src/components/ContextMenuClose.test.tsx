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
