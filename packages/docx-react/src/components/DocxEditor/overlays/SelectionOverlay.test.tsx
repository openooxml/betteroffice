import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { SelectionOverlay } from './SelectionOverlay';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const rect = { x: 10, y: 20, width: 30, height: 12, pageIndex: 0 };
const caret = { x: 10, y: 20, height: 12, pageIndex: 0 };

test('read-only draws a range selection and never the caret', () => {
  const range = render(
    <SelectionOverlay selectionRects={[rect]} caretPosition={null} isFocused readOnly />
  );
  expect(range.queryByTestId('selection-rect-0')).not.toBeNull();
  range.unmount();

  const collapsed = render(
    <SelectionOverlay selectionRects={[]} caretPosition={caret} isFocused readOnly />
  );
  expect(collapsed.queryByTestId('selection-overlay')).toBeNull();
  expect(collapsed.queryByTestId('caret')).toBeNull();
});

test('an editable overlay still draws the caret', () => {
  const view = render(<SelectionOverlay selectionRects={[]} caretPosition={caret} isFocused />);
  expect(view.queryByTestId('caret')).not.toBeNull();
});
