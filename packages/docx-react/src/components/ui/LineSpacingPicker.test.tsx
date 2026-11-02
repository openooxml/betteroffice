import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { LineSpacingPicker } from './LineSpacingPicker';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

for (const spacing of [undefined, 0, 240]) {
  test(`shows and applies ${spacing ? 'Remove' : 'Add'} actions for spacing ${spacing}`, () => {
    const onParagraphSpacingChange = mock(() => {});
    const view = render(<LineSpacingPicker spaceBefore={spacing} spaceAfter={spacing}
      onParagraphSpacingChange={onParagraphSpacingChange} onOpenOptions={() => {}} />);
    const trigger = view.getByRole('button', { name: 'Line spacing' });
    fireEvent.click(trigger);
    const action = spacing ? 'Remove' : 'Add';
    fireEvent.click(view.getByRole('menuitem', { name: `${action} Space Before Paragraph` }));
    expect(onParagraphSpacingChange).toHaveBeenLastCalledWith({ before: spacing ? 0 : 240 });
    fireEvent.click(trigger);
    fireEvent.click(view.getByRole('menuitem', { name: `${action} Space After Paragraph` }));
    expect(onParagraphSpacingChange).toHaveBeenLastCalledWith({ after: spacing ? 0 : 240 });
  });
}

test('preserves presets and reaches paragraph actions with arrow keys', () => {
  const onChange = mock(() => {});
  const onOpenOptions = mock(() => {});
  const view = render(<LineSpacingPicker value={360} onChange={onChange} onOpenOptions={onOpenOptions}
    onParagraphSpacingChange={() => {}} />);
  const trigger = view.getByRole('button', { name: 'Line spacing' });
  fireEvent.keyDown(trigger, { key: 'ArrowDown' });
  expect(view.getByRole('menuitemradio', { name: '1.5' }).getAttribute('aria-checked')).toBe('true');
  const first = view.getByRole('menuitemradio', { name: 'Single' });
  expect(document.activeElement).toBe(first);
  fireEvent.keyDown(first, { key: 'ArrowDown' });
  const next = view.getByRole('menuitemradio', { name: '1.15' });
  expect(document.activeElement).toBe(next);
  fireEvent.click(next);
  expect(onChange).toHaveBeenCalledWith(276);
  fireEvent.keyDown(trigger, { key: 'ArrowUp' });
  const options = view.getByRole('menuitem', { name: 'Line Spacing Options...' });
  expect(document.activeElement).toBe(options);
  fireEvent.keyDown(options, { key: 'ArrowUp' });
  expect(document.activeElement).toBe(view.getByRole('menuitem', { name: 'Add Space After Paragraph' }));
  fireEvent.click(options);
  expect(onOpenOptions).toHaveBeenCalledTimes(1);
});

test('disabled picker does not open', () => {
  const view = render(<LineSpacingPicker disabled description="Read only" />);
  fireEvent.keyDown(view.getByRole('button', { name: 'Line spacing' }), { key: 'ArrowDown' });
  fireEvent.click(view.getByRole('button', { name: 'Line spacing' }));
  expect(view.queryByRole('menu')).toBeNull();
});
