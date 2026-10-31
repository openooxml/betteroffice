import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { ParagraphSpacingDialog } from './ParagraphSpacingDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

test('shows current paragraph spacing in points and applies twips', () => {
  const onApply = mock(() => {});
  const onClose = mock(() => {});
  const view = render(<ParagraphSpacingDialog isOpen currentProps={{ before: 240, after: 160, lineSpacing: 276 }}
    onApply={onApply} onClose={onClose} />);
  const before = view.getByLabelText('Spacing Before') as HTMLInputElement;
  const after = view.getByLabelText('Spacing After') as HTMLInputElement;
  expect(before.value).toBe('12');
  expect(after.value).toBe('8');
  expect(document.activeElement).toBe(before);
  expect((view.getByLabelText('Line spacing') as HTMLSelectElement).value).toBe('1.15');
  fireEvent.change(before, { target: { value: '18' } });
  fireEvent.change(after, { target: { value: '6' } });
  fireEvent.change(view.getByLabelText('Line spacing'), { target: { value: '1.5' } });
  fireEvent.keyDown(before, { key: 'Enter' });
  expect(onApply).toHaveBeenCalledWith({ before: 360, after: 120, lineSpacing: 360 });
  expect(onClose).toHaveBeenCalledTimes(1);
});

test('Escape closes without applying', () => {
  const onApply = mock(() => {});
  const onClose = mock(() => {});
  const view = render(<ParagraphSpacingDialog isOpen onApply={onApply} onClose={onClose} />);
  fireEvent.keyDown(view.getByLabelText('Spacing Before'), { key: 'Escape' });
  expect(onClose).toHaveBeenCalledTimes(1);
  expect(onApply).not.toHaveBeenCalled();
});

test('accepts a custom line multiple and preserves unchanged line spacing', () => {
  const onApply = mock(() => {});
  const view = render(<ParagraphSpacingDialog isOpen currentProps={{ lineSpacing: 300 }} onApply={onApply} onClose={() => {}} />);
  const multiple = view.getByLabelText('Multiple') as HTMLInputElement;
  expect(multiple.value).toBe('1.25');
  fireEvent.click(view.getByRole('button', { name: 'Apply' }));
  expect(onApply).toHaveBeenLastCalledWith({ before: 0, after: 0 });
  fireEvent.change(multiple, { target: { value: '1.75' } });
  fireEvent.submit(view.getByRole('dialog'));
  expect(onApply).toHaveBeenLastCalledWith({ before: 0, after: 0, lineSpacing: 420 });
});

test('rejects empty, negative and excessive spacing', () => {
  const onApply = mock(() => {});
  const view = render(<ParagraphSpacingDialog isOpen onApply={onApply} onClose={() => {}} />);
  const before = view.getByLabelText('Spacing Before');
  for (const value of ['', '-1', '1585']) {
    fireEvent.change(before, { target: { value } });
    expect((view.getByRole('button', { name: 'Apply' }) as HTMLButtonElement).disabled).toBe(true);
    fireEvent.submit(view.getByRole('dialog'));
  }
  expect(onApply).not.toHaveBeenCalled();
  fireEvent.change(before, { target: { value: '1584' } });
  fireEvent.submit(view.getByRole('dialog'));
  expect(onApply).toHaveBeenCalledWith({ before: 31680, after: 0 });
});
