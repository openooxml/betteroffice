import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import type { SectionProperties } from '@betteroffice/docx/types/document';
import { PageSetupDialog } from './PageSetupDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const current: SectionProperties = {
  pageWidth: 11906, pageHeight: 16838, marginTop: 1440, marginBottom: 1134, marginLeft: 1418, marginRight: 720,
};

test('shows and applies margins in centimeters', () => {
  const onApply = mock((_props: Partial<SectionProperties>) => {});
  const view = render(<PageSetupDialog isOpen onClose={() => {}} onApply={onApply} currentProps={current} unit="cm" />);
  const value = (label: string) => (view.getByLabelText(label) as HTMLInputElement).value;
  expect([value('Top'), value('Bottom'), value('Left'), value('Right')]).toEqual(['2.54', '2', '2.5', '1.27']);
  expect(view.getAllByText('cm')).toHaveLength(4);
  expect(view.getByLabelText('Left').getAttribute('max')).toBe('25.4');
  fireEvent.change(view.getByLabelText('Left'), { target: { value: '3' } });
  fireEvent.click(view.getByRole('button', { name: 'Apply' }));
  expect(onApply).toHaveBeenCalledTimes(1);
  expect(onApply.mock.calls[0][0]).toMatchObject({ marginTop: 1440, marginBottom: 1134, marginLeft: 1701, marginRight: 720 });
});

test('keeps inches by default', () => {
  const onApply = mock((_props: Partial<SectionProperties>) => {});
  const view = render(<PageSetupDialog isOpen onClose={() => {}} onApply={onApply} currentProps={current} />);
  expect((view.getByLabelText('Top') as HTMLInputElement).value).toBe('1');
  expect(view.getAllByText('in')).toHaveLength(4);
  fireEvent.change(view.getByLabelText('Top'), { target: { value: '1.5' } });
  fireEvent.keyDown(view.getByLabelText('Top'), { key: 'Enter' });
  expect(onApply.mock.calls[0][0]).toMatchObject({ marginTop: 2160, marginLeft: 1418 });
});
