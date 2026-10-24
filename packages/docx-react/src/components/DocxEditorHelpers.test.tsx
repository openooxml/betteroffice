import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { de } from '@betteroffice/docx-i18n';
import { LocaleProvider } from '../i18n';
import { DefaultLoadingIndicator, ParseError } from './DocxEditorHelpers';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { cleanup, render } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a document that cannot open shows a short explanation and keeps the technical message in details', () => {
  const view = render(<ParseError message="zip: central directory not found" />);
  const alert = view.getByRole('alert');
  expect(view.getByRole('heading', { name: "Couldn't open this document" })).toBeTruthy();
  expect(view.getByText("The file may be damaged, or it uses something this editor can't read yet.")).toBeTruthy();
  const details = alert.querySelector('details')!;
  expect(details.open).toBe(false);
  expect(details.querySelector('summary')!.textContent).toBe('Show details');
  expect(details.textContent).toContain('zip: central directory not found');
});

test('the open failure follows the editor locale', () => {
  const view = render(<LocaleProvider i18n={de}><ParseError message="x" /></LocaleProvider>);
  expect(view.getByRole('heading', { name: de.errors!.failedToLoad! })).toBeTruthy();
});

test('the default loading indicator is announced as a status', () => {
  const view = render(<DefaultLoadingIndicator />);
  expect(view.getByRole('status').textContent).toContain('Loading document...');
});
