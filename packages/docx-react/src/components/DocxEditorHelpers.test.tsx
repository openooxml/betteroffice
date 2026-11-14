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

test('the explanation says damaged or unsupported when the engine message tells them apart', () => {
  const description = (message: string) => {
    const view = render(<ParseError message={message} />);
    const text = view.getByRole('alert').textContent ?? '';
    view.unmount();
    return text;
  };
  for (const message of [
    'container error: bad zip: invalid Zip archive', 'container error: not a Word document: main document part is missing: word/document.xml',
    'container error: unsafe zip entry path: ../x', 'container error: duplicate normalized zip entry path: a', 'container error: read failed for a: truncated',
    'container error: word/media/image1.png inflates to other than its declared 123 bytes', 'unexpected EOF in word/styles.xml',
    'unexpected closing element in word/styles.xml', 'malformed XML in word/document.xml at byte 9: x',
    'invalid XML in word/document.xml: x', 'missing [Content_Types].xml', 'missing main document part: word/document.xml',
    'conflicting main document relationships: ["a"]',
  ]) expect(description(message)).toContain("The file is damaged or isn't a complete Word document.");
  for (const message of ['unsupported geometry projection', 'missing cell projection', 'missing cell paragraph', 'missing control projection']) {
    expect(description(message)).toContain("This document uses something this editor can't read yet.");
  }
  const either = "The file may be damaged, or it uses something this editor can't read yet.";
  for (const message of ['out of memory', 'container error: inflated size exceeds 536870912 bytes', 'container error: zip entry count exceeds 5000',
    'invalid [Content_Types].xml: XML depth limit exceeded in [Content_Types].xml', 'invalid _rels/.rels: package XML event limit exceeded in _rels/.rels',
    'invalid [Content_Types].xml: too many content type declarations']) {
    expect(description(message)).toContain(either);
  }
  const afterOpening = render(<ParseError opening={false} message="missing cell projection" />);
  expect(afterOpening.getByRole('alert').textContent).toContain(either);
});

test('the damaged and unsupported explanations follow the editor language', () => {
  const view = render(<LocaleProvider i18n={de}><ParseError message="bad zip: x" /></LocaleProvider>);
  expect(view.getByRole('alert').textContent).toContain(de.errors!.failedToLoadDamaged!);
  view.unmount();
  const next = render(<LocaleProvider i18n={de}><ParseError message="unsupported geometry projection" /></LocaleProvider>);
  expect(next.getByRole('alert').textContent).toContain(de.errors!.failedToLoadUnsupported!);
});
