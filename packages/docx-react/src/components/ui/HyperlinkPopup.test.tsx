import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { de } from '@betteroffice/docx-i18n';
import { Toaster } from 'sonner';
import { LocaleProvider } from '../../i18n';
import { HyperlinkPopup } from './HyperlinkPopup';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const data = { href: 'https://example.com/', displayText: 'Example', position: { top: 10, left: 10 } };

test('copying a link announces it in the editor locale', async () => {
  const onCopy = mock((_href: string) => {});
  const view = render(
    <LocaleProvider i18n={de}>
      <HyperlinkPopup data={data} onNavigate={() => {}} onCopy={onCopy} onEdit={() => {}} onRemove={() => {}} onClose={() => {}} />
      <Toaster />
    </LocaleProvider>
  );
  fireEvent.click(view.getByTitle(de.hyperlinkPopup!.copyLink!));
  expect(onCopy).toHaveBeenCalledWith('https://example.com/');
  await waitFor(() => expect(view.getByText(de.editor!.linkCopied!)).toBeTruthy());
});

test('a link with a ScreenTip shows it in the popup and on hover (7171990133)', () => {
  const view = render(
    <HyperlinkPopup data={{ ...data, tooltip: 'Open the example' }} onNavigate={() => {}} onCopy={() => {}} onEdit={() => {}}
      onRemove={() => {}} onClose={() => {}} />
  );
  expect(view.getByText('Open the example')).toBeTruthy();
  expect(view.getByText('https://example.com/').getAttribute('title')).toBe('Open the example');
});

test('a link without a ScreenTip keeps its address as the hover text', () => {
  const view = render(
    <HyperlinkPopup data={data} onNavigate={() => {}} onCopy={() => {}} onEdit={() => {}} onRemove={() => {}} onClose={() => {}} />
  );
  expect(view.getByText('https://example.com/').getAttribute('title')).toBe('https://example.com/');
  expect(view.container.querySelector('.oox-hyperlink-popup__tooltip')).toBeNull();
});
