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
