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

function editing(onEdit: (text: string, href: string) => void) {
  const view = render(
    <LocaleProvider i18n={de}>
      <HyperlinkPopup data={data} onNavigate={() => {}} onCopy={() => {}} onEdit={onEdit} onRemove={() => {}} onClose={() => {}} />
    </LocaleProvider>
  );
  fireEvent.click(view.container.querySelectorAll('.oox-hyperlink-popup__icon-btn')[1]!);
  const [text, address] = view.getAllByRole('textbox') as HTMLInputElement[];
  return { text: text!, address: address! };
}

test('an address typed after a late frame lands in the address field and leaves the text', () => {
  const frames: FrameRequestCallback[] = [];
  const requestFrame = globalThis.requestAnimationFrame;
  globalThis.requestAnimationFrame = (callback) => frames.push(callback);
  try {
    const onEdit = mock((_text: string, _href: string) => {});
    const { address } = editing(onEdit);
    address.focus();
    for (const frame of frames.splice(0)) frame(0);
    const typedInto = document.activeElement as HTMLInputElement;
    fireEvent.change(typedInto, { target: { value: 'https://example.com/popup' } });
    fireEvent.keyDown(typedInto, { key: 'Enter' });
    expect(onEdit.mock.calls).toEqual([['Example', 'https://example.com/popup']]);
    expect(typedInto).toBe(address);
  } finally {
    globalThis.requestAnimationFrame = requestFrame;
  }
});

test('entering edit mode puts the caret into the text field with its content selected, before any frame', () => {
  const requestFrame = globalThis.requestAnimationFrame;
  globalThis.requestAnimationFrame = () => 0;
  try {
    const { text } = editing(() => {});
    expect(document.activeElement).toBe(text);
    expect([text.selectionStart, text.selectionEnd]).toEqual([0, 'Example'.length]);
  } finally {
    globalThis.requestAnimationFrame = requestFrame;
  }
});
