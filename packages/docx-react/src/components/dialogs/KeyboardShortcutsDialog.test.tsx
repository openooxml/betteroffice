import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { LocaleProvider } from '../../i18n';
import { KeyboardShortcutsDialog } from './KeyboardShortcutsDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

test('search finds a shortcut by its localized key names', () => {
  const view = render(
    <LocaleProvider i18n={{ _lang: 'de', dialogs: { keyboardShortcuts: { keys: { ctrl: 'Strg' } } } }}>
      <KeyboardShortcutsDialog
        isOpen
        onClose={() => {}}
        customShortcuts={[{ id: 'probe', name: 'Probe', description: 'Probe shortcut', keys: 'Ctrl+J', category: 'other' }]}
      />
    </LocaleProvider>
  );
  const search = view.container.ownerDocument.querySelector('input') as HTMLInputElement;
  fireEvent.change(search, { target: { value: 'strg+j' } });
  expect(view.getByText('Probe shortcut')).toBeTruthy();
  fireEvent.change(search, { target: { value: 'ctrl+j' } });
  expect(view.getByText('Probe shortcut')).toBeTruthy();
});
