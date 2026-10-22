import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { HyperlinkDialog } from './HyperlinkDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { createRoot } = await import('react-dom/client');
const { flushSync } = await import('react-dom');

afterAll(async () => {
  await new Promise((resolve) => setTimeout(resolve, 10));
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('initializes display text from the selection before the first paint', () => {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);

  try {
    flushSync(() => {
      root.render(
        <HyperlinkDialog isOpen onClose={() => {}} onSubmit={() => {}} selectedText="Bare text" />
      );
    });

    expect(container.querySelector<HTMLInputElement>('#hyperlink-display-text')?.value).toBe(
      'Bare text'
    );
  } finally {
    root.unmount();
    container.remove();
  }
});

test('initializes URL and display text from existing link data before the first paint', () => {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);

  try {
    flushSync(() => {
      root.render(
        <HyperlinkDialog
          isOpen
          onClose={() => {}}
          onSubmit={() => {}}
          initialData={{ url: 'https://example.com/a', displayText: 'Edit me', tooltip: 'Tip' }}
        />
      );
    });

    expect(container.querySelector<HTMLInputElement>('#hyperlink-url')?.value).toBe(
      'https://example.com/a'
    );
    expect(container.querySelector<HTMLInputElement>('#hyperlink-display-text')?.value).toBe('Edit me');
  } finally {
    root.unmount();
    container.remove();
  }
});
