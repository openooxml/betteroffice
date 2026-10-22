import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { HyperlinkDialog } from './HyperlinkDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { createRoot } = await import('react-dom/client');
const { flushSync } = await import('react-dom');
const { act } = await import('react');

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

test.each([false, true])('the delayed URL focus leaves a field the user already focused, user focus=%s', (userFocus) => {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);
  const delayed: (() => void)[] = [];
  const setTimeoutOriginal = globalThis.setTimeout;
  const actEnvironment = globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean };
  const actEnvironmentOriginal = actEnvironment.IS_REACT_ACT_ENVIRONMENT;
  actEnvironment.IS_REACT_ACT_ENVIRONMENT = true;
  globalThis.setTimeout = ((callback: () => void, delay?: number) => {
    if (delay !== 100) return setTimeoutOriginal(callback, delay);
    delayed.push(callback);
    return 0;
  }) as typeof setTimeout;

  try {
    act(() => {
      root.render(<HyperlinkDialog isOpen onClose={() => {}} onSubmit={() => {}} selectedText="" />);
    });
    const display = container.querySelector<HTMLInputElement>('#hyperlink-display-text')!;
    const url = container.querySelector<HTMLInputElement>('#hyperlink-url')!;
    if (userFocus) display.focus();
    expect(delayed).toHaveLength(1);
    for (const callback of delayed) callback();
    expect(document.activeElement?.id).toBe(userFocus ? display.id : url.id);
  } finally {
    globalThis.setTimeout = setTimeoutOriginal;
    act(() => root.unmount());
    actEnvironment.IS_REACT_ACT_ENVIRONMENT = actEnvironmentOriginal;
    container.remove();
  }
});
