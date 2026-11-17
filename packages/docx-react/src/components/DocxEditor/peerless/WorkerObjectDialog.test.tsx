import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { WorkerObjectDialog } from './WorkerObjectDialog';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { createRoot } = await import('react-dom/client');
const { flushSync } = await import('react-dom');

afterAll(async () => {
  await new Promise((resolve) => setTimeout(resolve, 10));
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('the column width dialog lays out styled Cancel and Apply buttons apart', () => {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);
  try {
    flushSync(() => {
      root.render(
        <WorkerObjectDialog
          action="resizeColumn"
          target={{ handle: 'h', version: '1', x: 0, y: 0, width: 24, height: 10 }}
          apply={async () => true}
          close={() => {}}
        />
      );
    });
    const [cancel, apply] = Array.from(container.querySelectorAll('button'));
    const footer = cancel.parentElement!;
    expect(footer).toBe(apply.parentElement!);
    expect(footer.style.display).toBe('flex');
    expect(footer.style.gap).toBe('8px');
    for (const button of [cancel, apply]) {
      expect(button.style.padding).toBe('6px 16px');
      expect(button.style.borderRadius).toBe('4px');
    }
    expect(apply.style.backgroundColor).toBe('var(--doc-primary)');
    expect(apply.style.color).toBe('var(--doc-on-primary)');
  } finally {
    root.unmount();
    container.remove();
  }
});
