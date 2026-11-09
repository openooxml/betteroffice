import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { gestureFor, pluginActivationId, trackPluginGestures } from './gesture';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0)) cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function setup(trusted?: (event: Event) => boolean) {
  const ancestor = document.createElement('div');
  const root = document.createElement('div');
  const marker = document.createElement('div');
  const button = document.createElement('button');
  const id = pluginActivationId({});
  marker.setAttribute('data-plugin-activation', id);
  marker.append(button);
  root.append(marker);
  ancestor.append(root);
  document.body.append(ancestor);
  cleanups.push(() => ancestor.remove(), trackPluginGestures(root, trusted));
  return { ancestor, root, marker, button, id };
}

describe('plugin gestures', () => {
  for (const type of ['click', 'keydown', 'pointerup']) {
    test(`${type} stays open through ancestor bubble handlers and closes after dispatch`, async () => {
      const { ancestor, button, id } = setup(() => true);
      const seen: boolean[] = [];
      button.addEventListener(type, () => {
        seen.push(gestureFor(id));
      });
      ancestor.addEventListener(type, () => seen.push(gestureFor(id)));
      button.dispatchEvent(new Event(type, { bubbles: true }));
      expect(seen).toEqual([true, true]);
      expect(gestureFor(id)).toBe(false);
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(gestureFor(id)).toBe(false);
    });
  }

  test('synthetic events do not open a gesture', () => {
    const { button, id } = setup();
    button.addEventListener('click', () => expect(gestureFor(id)).toBe(false));
    button.click();
  });

  test('gestures belong to the marked activation, including across editor mounts', () => {
    const first = setup(() => true);
    const second = setup(() => true);
    first.button.addEventListener('click', () => {
      expect(gestureFor(first.id)).toBe(true);
      expect(gestureFor(second.id)).toBe(false);
    });
    first.button.click();
    second.button.addEventListener('click', () => {
      expect(gestureFor(first.id)).toBe(false);
      expect(gestureFor(second.id)).toBe(true);
    });
    second.button.click();
  });

  test('host chrome does not open a gesture', () => {
    const { root, id } = setup(() => true);
    root.addEventListener('click', () => expect(gestureFor(id)).toBe(false));
    root.click();
  });

  test('stopped propagation closes after dispatch and the fallback', async () => {
    const { button, id } = setup(() => true);
    button.addEventListener('click', (event) => {
      event.stopPropagation();
      expect(gestureFor(id)).toBe(true);
    });
    button.click();
    expect(gestureFor(id)).toBe(false);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(gestureFor(id)).toBe(false);
  });

  test('nested synthetic events cannot borrow an open gesture', () => {
    const { button, id } = setup((event) => event.type === 'click');
    button.addEventListener('click', () => {
      expect(gestureFor(id)).toBe(true);
      button.dispatchEvent(new Event('keydown', { bubbles: true }));
    });
    button.addEventListener('keydown', () => expect(gestureFor(id)).toBe(false));
    button.click();
  });

  test('nested events outside the editor cannot borrow an open gesture', () => {
    const { button, id } = setup(() => true);
    const hostButton = document.createElement('button');
    document.body.append(hostButton);
    cleanups.push(() => hostButton.remove());
    hostButton.addEventListener('click', () => expect(gestureFor(id)).toBe(false));
    button.addEventListener('click', () => {
      expect(gestureFor(id)).toBe(true);
      hostButton.click();
    });
    button.click();
  });

  test('the timeout clears the stopped event even if its dispatch phase is retained', async () => {
    const { button, id } = setup(() => true);
    const event = new Event('click', { bubbles: true });
    button.addEventListener('click', (event) => event.stopPropagation());
    button.dispatchEvent(event);
    Object.defineProperty(event, 'eventPhase', { value: Event.BUBBLING_PHASE });
    expect(gestureFor(id)).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(gestureFor(id)).toBe(false);
  });
});
