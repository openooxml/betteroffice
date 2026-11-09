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

let dispatching: Event | undefined;

function on(target: EventTarget, type: string, listener: (event: Event) => void) {
  target.addEventListener(type, (event) => {
    const previous = dispatching;
    dispatching = event;
    try {
      listener(event);
    } finally {
      dispatching = previous;
    }
  });
}

function setup(
  trusted?: (event: Event) => boolean,
  currentEvent: () => Event | undefined = () => dispatching
) {
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
  cleanups.push(() => ancestor.remove(), trackPluginGestures(root, trusted, currentEvent));
  return { ancestor, root, marker, button, id };
}

describe('plugin gestures', () => {
  for (const type of ['click', 'keydown', 'pointerup']) {
    test(`${type} stays open through ancestor bubble handlers and closes after dispatch`, async () => {
      const { ancestor, button, id } = setup(() => true);
      const seen: boolean[] = [];
      on(button, type, () => {
        seen.push(gestureFor(id));
      });
      on(ancestor, type, () => seen.push(gestureFor(id)));
      button.dispatchEvent(new Event(type, { bubbles: true }));
      expect(seen).toEqual([true, true]);
      expect(gestureFor(id)).toBe(false);
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(gestureFor(id)).toBe(false);
    });
  }

  test('synthetic events do not open a gesture', () => {
    const { button, id } = setup();
    on(button, 'click', () => expect(gestureFor(id)).toBe(false));
    button.click();
  });

  test('gestures belong to the marked activation, including across editor mounts', () => {
    const first = setup(() => true);
    const second = setup(() => true);
    on(first.button, 'click', () => {
      expect(gestureFor(first.id)).toBe(true);
      expect(gestureFor(second.id)).toBe(false);
    });
    first.button.click();
    on(second.button, 'click', () => {
      expect(gestureFor(first.id)).toBe(false);
      expect(gestureFor(second.id)).toBe(true);
    });
    second.button.click();
  });

  test('host chrome does not open a gesture', () => {
    const { root, id } = setup(() => true);
    on(root, 'click', () => expect(gestureFor(id)).toBe(false));
    root.click();
  });

  test('stopped propagation closes after dispatch and the fallback', async () => {
    const { button, id } = setup(() => true);
    on(button, 'click', (event) => {
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
    on(button, 'click', () => {
      expect(gestureFor(id)).toBe(true);
      button.dispatchEvent(new Event('keydown', { bubbles: true }));
    });
    on(button, 'keydown', () => expect(gestureFor(id)).toBe(false));
    button.click();
  });

  for (const type of ['input', 'plugin-action']) {
    test(`a nested synthetic ${type} event cannot borrow the trusted click`, () => {
      const { button, id } = setup(() => true);
      on(button, type, () => expect(gestureFor(id)).toBe(false));
      on(button, 'click', () => {
        expect(gestureFor(id)).toBe(true);
        button.dispatchEvent(type === 'input' ? new Event(type) : new CustomEvent(type));
        expect(gestureFor(id)).toBe(true);
      });
      button.click();
    });
  }

  test('a missing current event refuses even during a trusted click', () => {
    const { button, id } = setup(() => true, () => undefined);
    on(button, 'click', () => expect(gestureFor(id)).toBe(false));
    button.click();
  });

  test('nested events outside the editor cannot borrow an open gesture', () => {
    const { button, id } = setup(() => true);
    const hostButton = document.createElement('button');
    document.body.append(hostButton);
    cleanups.push(() => hostButton.remove());
    on(hostButton, 'click', () => expect(gestureFor(id)).toBe(false));
    on(button, 'click', () => {
      expect(gestureFor(id)).toBe(true);
      hostButton.click();
    });
    button.click();
  });

  test('the timeout clears the stopped event even if its phase and current event are retained', async () => {
    const event = new Event('click', { bubbles: true });
    const { button, id } = setup(() => true, () => event);
    on(button, 'click', (event) => event.stopPropagation());
    button.dispatchEvent(event);
    Object.defineProperty(event, 'eventPhase', { value: Event.BUBBLING_PHASE });
    expect(gestureFor(id)).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(gestureFor(id)).toBe(false);
  });
});
