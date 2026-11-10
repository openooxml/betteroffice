import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { gestureFor, pluginActivationId, trackPluginGestures } from './gesture';
import { createPluginRuntime, type PluginInvocation, type RuntimePlugin } from './runtime';

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
  for (const type of ['focus-change', 'viewport-change', 'proposal-change', 'locale-change']) {
    test(`${type} admits owner handler, microtasks, onEvent, run and nested notify only inside its window`, async () => {
      type Snapshot = { generation: string; version: string };
      type Invocation = PluginInvocation<Snapshot>;
      type Notification = { type: string; generation: string };
      type Plugin = RuntimePlugin<Invocation, Notification>;
      const seen: boolean[] = [];
      const own: boolean[] = [];
      let id = '';
      const runtime = createPluginRuntime<Plugin, Invocation, Notification, Snapshot, 'command'>({
        context: (invocation) => invocation,
        snapshot: (_pluginId, generation) => ({ generation, version: 'v1' }),
        currentVersion: () => 'v1',
        loadEvent: (snapshot) => ({ type: 'load', generation: snapshot.generation }),
        grantsEvent: (snapshot) => ({ type: 'grants-change', generation: snapshot.generation }),
        report: (error) => { throw error.error; },
      });
      runtime.setPlugins([{
        id: 'gesture', createState: () => null,
        async onEvent(invocation, event) {
          if (event.type !== type && event.type !== `${type}-nested`) return;
          if (event.type === type) runtime.notify({ type: `${type}-nested`, generation: 'g1' });
          seen.push(gestureFor(pluginActivationId(invocation.activation)));
          await Promise.resolve();
          seen.push(gestureFor(pluginActivationId(invocation.activation)));
          await invocation.run((next) => { seen.push(gestureFor(pluginActivationId(next.activation))); });
        },
      }]);
      runtime.open('g1');
      await runtime.notificationsIdle();
      const entry = runtime.activations()[0]!;
      const event = new Event('click', { bubbles: true });
      const ui = setup(() => true, () => event);
      id = pluginActivationId(entry.activation);
      ui.marker.setAttribute('data-plugin-activation', id);
      on(ui.button, 'click', (event) => {
        event.stopPropagation();
        own.push(gestureFor(pluginActivationId(entry.context.activation)));
        queueMicrotask(() => own.push(gestureFor(pluginActivationId(entry.context.activation))));
        runtime.notify({ type, generation: 'g1' });
      });
      try {
        ui.button.dispatchEvent(event);
        Object.defineProperty(event, 'eventPhase', { configurable: true, value: Event.BUBBLING_PHASE });
        await runtime.notificationsIdle();
        expect(seen).toEqual([true, true, true, true, true, true]);
        expect(own).toEqual([true, true]);
        Object.defineProperty(event, 'eventPhase', { value: Event.NONE });
        runtime.notify({ type, generation: 'g1' });
        await runtime.notificationsIdle();
        expect(seen.slice(6)).toEqual([false, false, false, false, false, false]);
      } finally {
        runtime.close('unmounted');
      }
    });
  }

  for (const inside of [true, false]) {
    test(`load delivery and its nested notify are ${inside ? 'admitted inside the owner window' : 'refused outside a gesture'}`, async () => {
      type Snapshot = { generation: string; version: string };
      type Invocation = PluginInvocation<Snapshot>;
      type Notification = { type: string; generation: string };
      type Plugin = RuntimePlugin<Invocation, Notification>;
      let resume!: () => void;
      const gate = new Promise<void>((resolve) => { resume = resolve; });
      let activation!: object;
      const seen: Array<{ type: string; admitted: boolean }> = [];
      const runtime = createPluginRuntime<Plugin, Invocation, Notification, Snapshot, 'command'>({
        context: (invocation) => invocation,
        snapshot: (_pluginId, generation) => ({ generation, version: 'v1' }),
        currentVersion: () => 'v1',
        loadEvent: (snapshot) => ({ type: 'load', generation: snapshot.generation }),
        grantsEvent: (snapshot) => ({ type: 'grants-change', generation: snapshot.generation }),
        report: (error) => { throw error.error; },
      });
      runtime.setPlugins([{
        id: 'gesture', createState: () => null,
        initialize(invocation) {
          activation = invocation.activation;
          return gate;
        },
        async onEvent(invocation, notification) {
          const admitted = () => gestureFor(pluginActivationId(invocation.activation));
          seen.push({ type: notification.type, admitted: admitted() });
          if (notification.type === 'load') {
            await Promise.resolve();
            seen.push({ type: 'load-microtask', admitted: admitted() });
            runtime.notify({ type: 'nested', generation: 'g1' });
          }
        },
      }]);
      runtime.open('g1');
      const event = new Event('click', { bubbles: true });
      const ui = setup(() => true, () => event);
      ui.marker.setAttribute('data-plugin-activation', pluginActivationId(activation));
      on(ui.button, 'click', (event) => {
        event.stopPropagation();
        resume();
      });
      try {
        if (inside) {
          ui.button.dispatchEvent(event);
          Object.defineProperty(event, 'eventPhase', { value: Event.BUBBLING_PHASE });
        } else {
          resume();
        }
        await runtime.notificationsIdle();
        expect(seen).toEqual([
          { type: 'load', admitted: inside },
          { type: 'load-microtask', admitted: inside },
          { type: 'nested', admitted: inside },
        ]);
      } finally {
        resume();
        runtime.close('unmounted');
      }
    });
  }

  test('an await past dispatch does not qualify for the handler invocation', async () => {
    const { button, id } = setup(() => true);
    let resume!: () => void;
    const gate = new Promise<void>((resolve) => { resume = resolve; });
    let result!: Promise<boolean>;
    on(button, 'click', () => {
      expect(gestureFor(id)).toBe(true);
      result = gate.then(() => gestureFor(id));
    });
    button.click();
    resume();
    expect(await result).toBe(false);
  });

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
