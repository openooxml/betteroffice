import { missingWebWorker } from './webWorker';

type Listener = EventListenerOrEventListenerObject;
type Handler = ((event: Event) => void) | null;

export function createDeferredSessionWorker(load: () => Promise<Worker>, inThread: string): Worker {
  const proc = (globalThis as { process?: { getBuiltinModule?: unknown } }).process;
  if (typeof proc?.getBuiltinModule !== 'function') throw missingWebWorker(inThread);
  const listeners = new Map<string, Set<Listener>>();
  const queue: Array<[unknown, Transferable[] | undefined]> = [];
  let target: Worker | undefined;
  let terminated = false;
  const worker = {
    onmessage: null,
    onerror: null,
    onmessageerror: null,
    postMessage(message: unknown, transfer?: Transferable[]) {
      if (terminated) return;
      if (target) target.postMessage(message, transfer as Transferable[]);
      else queue.push([message, transfer]);
    },
    terminate() {
      if (terminated) return;
      terminated = true;
      listeners.clear();
      queue.length = 0;
      target?.terminate();
    },
    addEventListener(type: string, listener: Listener | null) {
      if (!listener || terminated) return;
      let entries = listeners.get(type);
      if (!entries) listeners.set(type, entries = new Set());
      entries.add(listener);
    },
    removeEventListener(type: string, listener: Listener | null) {
      if (listener) listeners.get(type)?.delete(listener);
    },
  };
  const facade = worker as unknown as Worker;
  const dispatch = (type: 'message' | 'error' | 'messageerror', event: Event) => {
    if (terminated) return;
    for (const listener of [...(listeners.get(type) ?? [])]) {
      if (!listeners.get(type)?.has(listener)) continue;
      if (typeof listener === 'function') listener.call(facade, event);
      else listener.handleEvent(event);
    }
    const handler = (worker as unknown as Record<string, Handler>)[`on${type}`];
    if (!terminated) handler?.call(facade, event);
  };
  load().then((loaded) => {
    if (terminated) {
      loaded.terminate();
      return;
    }
    target = loaded;
    for (const type of ['message', 'error', 'messageerror'] as const) {
      loaded.addEventListener(type, (event) => dispatch(type, event));
    }
    for (const [message, transfer] of queue.splice(0)) loaded.postMessage(message, transfer as Transferable[]);
  }, (error: unknown) => {
    dispatch('error', { error, message: error instanceof Error ? error.message : String(error) } as unknown as Event);
  });
  return facade;
}
