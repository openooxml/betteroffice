import { workerThreads } from './nodeSessionScope';

type Listener = EventListenerOrEventListenerObject;

export function createSessionWorker(url: URL, options: WorkerOptions, inThread: string): Worker {
  const { type: _, ...threadOptions } = options;
  const thread = new (workerThreads(inThread).Worker)(url, threadOptions);
  const listeners = new Map<string, Set<Listener>>();
  let terminated = false;
  const worker = {
    onmessage: null,
    onerror: null,
    onmessageerror: null,
    postMessage(message: unknown, transfer?: Transferable[]) { thread.postMessage(message, transfer); },
    terminate() {
      if (terminated) return;
      terminated = true;
      listeners.clear();
      thread.terminate();
    },
    addEventListener(type: string, listener: Listener | null) {
      if (!listener) return;
      let entries = listeners.get(type);
      if (!entries) listeners.set(type, entries = new Set());
      entries.add(listener);
    },
    removeEventListener(type: string, listener: Listener | null) {
      if (listener) listeners.get(type)?.delete(listener);
    },
  } as unknown as Worker;
  for (const type of ['message', 'error', 'messageerror'] as const) {
    thread.on(type, (value) => {
      if (terminated) return;
      const event = (type === 'message' ? { data: value } : {
        error: value, message: value instanceof Error ? value.message : String(value),
      }) as unknown as Event;
      for (const listener of [...(listeners.get(type) ?? [])]) {
        if (!listeners.get(type)?.has(listener)) continue;
        if (typeof listener === 'function') listener.call(worker, event);
        else listener.handleEvent(event);
      }
      const handler = worker[`on${type}`] as ((event: Event) => void) | null;
      if (!terminated) handler?.call(worker, event);
    });
  }
  return worker;
}
