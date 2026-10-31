import type { SessionScope } from './transport';
import { missingWebWorker } from './webWorker';

type WorkerEvent = 'message' | 'error' | 'messageerror';
type Listener = EventListenerOrEventListenerObject;

interface ThreadWorker {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  terminate(): unknown;
  on(type: WorkerEvent, listener: (value: unknown) => void): void;
}

interface WorkerThreads {
  Worker: new (url: URL, options: { name?: string; credentials?: RequestCredentials }) => ThreadWorker;
  parentPort: SessionScope | null;
}

function workerThreads(inThread: string): WorkerThreads {
  const proc = (globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } }).process;
  if (typeof proc?.getBuiltinModule !== 'function') throw missingWebWorker(inThread);
  const threads = proc.getBuiltinModule('node:worker_threads') as WorkerThreads | undefined;
  if (!threads) throw missingWebWorker(inThread);
  return threads;
}

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

export function sessionWorkerScope(): SessionScope {
  const scope = workerThreads('the in-thread API').parentPort;
  if (!scope) throw missingWebWorker('the in-thread API');
  return scope;
}
