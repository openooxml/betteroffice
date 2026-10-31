/** The error worker sessions raise where the runtime has no Web Worker (Node.js), naming the in-thread API. */
export function missingWebWorker(inThread: string): Error {
  return new Error(`Worker sessions need a Web Worker (browsers, Bun, Deno), which this runtime lacks; use ${inThread} in-thread or pass the worker option`);
}

export function deferSessionWorker(create: () => Promise<Worker>, inThread: string): Worker {
  const proc = (globalThis as { process?: { getBuiltinModule?: unknown } }).process;
  if (typeof proc?.getBuiltinModule !== 'function') throw missingWebWorker(inThread);
  let current: Worker | undefined;
  let terminated = false;
  let failure: unknown;
  const pending: Array<(worker: Worker) => void> = [];
  const errors = new Set<EventListenerOrEventListenerObject>();
  function send(action: (worker: Worker) => void): void {
    if (terminated || failure) return;
    if (current) action(current);
    else pending.push(action);
  }
  const worker = {
    onmessage: null,
    onerror: null,
    onmessageerror: null,
    postMessage(message: unknown, transfer?: Transferable[]) {
      if (failure) throw failure;
      send(worker => worker.postMessage(message, transfer ?? []));
    },
    terminate() {
      terminated = true;
      pending.length = 0;
      errors.clear();
      current?.terminate();
    },
    addEventListener(type: string, listener: EventListenerOrEventListenerObject | null) {
      if (!listener) return;
      if (type === 'error') errors.add(listener);
      send(worker => worker.addEventListener(type, listener));
    },
    removeEventListener(type: string, listener: EventListenerOrEventListenerObject | null) {
      if (!listener) return;
      if (type === 'error') errors.delete(listener);
      send(worker => worker.removeEventListener(type, listener));
    },
  } as unknown as Worker;
  void create().then(opened => {
    if (terminated) { opened.terminate(); return; }
    current = opened;
    opened.onmessage = event => worker.onmessage?.call(worker, event);
    opened.onerror = event => worker.onerror?.call(worker, event);
    opened.onmessageerror = event => worker.onmessageerror?.call(worker, event);
    for (const action of pending.splice(0)) action(opened);
  }).catch(error => {
    if (terminated) return;
    failure = error;
    pending.length = 0;
    current?.terminate();
    const event = { error, message: error instanceof Error ? error.message : String(error) } as ErrorEvent;
    for (const listener of [...errors]) {
      if (!errors.has(listener)) continue;
      if (typeof listener === 'function') listener.call(worker, event);
      else listener.handleEvent(event);
    }
    worker.onerror?.call(worker, event);
  });
  return worker;
}
