export interface WorkerBootScope {
  postMessage(message: unknown): void;
  addEventListener(type: string, listener: EventListener): void;
  removeEventListener(type: string, listener: EventListener): void;
  dispatchEvent(event: Event): boolean;
}

const HELD_EVENTS = ['message', 'messageerror'];

/** Holds messages while `load` imports the worker, starts it, then delivers each held message once, in order. */
export function bootWorker<Scope extends WorkerBootScope>(
  scope: Scope,
  load: () => Promise<(scope: Scope) => void>,
  failure: (message: string) => unknown,
  guard: string,
): void {
  const flags = scope as unknown as Record<symbol, unknown>;
  const key = Symbol.for(guard);
  if (flags[key]) return;
  flags[key] = true;
  let queue: MessageEvent[] | null = [];
  const replaying = new WeakSet<Event>();
  const hold: EventListener = (event) => {
    if (replaying.has(event)) return;
    event.stopImmediatePropagation();
    queue?.push(event as MessageEvent);
  };
  for (const type of HELD_EVENTS) scope.addEventListener(type, hold);
  load().then((start) => {
    start(scope);
    for (let event = queue?.shift(); event; event = queue?.shift()) {
      const replay = new MessageEvent(event.type, {
        data: event.data, ports: [...event.ports], origin: event.origin, lastEventId: event.lastEventId,
      });
      replaying.add(replay);
      scope.dispatchEvent(replay);
    }
    queue = null;
    for (const type of HELD_EVENTS) scope.removeEventListener(type, hold);
  }).catch((error: unknown) => {
    queue = null;
    scope.postMessage(failure(error instanceof Error ? error.message : String(error)));
    setTimeout(() => { throw error; });
  });
}
