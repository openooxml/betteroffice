import { SessionFailure } from './types';

export interface SessionTransport {
  post(message: unknown, transfer?: Transferable[]): void;
  listen(listener: (message: unknown) => void): () => void;
  onError(listener: (error: unknown) => void): () => void;
  close(): void;
}

export interface SessionScope {
  postMessage(message: unknown, transfer: Transferable[]): void;
  addEventListener(type: string, listener: EventListener): void;
  removeEventListener(type: string, listener: EventListener): void;
  close?(): void;
}

function transportFor(scope: SessionScope, close: () => void): SessionTransport {
  const cleanups = new Set<() => void>();
  let closed = false;
  function subscribe(type: string, listener: EventListener): () => void {
    if (closed) return () => {};
    scope.addEventListener(type, listener);
    const remove = () => {
      scope.removeEventListener(type, listener);
      cleanups.delete(remove);
    };
    cleanups.add(remove);
    return remove;
  }
  return {
    post(message, transfer = []) {
      if (closed) throw new SessionFailure('disposed', 'Session transport is closed');
      scope.postMessage(message, transfer);
    },
    listen(listener) {
      return subscribe('message', (event) => listener((event as MessageEvent).data));
    },
    onError(listener) {
      const crash = subscribe('error', (event) => listener(new SessionFailure(
        'crash', (event as ErrorEvent).message || 'Session worker crashed'
      )));
      const message = subscribe('messageerror', () => listener(new SessionFailure(
        'message', 'Session transport received an unreadable message'
      )));
      return () => { crash(); message(); };
    },
    close() {
      if (closed) return;
      closed = true;
      for (const cleanup of [...cleanups]) cleanup();
      close();
    },
  };
}

/** Adapts an existing worker without constructing one. */
export function createWorkerTransport(worker: Worker): SessionTransport {
  return transportFor(worker, () => worker.terminate());
}

/** Adapts a dedicated worker scope. */
export function createScopeTransport(scope: SessionScope): SessionTransport {
  return transportFor(scope, () => scope.close?.());
}
