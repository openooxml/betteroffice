import { hasOwn, isHostMessage } from './protocol';
import type { SessionTransport } from './transport';
import { SessionFailure, type Promisified, type SessionEvents, type SessionMethods } from './types';

export const OFFICE_SESSION_SILENCE_MS = 60_000;

export function requestWasmCompile(transport: SessionTransport): void {
  try {
    transport.post({ protocol: 1, kind: 'wasm-compile' });
  } catch (error) {
    try { transport.close(); } catch {}
    throw error;
  }
}

export interface SessionClientOptions<M extends SessionMethods> {
  methods: { readonly [K in keyof M]-?: true };
  silenceMs?: number;
  now?(): number;
  timer?(callback: () => void, ms: number): () => void;
  onWasmModule?(url: string, module: WebAssembly.Module): void;
}

export interface SessionClient<M extends SessionMethods, E extends SessionEvents> {
  readonly call: Promisified<M>;
  callWithTransfer<K extends keyof M & string>(
    method: K, args: Parameters<M[K]>, transfer: Transferable[]
  ): Promise<Awaited<ReturnType<M[K]>>>;
  on<K extends keyof E & string>(name: K, listener: (payload: E[K]) => void): () => void;
  onFailure(listener: (failure: SessionFailure) => void): () => void;
  readonly failure: SessionFailure | undefined;
  dispose(): Promise<void>;
}

/** Creates a typed client over an existing transport. */
export function createSessionClient<M extends SessionMethods, E extends SessionEvents>(
  transport: SessionTransport, options: SessionClientOptions<M>
): SessionClient<M, E> {
  const pending = new Map<number, { resolve(value: unknown): void; reject(error: unknown): void }>();
  const events = new Map<string, Set<(payload: unknown) => void>>();
  const failures = new Set<(failure: SessionFailure) => void>();
  const methods = new Set<string>(Object.keys(options.methods));
  const now = options.now ?? Date.now;
  const timer = options.timer ?? ((callback: () => void, ms: number) => {
    const id = setTimeout(callback, ms);
    return () => clearTimeout(id);
  });
  const silenceMs = options.silenceMs ?? OFFICE_SESSION_SILENCE_MS;
  let nextId = 1;
  let failure: SessionFailure | undefined;
  let cancelWatchdog: (() => void) | undefined;
  let disposal: Promise<void> | undefined;
  let unlisten = () => {};
  let unerror = () => {};

  function disarm(): void {
    cancelWatchdog?.();
    cancelWatchdog = undefined;
  }
  function notify<T>(listeners: Set<(value: T) => void>, value: T, event = false): void {
    for (const listener of [...listeners]) {
      if (!listeners.has(listener) || (event && failure)) continue;
      try { listener(value); } catch {}
    }
  }
  function end(error: SessionFailure, notifyFailure = true): void {
    if (failure) return;
    failure = error;
    disarm();
    unlisten();
    unerror();
    for (const request of pending.values()) request.reject(error);
    pending.clear();
    events.clear();
    if (notifyFailure) {
      try { transport.close(); } catch {}
      notify(failures, error);
    }
    failures.clear();
  }
  function arm(): void {
    disarm();
    if (pending.size === 0 || failure) return;
    const started = now();
    cancelWatchdog = timer(() => end(new SessionFailure(
      'silence', `Session received no message within ${silenceMs}ms`,
      `Waiting for ${pending.size} call(s); silent for ${now() - started}ms`
    )), silenceMs);
  }
  function request(method: string, args: unknown[], transfer: Transferable[] = []): Promise<unknown> {
    if (failure) return Promise.reject(failure);
    if (!methods.has(method)) return Promise.reject(new Error(`Unknown session method: ${method}`));
    const id = nextId++;
    return new Promise((resolve, reject) => {
      const first = pending.size === 0;
      pending.set(id, { resolve, reject });
      if (first) arm();
      try {
        transport.post({ protocol: 1, kind: 'call', id, method, args }, transfer);
      } catch (error) {
        pending.delete(id);
        if (pending.size === 0) disarm();
        reject(error instanceof Error ? error :
          new Error(`Session call could not be posted: ${String(error)}`));
      }
    });
  }
  unlisten = transport.listen((message) => {
    if (failure) return;
    arm();
    if (!isHostMessage(message)) {
      if (message !== null && typeof message === 'object' &&
        'kind' in message && message.kind === 'wasm-module') return;
      end(new SessionFailure('message', 'Session received a malformed host message'));
    } else if (message.kind === 'failure') {
      end(new SessionFailure(message.code, message.message, message.diagnostics));
    } else if (message.kind === 'wasm-module') {
      try { options.onWasmModule?.(message.url, message.module); } catch {}
    } else if (message.kind === 'event') {
      const listeners = events.get(message.name);
      if (listeners) notify(listeners, message.payload, true);
    } else {
      const request = pending.get(message.id);
      if (!request) return;
      pending.delete(message.id);
      if (pending.size === 0) disarm();
      if (message.ok) request.resolve(message.value);
      else {
        const error = new Error(message.error.message);
        error.name = message.error.name;
        if (hasOwn(message.error, 'refusal')) {
          Object.assign(error, { refusal: message.error.refusal });
        }
        request.reject(error);
      }
    }
  });
  unerror = transport.onError((error) => end(error instanceof SessionFailure ? error :
    new SessionFailure('crash', error instanceof Error ? error.message : String(error))));
  const call = Object.create(null) as Promisified<M>;
  for (const method of methods) {
    Object.defineProperty(call, method, {
      enumerable: true, value: (...args: unknown[]) => request(method, args),
    });
  }
  return {
    call,
    callWithTransfer: (method, args, transfer) =>
      request(method, args, transfer) as Promise<Awaited<ReturnType<M[typeof method]>>>,
    on(name, listener) {
      if (failure) return () => {};
      const listeners = events.get(name) ?? new Set<(payload: unknown) => void>();
      events.set(name, listeners);
      const callback = listener as (payload: unknown) => void;
      listeners.add(callback);
      return () => { listeners.delete(callback); };
    },
    onFailure(listener) {
      if (!failure) failures.add(listener);
      return () => { failures.delete(listener); };
    },
    get failure() { return failure; },
    dispose() {
      if (disposal) return disposal;
      const ended = failure !== undefined;
      end(new SessionFailure('disposed', 'Session was disposed'), false);
      disposal = new Promise<void>((resolve) => {
        if (!ended) {
          try { transport.post({ protocol: 1, kind: 'dispose' }); } catch {}
        }
        setTimeout(() => {
          try { transport.close(); } catch {}
          resolve();
        }, 0);
      });
      return disposal;
    },
  };
}
