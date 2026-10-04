import {
  hasOwn, isClientMessage, isDeferredReply, isTransferResult,
  type ClientMessage, type HostMessage, type ReplyError,
} from './protocol';
import { createResidentScheduler } from './scheduler';
import type { SessionTransport } from './transport';
import {
  SESSION_SUPERSEDED, SessionFailure, type MethodHandlers, type MethodPolicies,
  type SessionEvents, type SessionMethods, type SessionScheduler,
} from './types';

export interface SessionHostOptions<M extends SessionMethods, C> {
  handlers: MethodHandlers<M, C>;
  policies: MethodPolicies<M>;
  context: C;
  onDispose?(): void;
}

export interface SessionHost<E extends SessionEvents> {
  readonly scheduler: SessionScheduler;
  emit<K extends keyof E & string>(name: K, payload: E[K], transfer?: Transferable[]): void;
}

function replyError(error: unknown): ReplyError {
  const result: ReplyError = {
    name: error instanceof Error ? error.name : 'Error',
    message: error instanceof Error ? error.message : String(error),
  };
  if (error !== null && typeof error === 'object' && 'refusal' in error) result.refusal = error.refusal;
  return result;
}

function terminal(error: unknown): boolean {
  return error instanceof SessionFailure ||
    (typeof WebAssembly !== 'undefined' && error instanceof WebAssembly.RuntimeError);
}

function classify(error: unknown): SessionFailure {
  if (error instanceof SessionFailure) return error;
  return new SessionFailure(terminal(error) ? 'trap' : 'crash', replyError(error).message);
}

/** Runs method handlers serially through the resident scheduler. */
export function createSessionHost<M extends SessionMethods, E extends SessionEvents, C>(
  transport: SessionTransport, options: SessionHostOptions<M, C>
): SessionHost<E> {
  let failure: SessionFailure | undefined;
  let disposed = false;
  let unlisten = () => {};
  let unerror = () => {};
  const timeouts = new Set<ReturnType<typeof setTimeout>>();
  const tasks = new Set<{ cancel(reason?: unknown): void }>();
  const ended = () => disposed || failure !== undefined;
  function cleanup(): void {
    unlisten();
    unerror();
    for (const id of timeouts) clearTimeout(id);
    for (const task of tasks) {
      try { task.cancel(); } catch {}
    }
    timeouts.clear();
    tasks.clear();
    try { options.onDispose?.(); } catch {}
  }
  function fail(error: SessionFailure): void {
    if (ended()) return;
    failure = error;
    try {
      transport.post({
        protocol: 1, kind: 'failure', code: error.code, message: error.message,
        ...(error.diagnostics === undefined ? {} : { diagnostics: error.diagnostics }),
      });
    } catch {}
    cleanup();
  }
  function send(message: HostMessage, transfer?: Transferable[]): void {
    if (ended()) return;
    try { transport.post(message, transfer); } catch (error) {
      fail(new SessionFailure('message', `Session could not post a message: ${replyError(error).message}`));
    }
  }
  function timeout(callback: () => void, ms: number): () => void {
    if (ended()) return () => {};
    const id = setTimeout(() => {
      timeouts.delete(id);
      if (!ended()) callback();
    }, ms);
    timeouts.add(id);
    return () => { clearTimeout(id); timeouts.delete(id); };
  }
  function track(cancel: (reason?: unknown) => void) {
    let ended = false;
    const task = {
      end() {
        ended = true;
        tasks.delete(task);
      },
      cancel(reason?: unknown) {
        if (ended) return;
        task.end();
        cancel(reason);
      },
    };
    tasks.add(task);
    return task;
  }
  const resident = createResidentScheduler({
    now: Date.now,
    turn: (callback) => { timeout(callback, 0); },
    timer: timeout,
    failed: (error) => fail(classify(error)),
  });
  const scheduler: SessionScheduler = {
    schedule(task) {
      const tracked = track(() => task.cancel?.());
      if (ended()) {
        try { tracked.cancel(); } catch {}
        return;
      }
      resident.schedule({
        kind: task.kind,
        idleAfterInputMs: task.idleAfterInputMs,
        get version() { return task.version; },
        set version(value) { task.version = value; },
        get generation() { return task.generation; },
        set generation(value) { task.generation = value; },
        onStale: task.onStale ? () => task.onStale!() : undefined,
        cancel: () => tracked.cancel(),
        fail: task.fail ? (error) => {
          if (terminal(error)) {
            fail(classify(error));
            return;
          }
          tracked.end();
          task.fail!(error);
        } : undefined,
        async run(budgetMs) {
          if (ended()) {
            tracked.cancel();
            return 'done';
          }
          try {
            const step = await task.run(budgetMs);
            if (ended()) {
              tracked.cancel();
              return 'done';
            }
            if (step === 'done') tracked.end();
            return step;
          } catch (error) {
            if (terminal(error)) {
              fail(classify(error));
              return 'done';
            }
            tracked.end();
            throw error;
          }
        },
      });
    },
    dispatch(task) {
      const tracked = track((reason) => {
        try { task.cancel?.(reason); } catch (error) {
          if (terminal(error)) {
            fail(classify(error));
            return;
          }
          throw error;
        }
      });
      if (ended()) {
        try { tracked.cancel(); } catch {}
        return;
      }
      resident.dispatch({
        ...task,
        cancel: (reason) => tracked.cancel(reason),
        compute(input) {
          if (ended()) {
            tracked.cancel();
            throw failure ?? new SessionFailure('disposed', 'Session was disposed');
          }
          try {
            return task.compute(input);
          } catch (error) {
            if (terminal(error)) fail(classify(error));
            throw error;
          }
        },
        async install(result) {
          if (ended()) {
            tracked.cancel();
            return;
          }
          try {
            await task.install(result);
            tracked.end();
          } catch (error) {
            if (terminal(error)) fail(classify(error));
            throw error;
          }
        },
      });
    },
    bump: (changes) => resident.bump(changes),
    budget: () => resident.budget(),
    pending: () => resident.pending(),
    get version() { return resident.version; },
    get generation() { return resident.generation; },
  };
  function resolve(id: number, value: unknown): void {
    if (isTransferResult(value)) {
      send({ protocol: 1, kind: 'reply', id, ok: true, value: value.value }, value.transfer);
    } else send({ protocol: 1, kind: 'reply', id, ok: true, value });
  }
  function reject(id: number, error: unknown): void {
    if (terminal(error)) fail(classify(error));
    else send({ protocol: 1, kind: 'reply', id, ok: false, error: replyError(error) });
  }
  async function run(message: Extract<ClientMessage, { kind: 'call' }>): Promise<void> {
    if (ended()) return;
    try {
      const handler = options.handlers[message.method as keyof M] as
        (context: C, ...args: unknown[]) => unknown;
      const value = await handler(options.context, ...message.args);
      if (isDeferredReply(value)) {
        void value.promise.then(
          (value) => resolve(message.id, value), (error) => reject(message.id, error)
        );
      } else resolve(message.id, value);
    } catch (error) {
      reject(message.id, error);
    }
  }
  unlisten = transport.listen((message) => {
    if (ended()) return;
    if (!isClientMessage(message)) {
      fail(new SessionFailure('message', 'Session received a malformed client message'));
    } else if (message.kind === 'dispose') {
      disposed = true;
      cleanup();
      transport.close();
    } else if (!failure) {
      const policy = hasOwn(options.policies, message.method)
        ? options.policies[message.method as keyof M] : undefined;
      if (!hasOwn(options.handlers, message.method) || !policy) {
        send({ protocol: 1, kind: 'reply', id: message.id, ok: false,
          error: { name: 'Error', message: `Unknown session method: ${message.method}` } });
        return;
      }
      resident.submit({ ...policy, run: () => run(message), supersede: () => send({
        protocol: 1, kind: 'reply', id: message.id, ok: false,
        error: { name: SESSION_SUPERSEDED, message: 'Superseded by a newer request' },
      }) });
    }
  });
  unerror = transport.onError((error) => fail(classify(error)));
  return {
    scheduler,
    emit: (name, payload, transfer) => send({ protocol: 1, kind: 'event', name, payload }, transfer),
  };
}
