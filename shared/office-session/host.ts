import {
  hasOwn, isClientMessage, isDeferredReply, isTransferResult,
  type ClientMessage, type HostMessage, type ReplyError,
} from './protocol';
import { createResidentScheduler, type TaskExecutor } from './scheduler';
import type { SessionTransport } from './transport';
import {
  SESSION_SUPERSEDED, SessionFailure, type MethodHandlers, type MethodPolicies, type MethodPolicy,
  type SessionEvents, type SessionMethods, type SessionScheduler,
} from './types';

export interface SessionHostOptions<M extends SessionMethods, C> {
  handlers: MethodHandlers<M, C>;
  policies: MethodPolicies<M>;
  context: C;
  executor?: TaskExecutor;
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
  const tasks = new Set<{ cancel(): void }>();
  const ended = () => disposed || failure !== undefined;
  function cleanup(): void {
    unlisten();
    unerror();
    for (const id of timeouts) clearTimeout(id);
    timeouts.clear();
    for (const task of tasks) {
      try { task.cancel(); } catch {}
    }
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
  function track(cancel?: () => void) {
    const task = {
      end() {
        tasks.delete(task);
      },
      cancel() {
        task.end();
        cancel?.();
      },
    };
    if (cancel) tasks.add(task);
    return task;
  }
  function invoke<T>(hook: () => T, fallback: T): T {
    if (ended()) return fallback;
    const reject = (error: unknown): T => {
      if (ended()) return fallback;
      if (!terminal(error)) throw error;
      fail(classify(error));
      return fallback;
    };
    try {
      const value = hook();
      return value instanceof Promise ? value.catch(reject) as T : value;
    } catch (error) {
      return reject(error);
    }
  }
  function settled<T>(promise: Promise<T>): Promise<T> {
    const guarded = promise.then();
    const guard = <A>(hook: (value: A) => unknown) => (value: A) => {
      if (ended()) return undefined;
      try { return hook(value); } catch (error) {
        fail(classify(error));
        return undefined;
      }
    };
    guarded.then = ((resolved?: (value: T) => unknown, rejected?: (error: unknown) => unknown) =>
      Promise.prototype.then.call(guarded, resolved && guard(resolved), rejected && guard(rejected))
    ) as typeof guarded.then;
    return guarded;
  }
  const executor = options.executor;
  const resident = createResidentScheduler({
    now: Date.now,
    turn: (callback) => { timeout(callback, 0); },
    timer: timeout,
    executor: executor && {
      run<Input, Result>(job: { kind: string; input: Input; transfer?: Transferable[] }): Promise<Result> {
        return settled(executor.run<Input, Result>(job));
      },
    },
    failed: (error) => fail(classify(error)),
  });
  const scheduler: SessionScheduler = {
    schedule(task) {
      if (ended()) {
        try { task.cancel?.(); } catch {}
        return;
      }
      const tracked = track(task.cancel && (() => task.cancel!()));
      resident.schedule({
        kind: task.kind,
        idleAfterInputMs: task.idleAfterInputMs,
        get version() { return task.version; },
        set version(value) { task.version = value; },
        get generation() { return task.generation; },
        set generation(value) { task.generation = value; },
        ...(task.onStale ? { onStale: () => invoke(() => task.onStale!(), 'cancel') } : {}),
        ...(task.cancel ? { cancel() {
          tracked.end();
          return invoke(() => task.cancel!(), undefined);
        } } : {}),
        ...(task.fail ? { fail(error: unknown) {
          tracked.end();
          return invoke(() => task.fail!(error), undefined);
        } } : {}),
        async run(budgetMs) {
          if (ended()) return 'done';
          try {
            const step = await task.run(budgetMs);
            if (step === 'done') tracked.end();
            return ended() ? 'done' : step;
          } catch (error) {
            tracked.end();
            if (ended()) return 'done';
            if (terminal(error)) {
              fail(classify(error));
              return 'done';
            }
            throw error;
          }
        },
      });
    },
    dispatch(task) {
      if (ended()) {
        try { task.cancel?.(); } catch {}
        return;
      }
      const tracked = track(task.cancel && (() => task.cancel!()));
      resident.dispatch({
        kind: task.kind,
        input: task.input,
        transfer: task.transfer,
        get version() { return task.version; },
        set version(value) { task.version = value; },
        get generation() { return task.generation; },
        set generation(value) { task.generation = value; },
        ...(task.cancel ? { cancel(reason?: unknown) {
          tracked.end();
          return invoke(() => task.cancel!(reason), undefined);
        } } : {}),
        compute(input) {
          return invoke(() => task.compute(input), undefined as ReturnType<typeof task.compute>);
        },
        async install(result) {
          if (ended()) return;
          try {
            await task.install(result);
          } catch (error) {
            tracked.end();
            if (ended()) return;
            if (terminal(error)) {
              fail(classify(error));
              return;
            }
            throw error;
          } finally {
            tracked.end();
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
    } else if (message.kind === 'call' && !failure) {
      const configured = hasOwn(options.policies, message.method)
        ? options.policies[message.method as keyof M] : undefined;
      if (!hasOwn(options.handlers, message.method) || !configured) {
        send({ protocol: 1, kind: 'reply', id: message.id, ok: false,
          error: { name: 'Error', message: `Unknown session method: ${message.method}` } });
        return;
      }
      let policy: MethodPolicy;
      try {
        policy = typeof configured === 'function'
          ? (configured as (...args: unknown[]) => MethodPolicy)(...message.args) : configured;
      } catch (error) {
        reject(message.id, error);
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
