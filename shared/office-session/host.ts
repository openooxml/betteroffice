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

/** Runs method handlers serially through the resident scheduler. */
export function createSessionHost<M extends SessionMethods, E extends SessionEvents, C>(
  transport: SessionTransport, options: SessionHostOptions<M, C>
): SessionHost<E> {
  let failure: SessionFailure | undefined;
  let disposed = false;
  let unlisten = () => {};
  let unerror = () => {};
  const ended = () => disposed || failure !== undefined;
  function cleanup(): void {
    unlisten();
    unerror();
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
  const scheduler = createResidentScheduler({
    now: Date.now,
    turn: (callback) => { setTimeout(() => { if (!ended()) callback(); }, 0); },
    timer: (callback, ms) => {
      const id = setTimeout(() => { if (!ended()) callback(); }, ms);
      return () => clearTimeout(id);
    },
    failed: (error) => fail(new SessionFailure('crash', replyError(error).message)),
  });
  function resolve(id: number, value: unknown): void {
    if (isTransferResult(value)) {
      send({ protocol: 1, kind: 'reply', id, ok: true, value: value.value }, value.transfer);
    } else send({ protocol: 1, kind: 'reply', id, ok: true, value });
  }
  function reject(id: number, error: unknown): void {
    if (error instanceof SessionFailure) fail(error);
    else if (typeof WebAssembly !== 'undefined' && error instanceof WebAssembly.RuntimeError) {
      fail(new SessionFailure('trap', error.message));
    } else send({ protocol: 1, kind: 'reply', id, ok: false, error: replyError(error) });
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
      scheduler.submit({ ...policy, run: () => run(message), supersede: () => send({
        protocol: 1, kind: 'reply', id: message.id, ok: false,
        error: { name: SESSION_SUPERSEDED, message: 'Superseded by a newer request' },
      }) });
    }
  });
  unerror = transport.onError((error) => fail(error instanceof SessionFailure ? error :
    new SessionFailure('crash', replyError(error).message)));
  return {
    scheduler,
    emit: (name, payload, transfer) => send({ protocol: 1, kind: 'event', name, payload }, transfer),
  };
}
