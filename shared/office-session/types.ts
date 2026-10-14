import type { DeferredReply, TransferResult } from './protocol';
import type { ResidentScheduler } from './scheduler';

export const SESSION_SUPERSEDED = 'SessionSuperseded';
export type SessionScheduler = Pick<ResidentScheduler,
  'schedule' | 'dispatch' | 'bump' | 'budget' | 'pending' | 'version' | 'generation'>;

/** Format-owned, structured-cloneable methods. */
export type SessionMethods = Record<string, (...args: any[]) => unknown>;

export type Promisified<M extends SessionMethods> = {
  [K in keyof M]: (...args: Parameters<M[K]>) => Promise<Awaited<ReturnType<M[K]>>>;
};

export type MethodHandlers<M extends SessionMethods, C> = {
  [K in keyof M]: (
    context: C,
    ...args: Parameters<M[K]>
  ) =>
    | Awaited<ReturnType<M[K]>>
    | TransferResult<Awaited<ReturnType<M[K]>>>
    | DeferredReply<Awaited<ReturnType<M[K]>>>
    | Promise<Awaited<ReturnType<M[K]>> | TransferResult<Awaited<ReturnType<M[K]>>> |
        DeferredReply<Awaited<ReturnType<M[K]>>>>;
};

export interface MethodPolicy {
  lane: 'input' | 'collab' | 'interactive';
  mutates?: boolean;
  reframes?: boolean;
  reorderable?: boolean;
  key?: string;
  replaceableBy?: string;
  userInput?: boolean;
  holdsIdleTasks?: boolean;
}

export type MethodPolicies<M extends SessionMethods> = {
  [K in keyof M]: MethodPolicy | ((...args: Parameters<M[K]>) => MethodPolicy);
};
export type SessionEvents = Record<string, unknown>;

export interface SessionState {
  format: string;
  stage: 'preview' | 'ready';
  version: number;
  dirty: boolean;
}

export type SessionFailureCode = 'crash' | 'trap' | 'out-of-memory' | 'message' | 'silence' | 'disposed';

/** A terminal session failure. */
export class SessionFailure extends Error {
  constructor(
    readonly code: SessionFailureCode,
    message: string,
    readonly diagnostics?: string
  ) {
    super(message);
    this.name = 'SessionFailure';
  }
}
