import type { SessionFailureCode } from './types';

export type ClientMessage =
  | { protocol: 1; kind: 'wasm-compile' }
  | { protocol: 1; kind: 'call'; id: number; method: string; args: unknown[] }
  | { protocol: 1; kind: 'dispose' };

export interface ReplyError {
  name: string;
  message: string;
  refusal?: unknown;
}

export type HostMessage =
  | { protocol: 1; kind: 'wasm-module'; url: string; module: WebAssembly.Module;
      hydration?: string; version?: string; sequence?: number }
  | { protocol: 1; kind: 'reply'; id: number; ok: true; value: unknown }
  | { protocol: 1; kind: 'reply'; id: number; ok: false; error: ReplyError }
  | { protocol: 1; kind: 'event'; name: string; payload: unknown }
  | { protocol: 1; kind: 'failure'; code: SessionFailureCode; message: string; diagnostics?: string };

const TRANSFER = Symbol('session-transfer');
const DEFERRED = Symbol('session-deferred');

export interface DeferredReply<T> {
  readonly [DEFERRED]: true;
  readonly promise: Promise<T | TransferResult<T>>;
}

/** Releases the foreground lane while a reply is pending. */
export function deferReply<T>(promise: Promise<T | TransferResult<T>>): DeferredReply<T> {
  return { [DEFERRED]: true, promise };
}

export function isDeferredReply(value: unknown): value is DeferredReply<unknown> {
  return record(value) && value[DEFERRED] === true;
}

export interface TransferResult<T> {
  readonly [TRANSFER]: true;
  readonly value: T;
  readonly transfer: Transferable[];
}

/** Transfers ownership of a handler result when replying. */
export function transferable<T>(value: T, transfer: Transferable[]): TransferResult<T> {
  return { [TRANSFER]: true, value, transfer };
}

export function isTransferResult(value: unknown): value is TransferResult<unknown> {
  return record(value) && value[TRANSFER] === true;
}

/** `Object.hasOwn` for ES2020 targets. */
export function hasOwn(value: object, key: PropertyKey): boolean {
  return Object.prototype.hasOwnProperty.call(value, key);
}

function record(value: unknown): value is Record<PropertyKey, unknown> {
  return value !== null && typeof value === 'object';
}

function requestId(value: unknown): boolean {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0;
}

export function isClientMessage(value: unknown): value is ClientMessage {
  if (!record(value) || value.protocol !== 1) return false;
  return value.kind === 'wasm-compile' || value.kind === 'dispose' || (
    value.kind === 'call' && requestId(value.id) &&
    typeof value.method === 'string' && Array.isArray(value.args)
  );
}

export function isHostMessage(value: unknown): value is HostMessage {
  if (!record(value) || value.protocol !== 1) return false;
  switch (value.kind) {
    case 'wasm-module':
      try {
        return typeof value.url === 'string' &&
          (value.hydration === undefined || typeof value.hydration === 'string') &&
          Object.prototype.toString.call(value.module) === '[object WebAssembly.Module]';
      } catch { return false; }
    case 'reply':
      return requestId(value.id) && (
        (value.ok === true && hasOwn(value, 'value')) ||
        (value.ok === false && record(value.error) &&
          typeof value.error.name === 'string' && typeof value.error.message === 'string')
      );
    case 'event':
      return typeof value.name === 'string' && hasOwn(value, 'payload');
    case 'failure':
      return ['crash', 'trap', 'out-of-memory', 'message', 'silence', 'disposed'].includes(
        value.code as string
      ) && typeof value.message === 'string' &&
        (value.diagnostics === undefined || typeof value.diagnostics === 'string');
    default:
      return false;
  }
}
