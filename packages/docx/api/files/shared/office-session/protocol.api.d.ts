// @betteroffice/docx declaration shared/office-session/protocol.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { SessionFailureCode } from './types';
export type ClientMessage = {
    protocol: 1;
    kind: 'wasm-compile';
    /** The client is fetching the engine and will send it: wait for it instead of fetching. */
    pending?: true;
    /** The engine's response body, fetched by the client; the worker compiles it as it streams. */
    stream?: ReadableStream<Uint8Array>;
    /** The engine's bytes, where the stream could not be transferred; the worker compiles them. */
    bytes?: ArrayBuffer;
} | {
    protocol: 1;
    kind: 'call';
    id: number;
    method: string;
    args: unknown[];
} | {
    protocol: 1;
    kind: 'dispose';
};
export interface ReplyError {
    name: string;
    message: string;
    code?: string;
    refusal?: unknown;
}
export type HostMessage = 
/** The worker waits for the engine body the client announced as pending; old workers never send it. */
{
    protocol: 1;
    kind: 'wasm-accept';
} | {
    protocol: 1;
    kind: 'wasm-module';
    url: string;
    module: WebAssembly.Module;
    hydration?: string;
    version?: string;
    sequence?: number;
} | {
    protocol: 1;
    kind: 'reply';
    id: number;
    ok: true;
    value: unknown;
} | {
    protocol: 1;
    kind: 'reply';
    id: number;
    ok: false;
    error: ReplyError;
} | {
    protocol: 1;
    kind: 'event';
    name: string;
    payload: unknown;
} | {
    protocol: 1;
    kind: 'progress';
    id: number;
} | {
    protocol: 1;
    kind: 'failure';
    code: SessionFailureCode;
    message: string;
    diagnostics?: string;
};
declare const TRANSFER: unique symbol;
declare const DEFERRED: unique symbol;
export interface DeferredReply<T> {
    readonly [DEFERRED]: true;
    readonly promise: Promise<T | TransferResult<T>>;
}
/** Releases the foreground lane while a reply is pending. */
export declare function deferReply<T>(promise: Promise<T | TransferResult<T>>): DeferredReply<T>;
export declare function isDeferredReply(value: unknown): value is DeferredReply<unknown>;
export interface TransferResult<T> {
    readonly [TRANSFER]: true;
    readonly value: T;
    readonly transfer: Transferable[];
}
/** Transfers ownership of a handler result when replying. */
export declare function transferable<T>(value: T, transfer: Transferable[]): TransferResult<T>;
export declare function isTransferResult(value: unknown): value is TransferResult<unknown>;
/** `Object.hasOwn` for ES2020 targets. */
export declare function hasOwn(value: object, key: PropertyKey): boolean;
export declare function isClientMessage(value: unknown): value is ClientMessage;
export declare function isHostMessage(value: unknown): value is HostMessage;
export {};
