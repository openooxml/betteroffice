/**
 * Wire format between a main-thread collaboration facade and the provider its document worker runs.
 * Main owns the host transport's lifecycle and forwards its bytes untouched; the worker runs the protocol and
 * answers with frames to send, state and the transport calls its own failures need.
 */

export type PipeErrorCode = 'backpressure' | 'protocol' | 'replica' | 'transport';

export interface PipeError {
  code: PipeErrorCode;
  message: string;
}

export interface PipeUser {
  name: string;
  color?: string;
}

export interface PipeProviderOptions {
  maxFrameBytes?: number;
  maxMessagesPerFrame?: number;
  syncChunkBytes?: number;
  maxSyncBytes?: number;
  user?: PipeUser;
}

/** The docx cursor; the default cursor type `C` of the pipe types. */
export interface PipeCursor {
  story: string;
  anchor: Uint8Array;
  head: Uint8Array;
}

export type PipeTransportEvent =
  | { type: 'open' }
  | { type: 'message'; data: ArrayBuffer }
  | { type: 'close'; reason?: string };

/**
 * Main to worker, in order; `seq` increases. `connect`, `disconnect`, `destroy` and `fail` replay what main already
 * did to the host transport, so the worker sends nothing and asks for no transport call while it runs them.
 */
export type PipeOp<C = PipeCursor> =
  | { op: 'create'; seq: number; options: PipeProviderOptions }
  | { op: 'connect'; seq: number }
  | { op: 'disconnect'; seq: number }
  | { op: 'destroy'; seq: number }
  | { op: 'fail'; seq: number; error: PipeError }
  | { op: 'cursor'; seq: number; cursor: C | null; broadcast: boolean }
  | { op: 'event'; seq: number; event: PipeTransportEvent };

/** The docx peer; the default peer type `P` of the pipe types. */
export interface PipePeer {
  clientId: number;
  clock: number;
  user: { name: string; color: string };
  cursor: PipeCursor | null;
  lastSeenAt: number;
  cursorMovedAt: number;
  inferredCursor: { clientId: number; story: string; paraId: string; endOffset: number } | null;
}

/**
 * Worker to main, in order. `through` is the last op the worker had run; main drops output older than its own
 * latest lifecycle call.
 */
export type PipeOutput<P = PipePeer> =
  | { do: 'unsubscribe'; through: number }
  | { do: 'disconnect'; through: number }
  | { do: 'send'; data: ArrayBuffer; through: number }
  /** The awareness leave main sends when the host disconnects first; null while there is none to send. */
  | { do: 'leave'; data: ArrayBuffer | null; through: number }
  | { do: 'status'; status: 'disconnected' | 'connecting' | 'connected' | 'destroyed'; synced: boolean; through: number }
  | { do: 'peers'; peers: P[]; timeOrigin: number; through: number }
  | { do: 'error'; error: PipeError; through: number }
  /** The worker has no provider under this id. */
  | { do: 'lost' };

export interface PipeOutputBatch<P = PipePeer> {
  provider: number;
  out: PipeOutput<P>[];
}

export interface PipeCommandPayload<C = PipeCursor> {
  provider: number;
  ops: PipeOp<C>[];
}

export interface PipePayloadShape {
  /** Checks a non-null cursor; its byte fields stay within `maxFrameBytes`. */
  validCursor(cursor: unknown, maxFrameBytes: number): boolean;
  /** The keys a `create` op's options may carry. */
  optionKeys: ReadonlySet<string>;
}

/** Buffers a pipe payload holds, for transfer. */
export function pipeOpBuffers<C>(ops: readonly PipeOp<C>[]): ArrayBuffer[] {
  const buffers: ArrayBuffer[] = [];
  for (const op of ops) {
    if (op.op === 'event' && op.event.type === 'message') buffers.push(op.event.data);
  }
  return buffers;
}

export function ownedBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength
    ? bytes.buffer
    : bytes.slice().buffer;
}

const ERROR_CODES = new Set<string>(['backpressure', 'protocol', 'replica', 'transport']);
export const PIPE_OPTION_KEYS: ReadonlySet<string> =
  new Set(['maxFrameBytes', 'maxMessagesPerFrame', 'syncChunkBytes', 'maxSyncBytes', 'user']);

const sequence = (value: unknown) => Number.isSafeInteger(value) && (value as number) >= 0;
const bytes = (value: unknown, max: number) => value instanceof Uint8Array && value.byteLength <= max;

function validError(value: unknown): boolean {
  const error = value as Record<string, unknown> | null;
  return typeof error === 'object' && error !== null && ERROR_CODES.has(error.code as string) &&
    typeof error.message === 'string';
}

function validEvent(event: unknown, maxFrameBytes: number): boolean {
  if (typeof event !== 'object' || event === null) return false;
  const value = event as Record<string, unknown>;
  switch (value.type) {
    case 'open': return true;
    case 'message': return value.data instanceof ArrayBuffer && value.data.byteLength <= maxFrameBytes;
    case 'close': return value.reason === undefined || typeof value.reason === 'string';
    default: return false;
  }
}

/** The docx cursor check. */
export function validPipeCursor(cursor: unknown, maxFrameBytes: number): boolean {
  if (typeof cursor !== 'object' || cursor === null) return false;
  const value = cursor as Record<string, unknown>;
  return typeof value.story === 'string' && bytes(value.anchor, maxFrameBytes) && bytes(value.head, maxFrameBytes);
}

/** Shape check for a pipe command payload; frames past `maxFrameBytes` fail it. */
export function validPipePayload<C = PipeCursor>(
  payload: unknown,
  maxFrameBytes: number,
  shape: PipePayloadShape
): payload is PipeCommandPayload<C> {
  if (typeof payload !== 'object' || payload === null) return false;
  const { provider, ops } = payload as Record<string, unknown>;
  if (!Number.isSafeInteger(provider) || (provider as number) < 0 || !Array.isArray(ops) || ops.length === 0) return false;
  return ops.every((entry: unknown) => {
    if (typeof entry !== 'object' || entry === null) return false;
    const op = entry as Record<string, unknown>;
    if (!sequence(op.seq)) return false;
    switch (op.op) {
      case 'create':
        return typeof op.options === 'object' && op.options !== null &&
          Object.keys(op.options).every((key) => shape.optionKeys.has(key));
      case 'connect':
      case 'disconnect':
      case 'destroy':
        return true;
      case 'fail':
        return validError(op.error);
      case 'cursor':
        return typeof op.broadcast === 'boolean' &&
          (op.cursor === null || shape.validCursor(op.cursor, maxFrameBytes));
      case 'event':
        return validEvent(op.event, maxFrameBytes);
      default:
        return false;
    }
  });
}

/** Epoch time of this context's provider clock zero: performance.now() when present, else Date.now(). */
export function clockOrigin(): number {
  const performance = globalThis.performance;
  if (typeof performance?.now !== 'function') return 0;
  return typeof performance.timeOrigin === 'number' ? performance.timeOrigin : Date.now() - performance.now();
}
