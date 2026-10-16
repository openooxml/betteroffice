import { MAX_COLLABORATION_FRAME_BYTES } from '../../../../shared/collaboration-limits';
export { MAX_COLLABORATION_FRAME_BYTES } from '../../../../shared/collaboration-limits';
import {
  MAX_AWARENESS_CURSOR_BYTES,
  MAX_AWARENESS_ENTRIES_PER_UPDATE,
  MAX_AWARENESS_STRING_LENGTH,
  MAX_AWARENESS_TRACKED_PEERS,
  resolvePresenceColor,
  type AwarenessUpdateEntry,
} from './awareness';

export const DEFAULT_MAX_FRAME_BYTES = MAX_COLLABORATION_FRAME_BYTES;
export const DEFAULT_MAX_MESSAGES_PER_FRAME = 4096;
/** Largest update a replica accepts (docx-edit `MAX_UPDATE_BYTES`), the per-update bound. */
export const MAX_UPDATE_BYTES = 64 * 1024 * 1024;
/** Largest configurable frame: one update at the per-update bound and its sync header. */
export const MAX_FRAME_BYTES = MAX_UPDATE_BYTES + 6;

const TOP_LEVEL_SYNC = 0;
const TOP_LEVEL_AWARENESS = 1;
const TOP_LEVEL_AUTH = 2;
const TOP_LEVEL_QUERY_AWARENESS = 3;

const SYNC_STEP_1 = 0;
const SYNC_STEP_2 = 1;
const SYNC_UPDATE = 2;

const AUTH_PERMISSION_DENIED = 0;
const MAX_VAR_UINT = Number.MAX_SAFE_INTEGER;
const MAX_STATE_VECTOR_VAR_UINT_BYTES = 10;
const MAX_NATIVE_CLOCK = 0xffff_ffffn;

export type DecodedProtocolMessage =
  | { type: 'sync-step-1'; stateVector: Uint8Array }
  | { type: 'sync-step-2'; update: Uint8Array }
  | { type: 'update'; update: Uint8Array }
  | { type: 'awareness'; update: Uint8Array }
  | { type: 'auth'; reason: string }
  | { type: 'query-awareness' };

export class ProtocolError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ProtocolError';
  }
}

class AwarenessLimitError extends ProtocolError {}

export interface DecodeAwarenessUpdateOptions {
  trackedClientIds?: Iterable<number>;
  localClientId?: number;
  onDiscard?(error: ProtocolError): void;
}

class Decoder {
  private offset = 0;

  constructor(private readonly bytes: Uint8Array) {}

  get done(): boolean {
    return this.offset === this.bytes.byteLength;
  }

  readVarUint(): number {
    let value = 0;
    let multiplier = 1;
    let count = 0;

    while (true) {
      if (this.offset >= this.bytes.byteLength) {
        throw new ProtocolError('Truncated varUint');
      }

      const byte = this.bytes[this.offset++];
      const digit = byte & 0x7f;
      if (digit > Math.floor((MAX_VAR_UINT - value) / multiplier)) {
        throw new ProtocolError('varUint exceeds Number.MAX_SAFE_INTEGER');
      }

      value += digit * multiplier;
      count += 1;
      if ((byte & 0x80) === 0) {
        if (count > 1 && digit === 0) {
          throw new ProtocolError('Non-canonical varUint');
        }
        return value;
      }

      if (count >= 8) {
        throw new ProtocolError('varUint exceeds Number.MAX_SAFE_INTEGER');
      }
      multiplier *= 128;
    }
  }

  readVarUint8Array(): Uint8Array {
    const length = this.readVarUint();
    const remaining = this.bytes.byteLength - this.offset;
    if (length > remaining) {
      throw new ProtocolError('Truncated varUint8Array');
    }

    const value = this.bytes.slice(this.offset, this.offset + length);
    this.offset += length;
    return value;
  }

  skipVarUint8Array(): void {
    const length = this.readVarUint();
    const remaining = this.bytes.byteLength - this.offset;
    if (length > remaining) {
      throw new ProtocolError('Truncated varUint8Array');
    }
    this.offset += length;
  }

  readVarString(): string {
    const bytes = this.readVarUint8Array();
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (cause) {
      throw new ProtocolError(
        `Invalid UTF-8 string${cause instanceof Error ? `: ${cause.message}` : ''}`
      );
    }
  }
}

export function encodeVarUint(value: number): Uint8Array {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new ProtocolError('varUint value must be a non-negative safe integer');
  }

  const bytes: number[] = [];
  let remaining = value;
  while (remaining >= 128) {
    bytes.push((remaining % 128) | 0x80);
    remaining = Math.floor(remaining / 128);
  }
  bytes.push(remaining);
  return Uint8Array.from(bytes);
}

function encodeFrame(parts: readonly Uint8Array[], maxFrameBytes: number): Uint8Array {
  let length = 0;
  for (const part of parts) {
    length += part.byteLength;
    if (length > maxFrameBytes) {
      throw new ProtocolError(`Frame exceeds ${maxFrameBytes} bytes`);
    }
  }

  const frame = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    frame.set(part, offset);
    offset += part.byteLength;
  }
  return frame;
}

function encodeVarString(value: string): Uint8Array {
  const bytes = new TextEncoder().encode(value);
  return encodeFrame(
    [encodeVarUint(bytes.byteLength), bytes],
    Number.MAX_SAFE_INTEGER
  );
}

function encodeSyncMessage(
  subtype: number,
  payload: Uint8Array,
  maxFrameBytes: number
): Uint8Array {
  return encodeFrame(
    [
      encodeVarUint(TOP_LEVEL_SYNC),
      encodeVarUint(subtype),
      encodeVarUint(payload.byteLength),
      payload,
    ],
    maxFrameBytes
  );
}

/** Largest sync payload that fits a frame of `maxFrameBytes`. */
export function maxSyncPayloadBytes(maxFrameBytes: number): number {
  return maxFrameBytes - 2 - encodeVarUint(maxFrameBytes).byteLength;
}

/** Bytes of a frame holding one sync message with a payload of `payloadBytes`. */
export function syncFrameBytes(payloadBytes: number): number {
  return 2 + encodeVarUint(payloadBytes).byteLength + payloadBytes;
}

/**
 * A state-vector entry (this client id at clock 0) by which a peer asks for a chunked answer to its SyncStep1. Peers
 * that do not know it read an empty client, which changes no diff.
 */
export const CHUNKED_SYNC_CLIENT = Number.MAX_SAFE_INTEGER;

function stateVectorEntries(stateVector: Uint8Array): Array<[number, number]> | null {
  try {
    const decoder = new Decoder(stateVector);
    const count = decoder.readVarUint();
    const entries: Array<[number, number]> = [];
    for (let index = 0; index < count; index++) entries.push([decoder.readVarUint(), decoder.readVarUint()]);
    return decoder.done ? entries : null;
  } catch {
    return null;
  }
}

/** `stateVector` with the chunked-sync entry, or unchanged when it is malformed or already has that client. */
export function withChunkedSyncRequest(stateVector: Uint8Array): Uint8Array {
  const entries = stateVectorEntries(stateVector);
  if (!entries || entries.some(([client]) => client === CHUNKED_SYNC_CLIENT)) return stateVector;
  const body = stateVector.subarray(encodeVarUint(entries.length).byteLength);
  return encodeFrame(
    [encodeVarUint(entries.length + 1), body, encodeVarUint(CHUNKED_SYNC_CLIENT), encodeVarUint(0)],
    Number.MAX_SAFE_INTEGER
  );
}

/** Whether a SyncStep1 state vector asks for a chunked answer. */
export function requestsChunkedSync(stateVector: Uint8Array): boolean {
  return stateVectorEntries(stateVector)?.some(
    ([client, clock]) => client === CHUNKED_SYNC_CLIENT && clock === 0
  ) ?? false;
}

export function encodeSyncStep1(
  stateVector: Uint8Array,
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES
): Uint8Array {
  return encodeSyncMessage(SYNC_STEP_1, stateVector, maxFrameBytes);
}

export function encodeSyncStep2(
  update: Uint8Array,
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES
): Uint8Array {
  return encodeSyncMessage(SYNC_STEP_2, update, maxFrameBytes);
}

export function encodeUpdate(
  update: Uint8Array,
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES
): Uint8Array {
  return encodeSyncMessage(SYNC_UPDATE, update, maxFrameBytes);
}

export function encodeEmptyAwarenessUpdate(
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES
): Uint8Array {
  return encodeAwarenessMessage([], maxFrameBytes);
}

function encodedAwarenessState(entry: AwarenessUpdateEntry): string {
  if (entry.state === null) return 'null';
  return JSON.stringify({
    user: entry.state.user,
    cursor: entry.state.cursor
      ? {
          story: entry.state.cursor.story,
          anchor: [...entry.state.cursor.anchor],
          head: [...entry.state.cursor.head],
        }
      : null,
  });
}

export function encodeAwarenessUpdate(entries: readonly AwarenessUpdateEntry[]): Uint8Array {
  const parts = [encodeVarUint(entries.length)];
  for (const entry of entries) {
    parts.push(
      encodeVarUint(entry.clientId),
      encodeVarUint(entry.clock),
      encodeVarString(encodedAwarenessState(entry))
    );
  }
  return encodeFrame(parts, Number.MAX_SAFE_INTEGER);
}

export function encodeAwarenessMessage(
  entries: readonly AwarenessUpdateEntry[],
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES
): Uint8Array {
  const update = encodeAwarenessUpdate(entries);
  return encodeFrame(
    [
      encodeVarUint(TOP_LEVEL_AWARENESS),
      encodeVarUint(update.byteLength),
      update,
    ],
    maxFrameBytes
  );
}

export function encodeQueryAwareness(
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES
): Uint8Array {
  return encodeFrame([encodeVarUint(TOP_LEVEL_QUERY_AWARENESS)], maxFrameBytes);
}

function decodeByteArray(value: unknown, label: string): Uint8Array {
  if (!Array.isArray(value)) {
    throw new ProtocolError(`${label} must be a byte array`);
  }
  if (value.length > MAX_AWARENESS_CURSOR_BYTES) {
    throw new AwarenessLimitError(`${label} exceeds ${MAX_AWARENESS_CURSOR_BYTES} bytes`);
  }
  if (value.some((byte) => !Number.isInteger(byte) || byte < 0 || byte > 255)) {
    throw new ProtocolError(`${label} must be a byte array`);
  }
  return Uint8Array.from(value as number[]);
}

function decodeAwarenessState(value: string, clientId: number): AwarenessUpdateEntry['state'] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    throw new ProtocolError('Invalid awareness JSON');
  }
  if (parsed === null) return null;
  if (typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new ProtocolError('Awareness state must be an object or null');
  }
  const state = parsed as Record<string, unknown>;
  const userValue = state.user;
  if (typeof userValue !== 'object' || userValue === null || Array.isArray(userValue)) {
    throw new ProtocolError('Awareness user must be an object');
  }
  const userRecord = userValue as Record<string, unknown>;
  if (typeof userRecord.name !== 'string') {
    throw new ProtocolError('Awareness user requires a string name');
  }
  if (userRecord.name.length > MAX_AWARENESS_STRING_LENGTH) {
    throw new AwarenessLimitError(
      `Awareness user name exceeds ${MAX_AWARENESS_STRING_LENGTH} characters`
    );
  }
  const decodedUser = {
    name: userRecord.name,
    color: resolvePresenceColor(clientId, userRecord.color),
  };
  const cursor = state.cursor;
  if (cursor === null || cursor === undefined) {
    return {
      user: decodedUser,
      cursor: null,
    };
  }
  if (typeof cursor !== 'object' || Array.isArray(cursor)) {
    throw new ProtocolError('Awareness cursor must be an object or null');
  }
  const cursorRecord = cursor as Record<string, unknown>;
  if (typeof cursorRecord.story !== 'string') {
    throw new ProtocolError('Awareness cursor requires a string story');
  }
  if (cursorRecord.story.length > MAX_AWARENESS_STRING_LENGTH) {
    throw new AwarenessLimitError(
      `Awareness cursor story exceeds ${MAX_AWARENESS_STRING_LENGTH} characters`
    );
  }
  return {
    user: decodedUser,
    cursor: {
      story: cursorRecord.story,
      anchor: decodeByteArray(cursorRecord.anchor, 'Awareness cursor anchor'),
      head: decodeByteArray(cursorRecord.head, 'Awareness cursor head'),
    },
  };
}

export function decodeAwarenessUpdate(
  update: Uint8Array,
  options: DecodeAwarenessUpdateOptions = {}
): AwarenessUpdateEntry[] {
  const decoder = new Decoder(update);
  const count = decoder.readVarUint();
  if (count > update.byteLength) {
    throw new ProtocolError('Invalid awareness entry count');
  }
  const trackedClientIds = new Set(options.trackedClientIds);
  if (options.localClientId !== undefined) trackedClientIds.delete(options.localClientId);
  const reported = new Set<string>();
  const reportDiscard = (error: ProtocolError): void => {
    if (reported.has(error.message)) return;
    reported.add(error.message);
    options.onDiscard?.(error);
  };
  if (count > MAX_AWARENESS_ENTRIES_PER_UPDATE) {
    reportDiscard(
      new AwarenessLimitError(
        `Awareness update exceeds ${MAX_AWARENESS_ENTRIES_PER_UPDATE} entries`
      )
    );
  }
  const entries: AwarenessUpdateEntry[] = [];
  const decodedCount = Math.min(count, MAX_AWARENESS_ENTRIES_PER_UPDATE);
  for (let index = 0; index < decodedCount; index += 1) {
    const clientId = decoder.readVarUint();
    const clock = decoder.readVarUint();
    const isTracked = clientId === options.localClientId || trackedClientIds.has(clientId);
    if (!isTracked && trackedClientIds.size >= MAX_AWARENESS_TRACKED_PEERS) {
      decoder.skipVarUint8Array();
      reportDiscard(
        new AwarenessLimitError(`Awareness peer limit of ${MAX_AWARENESS_TRACKED_PEERS} reached`)
      );
      continue;
    }
    const encodedState = decoder.readVarString();
    let state: AwarenessUpdateEntry['state'];
    try {
      state = decodeAwarenessState(encodedState, clientId);
    } catch (error) {
      if (!(error instanceof AwarenessLimitError)) throw error;
      reportDiscard(error);
      continue;
    }
    entries.push({ clientId, clock, state });
    if (clientId !== options.localClientId) trackedClientIds.add(clientId);
  }
  if (count <= MAX_AWARENESS_ENTRIES_PER_UPDATE && !decoder.done) {
    throw new ProtocolError('Trailing awareness data');
  }
  return entries;
}

function varUintReader(payload: Uint8Array, label: string): { read(): bigint; done(): boolean } {
  let offset = 0;
  return {
    read() {
      let value = 0n;
      for (let index = 0; index < MAX_STATE_VECTOR_VAR_UINT_BYTES; index += 1) {
        if (offset >= payload.byteLength) throw new ProtocolError(`Truncated ${label.toLowerCase()}`);
        const byte = payload[offset++];
        value |= BigInt(byte & 0x7f) << BigInt(7 * index);
        if (byte < 0x80) return value;
      }
      throw new ProtocolError(`${label} varUint exceeds 64 bits`);
    },
    done: () => offset === payload.byteLength,
  };
}

function decodeStateVector(payload: Uint8Array): Map<bigint, bigint> {
  const reader = varUintReader(payload, 'State vector');
  const clocks = new Map<bigint, bigint>();
  for (let count = reader.read(); count > 0n; count -= 1n) clocks.set(reader.read(), reader.read());
  if (!reader.done()) throw new ProtocolError('Trailing state vector bytes');
  return clocks;
}

/** Whether `payload` decodes as one v1 state vector. */
export function isStateVector(payload: Uint8Array): boolean {
  try {
    decodeStateVector(payload);
    return true;
  } catch {
    return false;
  }
}

/** Whether `vector` has every client clock of `other`; both are v1 state vectors. */
export function stateVectorDominates(vector: Uint8Array, other: Uint8Array): boolean {
  if (vector.byteLength === other.byteLength && vector.every((byte, index) => byte === other[index])) {
    return true;
  }
  const clocks = decodeStateVector(vector);
  for (const [client, clock] of decodeStateVector(other)) {
    if (clock > (clocks.get(client) ?? 0n)) return false;
  }
  return true;
}

/** `vector` with `clientId`'s clock one past its own (1 when absent); undefined past the native clock limit. */
export function withClientClockAhead(vector: Uint8Array, clientId: number): Uint8Array | undefined {
  if (!Number.isSafeInteger(clientId) || clientId < 0) {
    throw new ProtocolError('Client id must be a non-negative safe integer');
  }
  const clocks = decodeStateVector(vector);
  const client = BigInt(clientId);
  const clock = (clocks.get(client) ?? 0n) + 1n;
  if (clock > MAX_NATIVE_CLOCK) return undefined;
  clocks.set(client, clock);
  const bytes: number[] = [];
  for (let value of [BigInt(clocks.size), ...[...clocks].flat()]) {
    for (; value >= 0x80n; value >>= 7n) bytes.push(Number(value & 0x7fn) | 0x80);
    bytes.push(Number(value));
  }
  return Uint8Array.from(bytes);
}

export function decodeMessages(
  frame: Uint8Array,
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES,
  maxMessages = DEFAULT_MAX_MESSAGES_PER_FRAME
): DecodedProtocolMessage[] {
  if (!(frame instanceof Uint8Array)) {
    throw new ProtocolError('Frame must be a Uint8Array');
  }
  if (frame.byteLength === 0) {
    throw new ProtocolError('Frame is empty');
  }
  if (frame.byteLength > maxFrameBytes) {
    throw new ProtocolError(`Frame exceeds ${maxFrameBytes} bytes`);
  }
  if (!Number.isSafeInteger(maxMessages) || maxMessages < 1) {
    throw new ProtocolError('Message limit must be a positive integer');
  }

  const decoder = new Decoder(frame);
  const messages: DecodedProtocolMessage[] = [];
  while (!decoder.done) {
    if (messages.length >= maxMessages) {
      throw new ProtocolError(`Frame exceeds ${maxMessages} messages`);
    }
    const type = decoder.readVarUint();
    switch (type) {
      case TOP_LEVEL_SYNC: {
        const subtype = decoder.readVarUint();
        const payload = decoder.readVarUint8Array();
        if (subtype === SYNC_STEP_1) {
          messages.push({ type: 'sync-step-1', stateVector: payload });
        } else if (subtype === SYNC_STEP_2) {
          messages.push({ type: 'sync-step-2', update: payload });
        } else if (subtype === SYNC_UPDATE) {
          messages.push({ type: 'update', update: payload });
        } else {
          throw new ProtocolError(`Unknown sync message type ${subtype}`);
        }
        break;
      }
      case TOP_LEVEL_AWARENESS:
        messages.push({ type: 'awareness', update: decoder.readVarUint8Array() });
        break;
      case TOP_LEVEL_AUTH: {
        const subtype = decoder.readVarUint();
        if (subtype !== AUTH_PERMISSION_DENIED) {
          throw new ProtocolError(`Unknown auth message type ${subtype}`);
        }
        messages.push({ type: 'auth', reason: decoder.readVarString() });
        break;
      }
      case TOP_LEVEL_QUERY_AWARENESS:
        messages.push({ type: 'query-awareness' });
        break;
      default:
        throw new ProtocolError(`Unknown top-level message type ${type}`);
    }
  }

  return messages;
}
