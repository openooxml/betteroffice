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
 * Reserved state-vector entries at clock 0 that ask something of the peer answering a SyncStep1. Peers that do not
 * know them read an empty client, which changes no diff. CHUNKED asks for a chunked answer; OPEN marks the opening
 * request, which the peer answers with one SyncStep1 of its own.
 */
export const CHUNKED_SYNC_CLIENT = Number.MAX_SAFE_INTEGER;
export const OPEN_SYNC_CLIENT = Number.MAX_SAFE_INTEGER - 1;

/** `clientId`, refused when it is one of the reserved request entries. */
export function requireUnreservedClientId(clientId: number): number {
  if (clientId >= OPEN_SYNC_CLIENT) {
    throw new RangeError('clientId must be below Number.MAX_SAFE_INTEGER - 1');
  }
  return clientId;
}

interface StateVectorEntry {
  client: bigint;
  clock: bigint;
  start: number;
  end: number;
}

function stateVectorEntries(body: Uint8Array): { header: number; entries: StateVectorEntry[] } | undefined {
  let offset = 0;
  const read = (): bigint | undefined => {
    let value = 0n;
    for (let index = 0; index < MAX_STATE_VECTOR_VAR_UINT_BYTES && offset < body.byteLength; index += 1) {
      const byte = body[offset++];
      value |= BigInt(byte & 0x7f) << BigInt(7 * index);
      if (byte < 0x80) return value;
    }
    return undefined;
  };
  const count = read();
  if (count === undefined) return undefined;
  const header = offset;
  const entries: StateVectorEntry[] = [];
  for (let left = count; left > 0n; left -= 1n) {
    const start = offset;
    const client = read();
    const clock = read();
    if (client === undefined || clock === undefined) return undefined;
    entries.push({ client, clock, start, end: offset });
  }
  return offset === body.byteLength ? { header, entries } : undefined;
}

/** `stateVector` with the requested entries, each added once; unchanged when it is malformed. */
export function withSyncRequest(
  stateVector: Uint8Array,
  { open = false, chunked = false }: { open?: boolean; chunked?: boolean }
): Uint8Array {
  const decoded = stateVectorEntries(stateVector);
  if (!decoded) return stateVector;
  const clients = [...(open ? [OPEN_SYNC_CLIENT] : []), ...(chunked ? [CHUNKED_SYNC_CLIENT] : [])]
    .filter((client) => !decoded.entries.some((entry) => entry.client === BigInt(client)));
  if (!clients.length) return stateVector;
  return encodeFrame([
    encodeVarUint(decoded.entries.length + clients.length),
    stateVector.subarray(decoded.header),
    ...clients.flatMap((client) => [encodeVarUint(client), encodeVarUint(0)]),
  ], Number.MAX_SAFE_INTEGER);
}

/** Whether the history behind `stateVector` already holds the open request client, which then cannot carry the mark. */
export function holdsOpenSyncClient(stateVector: Uint8Array): boolean {
  return stateVectorEntries(stateVector)?.entries
    .some((entry) => entry.client === BigInt(OPEN_SYNC_CLIENT) && entry.clock !== 0n) ?? false;
}

/** The state vector of a SyncStep1 without its request entries, and which requests it carried. */
export function parseSyncRequest(
  stateVector: Uint8Array
): { stateVector: Uint8Array; open: boolean; chunked: boolean } {
  const decoded = stateVectorEntries(stateVector);
  const asks = (entry: StateVectorEntry, client: number) => entry.client === BigInt(client) && entry.clock === 0n;
  const open = decoded?.entries.some((entry) => asks(entry, OPEN_SYNC_CLIENT)) ?? false;
  const chunked = decoded?.entries.some((entry) => asks(entry, CHUNKED_SYNC_CLIENT)) ?? false;
  if (!decoded || (!open && !chunked)) return { stateVector, open, chunked };
  const kept = decoded.entries.filter((entry) => !asks(entry, OPEN_SYNC_CLIENT) && !asks(entry, CHUNKED_SYNC_CLIENT));
  return {
    stateVector: encodeFrame([
      encodeVarUint(kept.length), ...kept.map(({ start, end }) => stateVector.subarray(start, end)),
    ], Number.MAX_SAFE_INTEGER),
    open,
    chunked,
  };
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
