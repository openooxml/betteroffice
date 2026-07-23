import { sanitizePresenceColor, type AwarenessUpdateEntry } from './presence';
import type { PptxPresenceCursor, PptxPresenceState } from './types';

export const DEFAULT_MAX_FRAME_BYTES = 16 * 1024 * 1024;
export const DEFAULT_MAX_MESSAGES_PER_FRAME = 4096;

const TOP_LEVEL_SYNC = 0;
const TOP_LEVEL_AWARENESS = 1;
const TOP_LEVEL_AUTH = 2;
const TOP_LEVEL_QUERY_AWARENESS = 3;

const SYNC_STEP_1 = 0;
const SYNC_STEP_2 = 1;
const SYNC_UPDATE = 2;

const AUTH_PERMISSION_DENIED = 0;
const MAX_VAR_UINT = Number.MAX_SAFE_INTEGER;
const MAX_AWARENESS_STRING_LENGTH = 1024;

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

function encodeVarString(value: string): Uint8Array {
  const bytes = new TextEncoder().encode(value);
  return encodeFrame([encodeVarUint(bytes.byteLength), bytes], DEFAULT_MAX_FRAME_BYTES);
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
  return encodeAwarenessUpdate([], maxFrameBytes);
}

export function encodeAwarenessUpdate(
  entries: readonly AwarenessUpdateEntry[],
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES
): Uint8Array {
  if (entries.length > DEFAULT_MAX_MESSAGES_PER_FRAME) {
    throw new ProtocolError(
      `Awareness update exceeds ${DEFAULT_MAX_MESSAGES_PER_FRAME} entries`
    );
  }
  const updateParts: Uint8Array[] = [encodeVarUint(entries.length)];
  for (const entry of entries) {
    validateAwarenessEntry(entry);
    const encodedState = JSON.stringify(entry.state);
    if (encodedState === undefined) {
      throw new ProtocolError('Awareness state is not JSON serializable');
    }
    updateParts.push(
      encodeVarUint(entry.clientId),
      encodeVarUint(entry.clock),
      encodeVarString(encodedState)
    );
  }
  const update = encodeFrame(updateParts, maxFrameBytes);
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

export function decodeAwarenessUpdate(
  update: Uint8Array,
  maxEntries = DEFAULT_MAX_MESSAGES_PER_FRAME
): AwarenessUpdateEntry[] {
  if (!(update instanceof Uint8Array)) {
    throw new ProtocolError('Awareness update must be a Uint8Array');
  }
  if (!Number.isSafeInteger(maxEntries) || maxEntries < 1) {
    throw new ProtocolError('Awareness entry limit must be a positive integer');
  }
  const decoder = new Decoder(update);
  const count = decoder.readVarUint();
  if (count > maxEntries) {
    throw new ProtocolError(`Awareness update exceeds ${maxEntries} entries`);
  }

  const entries: AwarenessUpdateEntry[] = [];
  for (let index = 0; index < count; index += 1) {
    const clientId = decoder.readVarUint();
    const clock = decoder.readVarUint();
    let parsed: unknown;
    try {
      parsed = JSON.parse(decoder.readVarString());
    } catch (cause) {
      if (cause instanceof ProtocolError) throw cause;
      throw new ProtocolError(
        `Invalid awareness JSON${cause instanceof Error ? `: ${cause.message}` : ''}`
      );
    }
    const entry = {
      clientId,
      clock,
      state: parseAwarenessState(parsed, clientId, clock),
    };
    validateAwarenessEntry(entry);
    entries.push(entry);
  }
  if (!decoder.done) throw new ProtocolError('Awareness update has trailing bytes');
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

function parseAwarenessState(
  value: unknown,
  clientId: number,
  clock: number
): PptxPresenceState | null {
  if (value === null) return null;
  if (!isRecord(value)) throw new ProtocolError('Awareness state must be an object or null');
  if (value.clientId !== clientId || value.clock !== clock) {
    throw new ProtocolError('Awareness state identity does not match its entry');
  }
  if (!isRecord(value.user)) throw new ProtocolError('Awareness user must be an object');
  const name = requireAwarenessString(value.user.name, 'Awareness user name');
  const color = sanitizePresenceColor(clientId, value.user.color);
  return {
    clientId,
    clock,
    user: { name, color },
    cursor: parseAwarenessCursor(value.cursor),
  };
}

function parseAwarenessCursor(value: unknown): PptxPresenceCursor | null {
  if (value === null || value === undefined) return null;
  if (!isRecord(value)) throw new ProtocolError('Awareness cursor must be an object or null');
  const slideId = requireAwarenessString(value.slideId, 'Awareness slideId');
  const shapeId =
    value.shapeId === undefined
      ? undefined
      : requireAwarenessString(value.shapeId, 'Awareness shapeId');
  return shapeId ? { slideId, shapeId } : { slideId };
}

function validateAwarenessEntry(entry: AwarenessUpdateEntry): void {
  requireVarUint(entry.clientId, 'Awareness clientId');
  requireVarUint(entry.clock, 'Awareness clock');
  if (entry.state === null) return;
  parseAwarenessState(entry.state, entry.clientId, entry.clock);
}

function requireVarUint(value: number, name: string): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new ProtocolError(`${name} must be a non-negative safe integer`);
  }
}

function requireAwarenessString(value: unknown, name: string): string {
  if (
    typeof value !== 'string' ||
    value.length === 0 ||
    value.length > MAX_AWARENESS_STRING_LENGTH
  ) {
    throw new ProtocolError(
      `${name} must be between 1 and ${MAX_AWARENESS_STRING_LENGTH} characters`
    );
  }
  return value;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
