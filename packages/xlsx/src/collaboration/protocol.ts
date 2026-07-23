import type { AwarenessPayload, AwarenessUpdate } from './types';

export const DEFAULT_MAX_FRAME_BYTES = 64 * 1024 * 1024 + 16;
export const DEFAULT_MAX_MESSAGES_PER_FRAME = 4096;
export const DEFAULT_MAX_AWARENESS_STATES = 1024;

const TOP_LEVEL_SYNC = 0;
const TOP_LEVEL_AWARENESS = 1;
const TOP_LEVEL_AUTH = 2;
const TOP_LEVEL_QUERY_AWARENESS = 3;

const SYNC_STEP_1 = 0;
const SYNC_STEP_2 = 1;
const SYNC_UPDATE = 2;

const AUTH_PERMISSION_DENIED = 0;
const MAX_VAR_UINT = Number.MAX_SAFE_INTEGER;

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

function encodeVarString(value: string): Uint8Array[] {
  const bytes = new TextEncoder().encode(value);
  return [encodeVarUint(bytes.byteLength), bytes];
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function awarenessCell(value: unknown): { row: number; col: number } {
  if (!isObject(value)) throw new ProtocolError('Awareness cell must be an object');
  const { row, col } = value;
  if (
    !Number.isSafeInteger(row) ||
    !Number.isSafeInteger(col) ||
    (row as number) < 0 ||
    (col as number) < 0
  ) {
    throw new ProtocolError('Awareness cell coordinates must be non-negative safe integers');
  }
  return { row: row as number, col: col as number };
}

function awarenessPayload(value: unknown): AwarenessPayload | null {
  if (value === null) return null;
  if (!isObject(value) || !isObject(value.user)) {
    throw new ProtocolError('Awareness state must contain a user');
  }
  const { name, color } = value.user;
  if (typeof name !== 'string' || name.length === 0 || name.length > 128) {
    throw new ProtocolError('Awareness user name must contain 1 to 128 characters');
  }
  if (typeof color !== 'string' || !/^#[0-9a-f]{6}$/i.test(color)) {
    throw new ProtocolError('Awareness user color must be #RRGGBB');
  }

  const cursor = value.cursor;
  if (cursor === null || cursor === undefined) {
    return { user: { name, color }, cursor: null };
  }
  if (!isObject(cursor) || typeof cursor.sheet !== 'string') {
    throw new ProtocolError('Awareness cursor must contain a sheet identifier');
  }
  if (cursor.sheet.length === 0 || cursor.sheet.length > 256) {
    throw new ProtocolError('Awareness sheet identifier must contain 1 to 256 characters');
  }
  return {
    user: { name, color },
    cursor: {
      sheet: cursor.sheet,
      anchor: awarenessCell(cursor.anchor),
      head: awarenessCell(cursor.head),
    },
  };
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
  updates: readonly AwarenessUpdate[],
  maxFrameBytes = DEFAULT_MAX_FRAME_BYTES
): Uint8Array {
  if (updates.length > DEFAULT_MAX_AWARENESS_STATES) {
    throw new ProtocolError(
      `Awareness update exceeds ${DEFAULT_MAX_AWARENESS_STATES} states`
    );
  }
  const payloadParts: Uint8Array[] = [encodeVarUint(updates.length)];
  for (const update of updates) {
    if (!Number.isSafeInteger(update.clientId) || update.clientId <= 0) {
      throw new ProtocolError('Awareness clientId must be a positive safe integer');
    }
    if (!Number.isSafeInteger(update.clock) || update.clock < 0) {
      throw new ProtocolError('Awareness clock must be a non-negative safe integer');
    }
    const state = awarenessPayload(update.state);
    payloadParts.push(
      encodeVarUint(update.clientId),
      encodeVarUint(update.clock),
      ...encodeVarString(JSON.stringify(state))
    );
  }
  const payload = encodeFrame(payloadParts, maxFrameBytes);
  return encodeFrame(
    [
      encodeVarUint(TOP_LEVEL_AWARENESS),
      encodeVarUint(payload.byteLength),
      payload,
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
  maxStates = DEFAULT_MAX_AWARENESS_STATES
): AwarenessUpdate[] {
  if (!(update instanceof Uint8Array)) {
    throw new ProtocolError('Awareness update must be a Uint8Array');
  }
  if (!Number.isSafeInteger(maxStates) || maxStates < 1) {
    throw new ProtocolError('Awareness state limit must be a positive integer');
  }
  const decoder = new Decoder(update);
  const count = decoder.readVarUint();
  if (count > maxStates) {
    throw new ProtocolError(`Awareness update exceeds ${maxStates} states`);
  }
  const updates: AwarenessUpdate[] = [];
  for (let index = 0; index < count; index++) {
    const clientId = decoder.readVarUint();
    const clock = decoder.readVarUint();
    if (clientId <= 0) {
      throw new ProtocolError('Awareness clientId must be a positive safe integer');
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(decoder.readVarString());
    } catch (cause) {
      if (cause instanceof ProtocolError) throw cause;
      throw new ProtocolError(
        `Invalid awareness JSON${cause instanceof Error ? `: ${cause.message}` : ''}`
      );
    }
    updates.push({ clientId, clock, state: awarenessPayload(parsed) });
  }
  if (!decoder.done) throw new ProtocolError('Trailing awareness update data');
  return updates;
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
