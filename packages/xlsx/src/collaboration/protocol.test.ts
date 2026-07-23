import { describe, expect, it } from 'bun:test';
import {
  decodeAwarenessUpdate,
  decodeMessages,
  encodeAwarenessUpdate,
  encodeEmptyAwarenessUpdate,
  encodeQueryAwareness,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
  encodeVarUint,
  ProtocolError,
} from './protocol';
import { colorForClientId } from './awareness';

function concat(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((length, part) => length + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.byteLength;
  }
  return output;
}

describe('collaboration protocol encoding', () => {
  it('matches sync-v1 golden frames', () => {
    expect([...encodeSyncStep1(Uint8Array.of(1, 2))]).toEqual([0, 0, 2, 1, 2]);
    expect([...encodeSyncStep2(Uint8Array.of(3, 4))]).toEqual([0, 1, 2, 3, 4]);
    expect([...encodeUpdate(Uint8Array.of(5, 6))]).toEqual([0, 2, 2, 5, 6]);
    expect([...encodeEmptyAwarenessUpdate()]).toEqual([1, 1, 0]);
    expect([...encodeQueryAwareness()]).toEqual([3]);
  });

  it('round-trips awareness states and explicit leaves', () => {
    const updates = [
      {
        clientId: 42,
        clock: 7,
        state: {
          user: { name: 'Quiet Otter', color: '#0B57D0' },
          cursor: {
            sheet: 'sheet:0',
            anchor: { row: 2, col: 3 },
            head: { row: 5, col: 8 },
          },
        },
      },
      { clientId: 43, clock: 9, state: null },
    ];
    const [message] = decodeMessages(encodeAwarenessUpdate(updates));
    expect(message.type).toBe('awareness');
    if (message.type !== 'awareness') throw new Error('expected awareness');
    expect(decodeAwarenessUpdate(message.update)).toEqual(updates);
  });

  it('normalizes an absent cursor to present without a selection', () => {
    const state = JSON.stringify({
      user: { name: 'Quiet Otter', color: '#0B57D0' },
    });
    const bytes = new TextEncoder().encode(state);
    expect(
      decodeAwarenessUpdate(
        Uint8Array.from([1, 42, 7, bytes.byteLength, ...bytes])
      )
    ).toEqual([
      {
        clientId: 42,
        clock: 7,
        state: {
          user: { name: 'Quiet Otter', color: '#0B57D0' },
          cursor: null,
        },
      },
    ]);
  });

  it('sanitizes peer colors to a strict hex value at decode time', () => {
    const state = JSON.stringify({
      user: { name: 'Quiet Otter', color: 'url(https://example.invalid/peer)' },
    });
    const bytes = new TextEncoder().encode(state);
    expect(
      decodeAwarenessUpdate(
        Uint8Array.from([1, 42, 7, bytes.byteLength, ...bytes])
      )[0].state?.user.color
    ).toBe(colorForClientId(42));

    const shortHex = new TextEncoder().encode(
      JSON.stringify({ user: { name: 'Quiet Otter', color: '#aBc' } })
    );
    expect(
      decodeAwarenessUpdate(
        Uint8Array.from([1, 42, 8, shortHex.byteLength, ...shortHex])
      )[0].state?.user.color
    ).toBe('#AABBCC');
  });

  it('uses standard multibyte varUint lengths', () => {
    const update = new Uint8Array(130).fill(7);
    const frame = encodeUpdate(update);
    expect([...frame.subarray(0, 4)]).toEqual([0, 2, 130, 1]);
    expect(decodeMessages(frame)).toEqual([{ type: 'update', update }]);
  });

  it('encodes the full safe varUint range canonically', () => {
    expect([...encodeVarUint(127)]).toEqual([127]);
    expect([...encodeVarUint(128)]).toEqual([128, 1]);
    expect(encodeVarUint(Number.MAX_SAFE_INTEGER).byteLength).toBe(8);
    expect(() => encodeVarUint(-1)).toThrow(ProtocolError);
    expect(() => encodeVarUint(Number.MAX_SAFE_INTEGER + 1)).toThrow(ProtocolError);
  });
});

describe('collaboration protocol decoding', () => {
  it('safely parses concatenated standard messages', () => {
    const frame = concat(
      encodeSyncStep1(Uint8Array.of(1)),
      encodeSyncStep2(Uint8Array.of(2)),
      encodeUpdate(Uint8Array.of(3)),
      Uint8Array.of(1, 1, 0),
      Uint8Array.of(3)
    );

    expect(decodeMessages(frame)).toEqual([
      { type: 'sync-step-1', stateVector: Uint8Array.of(1) },
      { type: 'sync-step-2', update: Uint8Array.of(2) },
      { type: 'update', update: Uint8Array.of(3) },
      { type: 'awareness', update: Uint8Array.of(0) },
      { type: 'query-awareness' },
    ]);
  });

  it('decodes permission-denied auth messages', () => {
    expect(decodeMessages(Uint8Array.of(2, 0, 6, 100, 101, 110, 105, 101, 100))).toEqual([
      { type: 'auth', reason: 'denied' },
    ]);
  });

  it('copies decoded payloads from the physical frame', () => {
    const frame = encodeUpdate(Uint8Array.of(4, 5));
    const messages = decodeMessages(frame);
    frame.fill(9);
    expect(messages).toEqual([{ type: 'update', update: Uint8Array.of(4, 5) }]);
  });

  it('rejects empty and truncated input', () => {
    const frames = [
      Uint8Array.of(),
      Uint8Array.of(0),
      Uint8Array.of(0, 2),
      Uint8Array.of(0, 2, 2, 1),
      Uint8Array.of(1, 2, 0),
      Uint8Array.of(2, 0, 1),
    ];
    for (const frame of frames) expect(() => decodeMessages(frame)).toThrow(ProtocolError);
  });

  it('rejects malformed awareness payloads and bounded state floods', () => {
    expect(() => decodeAwarenessUpdate(Uint8Array.of())).toThrow('Truncated varUint');
    expect(() => decodeAwarenessUpdate(Uint8Array.of(1, 1, 0, 2, 123, 125))).toThrow(
      'Awareness state must contain a user'
    );
    expect(() => decodeAwarenessUpdate(Uint8Array.of(0, 1))).toThrow(
      'Trailing awareness update data'
    );
    expect(() => decodeAwarenessUpdate(Uint8Array.of(2), 1)).toThrow(
      'Awareness update exceeds 1 states'
    );
  });

  it('rejects overflowing and non-canonical varUints', () => {
    const overflow = Uint8Array.from([0, 2, ...new Array(7).fill(0xff), 0x10]);
    expect(() => decodeMessages(overflow)).toThrow('varUint exceeds Number.MAX_SAFE_INTEGER');
    expect(() => decodeMessages(Uint8Array.of(0, 2, 0x80, 0))).toThrow(
      'Non-canonical varUint'
    );
  });

  it('rejects unknown top-level, sync, and auth types', () => {
    expect(() => decodeMessages(Uint8Array.of(4))).toThrow('Unknown top-level message type 4');
    expect(() => decodeMessages(Uint8Array.of(0, 3, 0))).toThrow(
      'Unknown sync message type 3'
    );
    expect(() => decodeMessages(Uint8Array.of(2, 1))).toThrow('Unknown auth message type 1');
  });

  it('enforces the physical frame limit', () => {
    expect(() => decodeMessages(encodeUpdate(Uint8Array.of(1)), 3)).toThrow(
      'Frame exceeds 3 bytes'
    );
    expect(() => encodeUpdate(Uint8Array.of(1, 2), 4)).toThrow('Frame exceeds 4 bytes');
  });

  it('bounds concatenated message counts independently of frame bytes', () => {
    const flood = new Uint8Array(5).fill(3);
    expect(() => decodeMessages(flood, 1024, 4)).toThrow('Frame exceeds 4 messages');
    expect(decodeMessages(flood.subarray(0, 4), 1024, 4)).toHaveLength(4);
    expect(() => decodeMessages(Uint8Array.of(3), 1024, 0)).toThrow(
      'Message limit must be a positive integer'
    );
  });
});
