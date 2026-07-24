import { describe, expect, it } from 'bun:test';
import {
  decodeAwarenessUpdate,
  decodeMessages,
  encodeAwarenessMessage,
  encodeAwarenessUpdate,
  encodeEmptyAwarenessUpdate,
  encodeQueryAwareness,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
  encodeVarUint,
  ProtocolError,
} from './protocol';

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

  it('round-trips sticky cursor awareness and leave states', () => {
    const states = [
      {
        clientId: 17,
        clock: 4,
        state: {
          user: { name: 'Calm Otter', color: '#0B57D0' },
          cursor: {
            story: 'body',
            anchor: Uint8Array.of(1, 2, 3),
            head: Uint8Array.of(4, 5),
          },
        },
      },
      { clientId: 18, clock: 9, state: null },
    ];
    const [message] = decodeMessages(encodeAwarenessMessage(states));
    expect(message.type).toBe('awareness');
    if (message.type !== 'awareness') return;
    expect(decodeAwarenessUpdate(message.update)).toEqual(states);
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

  it('rejects malformed awareness payloads', () => {
    expect(() => decodeAwarenessUpdate(Uint8Array.of())).toThrow('Truncated varUint');
    expect(() => decodeAwarenessUpdate(Uint8Array.of(1, 1, 1, 1, 123))).toThrow(
      'Invalid awareness JSON'
    );
    expect(() =>
      decodeAwarenessUpdate(
        encodeAwarenessUpdate([
          {
            clientId: 1,
            clock: 1,
            state: {
              user: { name: 'A', color: '#000000' },
              cursor: null,
            },
          },
        ]).slice(0, -1)
      )
    ).toThrow(ProtocolError);
  });

  it('falls back to the client palette for unsafe awareness colors', () => {
    const update = encodeAwarenessUpdate([
      {
        clientId: 12,
        clock: 1,
        state: {
          user: { name: 'Bright Fox', color: 'url("https://attacker/pixel")' },
          cursor: null,
        },
      },
    ]);

    expect(decodeAwarenessUpdate(update)[0]?.state?.user.color).toBe('#A142F4');
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
