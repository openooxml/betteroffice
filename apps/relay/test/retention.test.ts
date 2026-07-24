import { describe, expect, test } from "bun:test";
import { RetainedUpdateLog, isDocumentFrame } from "../src/retention";

function encodeVarUint(value: number): Uint8Array {
  const bytes: number[] = [];
  let remaining = value;
  while (remaining >= 128) {
    bytes.push((remaining % 128) | 0x80);
    remaining = Math.floor(remaining / 128);
  }
  bytes.push(remaining);
  return Uint8Array.from(bytes);
}

function frame(...parts: readonly Uint8Array[]): Uint8Array {
  const bytes = new Uint8Array(
    parts.reduce((length, part) => length + part.byteLength, 0),
  );
  let offset = 0;
  for (const part of parts) {
    bytes.set(part, offset);
    offset += part.byteLength;
  }
  return bytes;
}

function syncFrame(subtype: number, payload: Uint8Array): Uint8Array {
  return frame(
    encodeVarUint(0),
    encodeVarUint(subtype),
    encodeVarUint(payload.byteLength),
    payload,
  );
}

function awarenessFrame(payload: Uint8Array): Uint8Array {
  return frame(
    encodeVarUint(1),
    encodeVarUint(payload.byteLength),
    payload,
  );
}

describe("RetainedUpdateLog", () => {
  test("retains and replays only document frames", () => {
    const first = syncFrame(1, Uint8Array.of(10, 11));
    const second = syncFrame(2, Uint8Array.of(12));
    const awareness = awarenessFrame(Uint8Array.of(13));
    const queryAwareness = encodeVarUint(3);
    const malformed = Uint8Array.of(0, 2, 2, 14);
    const log = new RetainedUpdateLog(512, 1024);

    expect(log.retain(first)).toBe(true);
    expect(log.retain(awareness)).toBe(false);
    expect(log.retain(queryAwareness)).toBe(false);
    expect(log.retain(malformed)).toBe(false);
    expect(log.retain(second)).toBe(true);
    expect(log.snapshot()).toEqual([first, second]);

    const replayed: Uint8Array[] = [];
    log.replay((update) => replayed.push(update));
    expect(replayed).toEqual([first, second]);
  });

  test("filters awareness and malformed entries from stored logs", () => {
    const document = syncFrame(0, Uint8Array.of(20));
    const awareness = awarenessFrame(Uint8Array.of(21));
    const log = new RetainedUpdateLog(512, 1024);

    expect(
      log.restore([awareness, Uint8Array.of(0x80), document]),
    ).toBe(true);
    expect(log.snapshot()).toEqual([document]);
  });
});

describe("isDocumentFrame", () => {
  test("validates varints and complete sync payloads", () => {
    expect(isDocumentFrame(syncFrame(2, Uint8Array.of(1)))).toBe(true);
    expect(isDocumentFrame(new Uint8Array())).toBe(false);
    expect(isDocumentFrame(Uint8Array.of(0))).toBe(false);
    expect(isDocumentFrame(Uint8Array.of(0x80))).toBe(false);
    expect(isDocumentFrame(Uint8Array.of(0x80, 0))).toBe(false);
    expect(isDocumentFrame(Uint8Array.of(0, 3, 0))).toBe(false);
    expect(isDocumentFrame(encodeVarUint(128))).toBe(false);
  });
});
