import { describe, expect, mock, test } from "bun:test";
import { decodeMessages } from "../../../packages/docx/src/collaboration/protocol";
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

function documentMessages(protocolFrame: Uint8Array) {
  return decodeMessages(protocolFrame).filter(
    ({ type }) =>
      type === "sync-step-1" ||
      type === "sync-step-2" ||
      type === "update",
  );
}

describe("RetainedUpdateLog", () => {
  test("retains sync from a mixed awareness frame and replays it", () => {
    const document = syncFrame(2, Uint8Array.of(10, 11));
    const mixed = frame(document, awarenessFrame(Uint8Array.of(12)));
    const log = new RetainedUpdateLog(512, 1024);

    expect(log.retain(mixed)).toBe(true);
    const [retained] = log.snapshot();
    expect(retained).toEqual(document);
    expect(decodeMessages(retained)).toEqual(documentMessages(mixed));

    const newSocket = { send: mock((_: Uint8Array) => {}) };
    log.replay((update) => newSocket.send(update));
    expect(newSocket.send).toHaveBeenCalledTimes(1);
    const replayed = newSocket.send.mock.calls[0][0];
    expect(replayed).toEqual(document);
    expect(decodeMessages(replayed)).toEqual(documentMessages(mixed));
  });

  test("does not retain awareness-only frames", () => {
    const log = new RetainedUpdateLog(512, 1024);

    expect(log.retain(awarenessFrame(Uint8Array.of(13)))).toBe(false);
    expect(log.snapshot()).toEqual([]);
  });

  test("keeps sync-only frames byte-identical in a fresh buffer", () => {
    const document = syncFrame(1, new Uint8Array(130).fill(14));
    const log = new RetainedUpdateLog(512, 1024);

    expect(log.retain(document)).toBe(true);
    const [retained] = log.snapshot();
    expect(retained).toEqual(document);
    expect(retained).not.toBe(document);
    expect(decodeMessages(retained)).toEqual(decodeMessages(document));
  });

  test("retains sync from a mixed query-awareness frame", () => {
    const document = syncFrame(0, Uint8Array.of(15));
    const mixed = frame(encodeVarUint(3), document);
    const log = new RetainedUpdateLog(512, 1024);

    expect(log.retain(mixed)).toBe(true);
    const [retained] = log.snapshot();
    expect(retained).toEqual(document);
    expect(decodeMessages(retained)).toEqual(documentMessages(mixed));
  });

  test("retains multiple sync messages in their original order", () => {
    const first = syncFrame(0, Uint8Array.of(16));
    const second = syncFrame(2, Uint8Array.of(17, 18));
    const mixed = frame(
      first,
      awarenessFrame(Uint8Array.of(19)),
      second,
    );
    const expected = frame(first, second);
    const log = new RetainedUpdateLog(512, 1024);

    expect(log.retain(mixed)).toBe(true);
    const [retained] = log.snapshot();
    expect(retained).toEqual(expected);
    expect(decodeMessages(retained)).toEqual(documentMessages(mixed));
  });

  test("rejects malformed frames without throwing or retaining", () => {
    const document = syncFrame(2, Uint8Array.of(20));
    const frames = [
      frame(document, Uint8Array.of(1, 2, 21)),
      frame(document, Uint8Array.of(4)),
      frame(document, Uint8Array.of(0x80, 0)),
    ];
    const log = new RetainedUpdateLog(512, 1024);

    for (const malformed of frames) {
      let retained = true;
      expect(() => {
        retained = log.retain(malformed);
      }).not.toThrow();
      expect(retained).toBe(false);
    }
    expect(log.snapshot()).toEqual([]);
  });

  test("normalizes mixed stored frames and removes invalid entries", () => {
    const document = syncFrame(0, Uint8Array.of(22));
    const awareness = awarenessFrame(Uint8Array.of(23));
    const mixed = frame(document, awareness);
    const log = new RetainedUpdateLog(512, 1024);

    expect(
      log.restore([awareness, Uint8Array.of(0x80), mixed]),
    ).toBe(true);
    expect(log.snapshot()).toEqual([document]);
  });
});

describe("isDocumentFrame", () => {
  test("accepts only complete, canonical sync messages", () => {
    const document = syncFrame(2, Uint8Array.of(1));
    expect(isDocumentFrame(document)).toBe(true);
    expect(
      isDocumentFrame(frame(document, awarenessFrame(Uint8Array.of(2)))),
    ).toBe(false);
    expect(isDocumentFrame(new Uint8Array())).toBe(false);
    expect(isDocumentFrame(Uint8Array.of(0))).toBe(false);
    expect(isDocumentFrame(Uint8Array.of(0x80))).toBe(false);
    expect(isDocumentFrame(Uint8Array.of(0x80, 0))).toBe(false);
    expect(isDocumentFrame(Uint8Array.of(0, 3, 0))).toBe(false);
    expect(isDocumentFrame(encodeVarUint(128))).toBe(false);
  });
});
