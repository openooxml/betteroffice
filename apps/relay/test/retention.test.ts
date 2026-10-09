import { describe, expect, mock, test } from "bun:test";
import { MAX_AWARENESS_PAYLOAD_BYTES } from "../../../shared/collaboration-limits";
import { decodeMessages } from "../../../packages/docx/src/collaboration/protocol";
import * as Y from "yjs";
import { documentFrame, rehydrate } from "./fixtures";
import { RetainedUpdateLog, classifyFrame } from "../src/retention";

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

function authFrame(reason: Uint8Array): Uint8Array {
  return frame(
    encodeVarUint(2),
    encodeVarUint(0),
    encodeVarUint(reason.byteLength),
    reason,
  );
}

function documentMessages(protocolFrame: Uint8Array) {
  return decodeMessages(protocolFrame).filter(
    ({ type }) => type === "sync-step-2" || type === "update",
  );
}

describe("RetainedUpdateLog", () => {
  test("retains document messages and excludes transient traffic", () => {
    const document = documentFrame();
    const log = new RetainedUpdateLog(512, 1024);
    expect(log.retain(frame(document, awarenessFrame(Uint8Array.of(12))))?.puts[0].bytes).toEqual(document);
    expect(log.retain(awarenessFrame(Uint8Array.of(12)))).toBeNull();
    expect(log.retain(syncFrame(0, Uint8Array.of(0)))).toBeNull();
    expect(log.snapshot()).toEqual([document]);
  });

  test("checkpoints rather than discarding the oldest edits", () => {
    const log = new RetainedUpdateLog(2, 4096);
    const first = documentFrame('first', 1);
    const second = documentFrame('second', 2);
    log.retain(first);
    const mutation = log.retain(second)!;
    expect(mutation.checkpoint?.seq).toBe(1);
    expect(mutation.deletes).toEqual([0, 1]);
    expect(log.snapshot()).toHaveLength(1);
    const restored = rehydrate(log.snapshot());
    expect(restored.getText('body').toString()).toBe('firstsecond');
    restored.destroy();
  });

  test("restores a checkpoint plus its newer tail with monotonic sequence numbers", () => {
    const source = new RetainedUpdateLog(2, 4096);
    source.retain(documentFrame('first', 1));
    const checkpoint = source.retain(documentFrame('second', 2))!.checkpoint!;
    const third = documentFrame('third', 3);
    const restored = new RetainedUpdateLog(2, 4096);
    const repair = restored.restore([{ seq: 0, bytes: documentFrame('first', 1) }, { seq: 2, bytes: third }], checkpoint);
    expect(repair?.deletes).toEqual([0]);
    expect(restored.snapshot()).toEqual([checkpoint.bytes, third]);
    expect(restored.retain(documentFrame('fourth', 4))!.checkpoint?.seq).toBe(3);
    const doc = rehydrate(restored.snapshot());
    expect(doc.getText('body').toString()).toBe('firstsecondthirdfourth');
    doc.destroy();
  });

  test("rehydrates dependent edits received out of order", () => {
    const doc = new Y.Doc();
    const frames: Uint8Array[] = [];
    doc.on('update', update => frames.push(syncFrame(2, update)));
    doc.getText('body').insert(0, 'base');
    doc.getText('body').insert(4, ' tail');
    const log = new RetainedUpdateLog(1, 4096);
    log.retain(frames[1]);
    const checkpoint = log.checkpoint();
    expect(checkpoint).toBeNull();
    log.retain(frames[0]);
    const restored = rehydrate(log.snapshot());
    expect(restored.getText('body').toString()).toBe('base tail');
    restored.destroy();
    doc.destroy();
  });

  test("keeps offline edits mergeable after multiple checkpoints", () => {
    const first = rehydrate([documentFrame('base')]);
    const offline = rehydrate([documentFrame('base')]);
    const log = new RetainedUpdateLog(1, 4096);
    log.retain(documentFrame('base'));
    first.on('update', update => log.retain(syncFrame(2, update)));
    first.getText('body').insert(4, ' online');
    first.getText('body').delete(0, 1);
    offline.getText('body').insert(0, 'offline ');
    log.retain(syncFrame(2, Y.encodeStateAsUpdate(offline)));
    const restored = rehydrate(log.snapshot());
    Y.applyUpdate(first, Y.encodeStateAsUpdate(offline));
    expect(restored.getText('body').toString()).toBe(first.getText('body').toString());
    for (const doc of [first, offline, restored]) doc.destroy();
  });

  test("compacts duplicate traffic under the byte budget", () => {
    const frame = documentFrame('a'.repeat(200));
    const log = new RetainedUpdateLog(512, 512);
    for (let i = 0; i < 20; i++) log.retain(frame);
    expect(log.snapshot().reduce((sum, bytes) => sum + bytes.length, 0)).toBeLessThanOrEqual(512);
    expect(rehydrate(log.snapshot()).getText('body').toString()).toBe('a'.repeat(200));
  });

  test("refuses capacity overflow without losing the previous state", () => {
    const log = new RetainedUpdateLog(1, 256);
    log.retain(documentFrame('a'.repeat(100), 1));
    const before = log.snapshot();
    expect(() => log.retain(documentFrame('b'.repeat(200), 2))).toThrow('capacity');
    expect(log.snapshot()).toEqual(before);
  });

  test("rejects malformed document payloads without adopting them", () => {
    const log = new RetainedUpdateLog(2, 4096);
    log.retain(documentFrame());
    expect(() => log.retain(syncFrame(2, Uint8Array.of(255)))).toThrow();
    expect(log.snapshot()).toEqual([documentFrame()]);
    expect(log.restore([{ seq: 0, bytes: syncFrame(2, Uint8Array.of(255)) }])?.deletes).toEqual([0]);
  });

  test("answers state vectors and requests the seed missing from a fresh room", () => {
    const log = new RetainedUpdateLog(2, 4096);
    expect(decodeMessages(log.syncRequest())).toEqual([{ type: 'sync-step-1', stateVector: Uint8Array.of(0) }]);
    log.retain(documentFrame());
    const answers = log.responses(syncFrame(0, Uint8Array.of(0)));
    expect(rehydrate(answers).getText('body').toString()).toBe('hello');
    const peer = rehydrate(answers);
    const [response] = log.responses(syncFrame(0, Y.encodeStateVector(peer)));
    expect(decodeMessages(response)).toEqual([{ type: 'sync-step-2', update: Uint8Array.of(0, 0) }]);
    peer.destroy();
  });

  test("sync responses include edits after querying cached state", () => {
    const log = new RetainedUpdateLog(2, 4096);
    const query = syncFrame(0, Uint8Array.of(0));
    log.retain(documentFrame('first', 1));
    expect(rehydrate(log.responses(query)).getText('body').toString()).toBe('first');
    expect(rehydrate(log.responses(query)).getText('body').toString()).toBe('first');
    log.retain(documentFrame('second', 2));
    expect(rehydrate(log.responses(query)).getText('body').toString()).toBe('firstsecond');
    log.retain(documentFrame('third', 3));
    expect(rehydrate(log.responses(query)).getText('body').toString()).toBe('firstsecondthird');
  });

  test("clears both checkpoint and tail", () => {
    const log = new RetainedUpdateLog(1, 4096);
    log.retain(documentFrame());
    log.clear();
    expect(log.snapshot()).toEqual([]);
    expect(log.retain(documentFrame())!.checkpoint?.seq).toBe(0);
  });
});

describe("classifyFrame", () => {
  test("reports frames carrying document state as document", () => {
    const document = syncFrame(2, Uint8Array.of(1));
    expect(classifyFrame(document).kind).toBe("document");
    expect(
      classifyFrame(frame(document, awarenessFrame(Uint8Array.of(2)))).kind,
    ).toBe("document");
  });

  test("reports valid but unretained frames as transient", () => {
    expect(classifyFrame(awarenessFrame(Uint8Array.of(2))).kind).toBe("transient");
    expect(classifyFrame(encodeVarUint(3)).kind).toBe("transient");
    expect(classifyFrame(syncFrame(0, Uint8Array.of(3))).kind).toBe("transient");
  });

  test("reports auth-bearing frames as auth even alongside sync", () => {
    const denial = authFrame(Uint8Array.of(104, 105));
    expect(classifyFrame(denial).kind).toBe("auth");
    expect(classifyFrame(frame(syncFrame(2, Uint8Array.of(1)), denial)).kind).toBe(
      "auth",
    );
  });

  test("reports truncated or unknown frames as invalid", () => {
    expect(classifyFrame(new Uint8Array()).kind).toBe("invalid");
    expect(classifyFrame(Uint8Array.of(0)).kind).toBe("invalid");
    expect(classifyFrame(Uint8Array.of(0x80)).kind).toBe("invalid");
    expect(classifyFrame(Uint8Array.of(0x80, 0)).kind).toBe("invalid");
    expect(classifyFrame(Uint8Array.of(0, 3, 0)).kind).toBe("invalid");
    expect(classifyFrame(encodeVarUint(128)).kind).toBe("invalid");
    expect(classifyFrame(authFrame(Uint8Array.of(0xff))).kind).toBe("invalid");
  });

  test("accepts awareness at the cap and flags larger payloads", () => {
    const atLimit = awarenessFrame(new Uint8Array(MAX_AWARENESS_PAYLOAD_BYTES));
    expect(classifyFrame(atLimit)).toEqual({
      kind: "transient", hasAwareness: true, awarenessBytes: MAX_AWARENESS_PAYLOAD_BYTES,
    });
    const oversize = awarenessFrame(new Uint8Array(MAX_AWARENESS_PAYLOAD_BYTES + 1));
    expect(classifyFrame(oversize).kind).toBe("oversize-awareness");
    expect(classifyFrame(frame(documentFrame(), oversize)).kind).toBe("oversize-awareness");
  });

  test("caps the sum of awareness payloads within mixed frames", () => {
    const half = awarenessFrame(new Uint8Array(MAX_AWARENESS_PAYLOAD_BYTES / 2));
    expect(classifyFrame(frame(documentFrame(), half, half))).toEqual({
      kind: "document", hasAwareness: true, awarenessBytes: MAX_AWARENESS_PAYLOAD_BYTES,
    });
    expect(classifyFrame(frame(documentFrame(), half, half, awarenessFrame(Uint8Array.of(0)))).kind)
      .toBe("oversize-awareness");
  });

  test("identifies empty awareness payloads without marking document-only frames", () => {
    expect(classifyFrame(awarenessFrame(new Uint8Array()))).toEqual({
      kind: "transient", hasAwareness: true, awarenessBytes: 0,
    });
    expect(classifyFrame(documentFrame())).toEqual({
      kind: "document", hasAwareness: false, awarenessBytes: 0,
    });
  });
});
