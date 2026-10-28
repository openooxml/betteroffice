import { expect, test } from "bun:test";
import { decodeSaveBarrier } from "../sessionApi";
import { decodeSequencedFrame, encodeSequencedFrame } from "./sequence";

test("sequence envelopes preserve the collaboration payload", () => {
  const barrier = { connectionId: crypto.randomUUID(), seq: Number.MAX_SAFE_INTEGER };
  const payload = Uint8Array.of(0, 2, 2, 0, 0);
  expect(decodeSequencedFrame(encodeSequencedFrame(barrier, payload))).toEqual({ ...barrier, payload });
  expect(decodeSequencedFrame(encodeSequencedFrame({ ...barrier, seq: 0 }))).toEqual({
    ...barrier, seq: 0, payload: new Uint8Array(),
  });
  expect(decodeSequencedFrame(payload)).toBeNull();
});

test("save barriers validate the connection id and sequence", () => {
  const barrier = { connectionId: crypto.randomUUID(), seq: 3 };
  const body = (value: unknown) => new TextEncoder().encode(JSON.stringify(value));
  expect(decodeSaveBarrier(body(barrier))).toEqual(barrier);
  for (const value of [
    { ...barrier, seq: -1 }, { ...barrier, seq: 0.5 }, { ...barrier, seq: Number.MAX_SAFE_INTEGER + 1 },
    { ...barrier, seq: "3" }, { ...barrier, connectionId: "" }, { connectionId: barrier.connectionId },
    { seq: 3, stateVector: "AA==" }, { ...barrier, connectionId: "invalid", stateVector: "AA==" },
  ]) {
    expect(() => decodeSaveBarrier(body(value))).toThrow("invalidSaveBarrier");
  }
});

test("malformed sequence envelopes are rejected", () => {
  const frame = encodeSequencedFrame({ connectionId: crypto.randomUUID(), seq: 1 }, Uint8Array.of(0));
  expect(() => decodeSequencedFrame(frame.slice(0, 4))).toThrow();
  frame[1] = 0;
  expect(() => decodeSequencedFrame(frame)).toThrow();
});
