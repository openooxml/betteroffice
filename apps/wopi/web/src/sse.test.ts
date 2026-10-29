import { describe, expect, test } from "bun:test";

import { SseParser, toSessionEvent } from "./sse";

describe("SseParser", () => {
  test("parses events split across chunks and ignores comments", () => {
    const parser = new SseParser();
    expect(parser.feed(": connected\n\nevent: saved\nda")).toEqual([]);
    const events = parser.feed('ta: {"itemVersion":"2"}\n\n');
    expect(events).toEqual([{ event: "saved", data: '{"itemVersion":"2"}' }]);
  });
});

describe("toSessionEvent", () => {
  test("maps effective writability snapshots and rejects snapshots without a boolean permission", () => {
    expect(toSessionEvent({ event: "snapshot", data: '{"canWrite":false,"viewOnlyReason":"lockLost"}' })).toEqual({
      type: "snapshot", canWrite: false, viewOnlyReason: "lockLost",
    });
    expect(toSessionEvent({ event: "snapshot", data: '{"canWrite":true,"viewOnlyReason":null}' })).toEqual({
      type: "snapshot", canWrite: true, viewOnlyReason: null,
    });
    expect(toSessionEvent({ event: "snapshot", data: '{"viewOnlyReason":"lockLost"}' })).toBeNull();
  });

  test("recognizes a replacement baseline event", () => {
    expect(toSessionEvent({ event: "rejoinRequired", data: '{"reason":"sourceChanged"}' })).toEqual({
      type: "rejoinRequired", reason: "sourceChanged",
    });
  });

  test("maps the session event kinds", () => {
    expect(toSessionEvent({ event: "saved", data: '{"itemVersion":"3"}' })).toEqual({ type: "saved", itemVersion: "3" });
    expect(toSessionEvent({ event: "conflict", data: '{"copyName":"book (conflict).xlsx","name":"book.xlsx"}' })).toEqual({ type: "conflict", copyName: "book (conflict).xlsx" });
    expect(toSessionEvent({ event: "viewOnly", data: '{"reason":"lockedElsewhere"}' })).toEqual({ type: "viewOnly", reason: "lockedElsewhere" });
    expect(toSessionEvent({ event: "writable", data: '{"canWrite":true}' })).toEqual({ type: "writable", canWrite: true });
    expect(toSessionEvent({ event: "writable", data: '{"canWrite":false}' })).toEqual({ type: "writable", canWrite: false });
    expect(toSessionEvent({ event: "writable", data: '{}' })).toBeNull();
    expect(toSessionEvent({ event: "tokenExpiring", data: '{"expiresAt":123}' })).toEqual({ type: "tokenExpiring", expiresAt: 123 });
    expect(toSessionEvent({ event: "leaderChanged", data: "{}" })).toBeNull();
    expect(toSessionEvent({ event: "saved", data: "nope" })).toBeNull();
  });
});
