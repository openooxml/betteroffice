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
  test("maps the session event kinds", () => {
    expect(toSessionEvent({ event: "saved", data: '{"itemVersion":"3"}' })).toEqual({ type: "saved", itemVersion: "3" });
    expect(toSessionEvent({ event: "conflict", data: '{"copyName":"book (conflict).xlsx","name":"book.xlsx"}' })).toEqual({ type: "conflict", copyName: "book (conflict).xlsx" });
    expect(toSessionEvent({ event: "viewOnly", data: '{"reason":"lockedElsewhere"}' })).toEqual({ type: "viewOnly", reason: "lockedElsewhere" });
    expect(toSessionEvent({ event: "tokenExpiring", data: '{"expiresAt":123}' })).toEqual({ type: "tokenExpiring", expiresAt: 123 });
    expect(toSessionEvent({ event: "leaderChanged", data: "{}" })).toBeNull();
    expect(toSessionEvent({ event: "saved", data: "nope" })).toBeNull();
  });
});
