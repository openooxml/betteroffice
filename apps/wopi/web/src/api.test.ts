import { describe, expect, test } from "bun:test";

import { SessionApi, decodeJoinResponse } from "./api";
import type { SessionEvent } from "./sse";

function sseResponse(text: string): Response {
  return new Response(text, { status: 200, headers: { "Content-Type": "text/event-stream" } });
}

describe("decodeJoinResponse", () => {
  test("decodes the base64 initial update", () => {
    const decoded = decodeJoinResponse({ format: "xlsx", initialUpdate: Buffer.from([1, 2, 3]).toString("base64") });
    expect(decoded.format).toBe("xlsx");
    expect(decoded.initialUpdate).toEqual(Uint8Array.of(1, 2, 3));
  });
});

describe("SessionApi default fetch", () => {
  test("calls the global fetch with the global this binding", async () => {
    const original = globalThis.fetch;
    function fake(this: unknown, _input: RequestInfo | URL, _init?: RequestInit): Promise<Response> {
      if (this !== globalThis) throw new TypeError("Failed to execute 'fetch' on 'Window': Illegal invocation");
      return Promise.resolve(new Response(new Uint8Array([1, 2, 3])));
    }
    globalThis.fetch = fake as unknown as typeof fetch;
    try {
      const api = new SessionApi("handle-1");
      expect(await api.source()).toEqual(Uint8Array.of(1, 2, 3));
    } finally {
      globalThis.fetch = original;
    }
  });
});

describe("SessionApi events", () => {
  test("reconnects after a dropped stream until aborted", async () => {
    let calls = 0;
    const fetchFn = (async () => {
      calls += 1;
      if (calls === 1) return sseResponse('event: saved\ndata: {"itemVersion":"1"}\n\n');
      return sseResponse('event: saved\ndata: {"itemVersion":"2"}\n\n');
    }) as unknown as typeof fetch;
    const api = new SessionApi("handle-1", fetchFn);
    const controller = new AbortController();
    const seen: SessionEvent[] = [];
    const done = api.events(controller.signal, (event) => {
      seen.push(event);
      if (seen.length === 2) controller.abort();
    }, { retryMs: 1 });
    await done;
    expect(calls).toBeGreaterThanOrEqual(2);
    expect(seen).toEqual([
      { type: "saved", itemVersion: "1" },
      { type: "saved", itemVersion: "2" },
    ]);
  });

  test("stops reconnecting after a closed event", async () => {
    let calls = 0;
    const fetchFn = (async () => {
      calls += 1;
      return sseResponse("event: closed\ndata: {}\n\n");
    }) as unknown as typeof fetch;
    const api = new SessionApi("handle-1", fetchFn);
    const controller = new AbortController();
    await api.events(controller.signal, () => {}, { retryMs: 1 });
    expect(calls).toBe(1);
  });
});

test("save posts the client state vector as base64 JSON", async () => {
  let request: RequestInit | undefined;
  const api = new SessionApi("handle-1", (async (_input: RequestInfo | URL, init?: RequestInit) => {
    request = init;
    return Response.json({ itemVersion: "2" });
  }) as unknown as typeof fetch);
  expect(await api.save(Uint8Array.of(1, 2, 255))).toEqual({ itemVersion: "2" });
  expect(request?.method).toBe("POST");
  expect(new Headers(request?.headers).get("Content-Type")).toBe("application/json");
  expect(JSON.parse(String(request?.body))).toEqual({ stateVector: "AQL/" });
});

test("save exposes a short typed synchronization failure", async () => {
  const api = new SessionApi("handle-1", (async () => Response.json({ error: { kind: "notSynced" } }, { status: 409 })) as unknown as typeof fetch);
  await expect(api.save(Uint8Array.of(0))).rejects.toThrow("notSynced");
});

test("save forwards the request cancellation signal", async () => {
  const controller = new AbortController();
  let signal: AbortSignal | null | undefined;
  const api = new SessionApi("handle-1", (async (_input: RequestInfo | URL, init?: RequestInit) => {
    signal = init?.signal;
    return Response.json({ itemVersion: "2" });
  }) as unknown as typeof fetch);
  await api.save(Uint8Array.of(0), controller.signal);
  expect(signal).toBe(controller.signal);
});

test("close rejects unsuccessful responses", async () => {
  const api = new SessionApi("handle-1", (async () => new Response(null, { status: 502 })) as unknown as typeof fetch);
  await expect(api.close()).rejects.toThrow("Close failed");
});
