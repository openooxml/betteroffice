import { describe, expect, test } from "bun:test";

import { SessionApi, SessionEndedError, decodeJoinResponse } from "./api";
import type { SessionEvent } from "./sse";

function sseResponse(text: string): Response {
  return new Response(text, { status: 200, headers: { "Content-Type": "text/event-stream" } });
}

describe("decodeJoinResponse", () => {
  test.each([false, true])("decodes the base64 initial update and effective canWrite=%s", (canWrite) => {
    const decoded = decodeJoinResponse({ format: "xlsx", baselineGeneration: 1, initialUpdate: Buffer.from([1, 2, 3]).toString("base64"), canWrite });
    expect(decoded.format).toBe("xlsx");
    expect(decoded.initialUpdate).toEqual(Uint8Array.of(1, 2, 3));
    expect(decoded.baselineGeneration).toBe(1);
    expect(decoded.canWrite).toBe(canWrite);
  });

  test.each([undefined, null, "true", 1])("rejects invalid effective writability %s", (canWrite) => {
    expect(() => decodeJoinResponse({ format: "xlsx", baselineGeneration: 1, initialUpdate: "AA==", canWrite })).toThrow("Invalid join response writability.");
  });
});

describe("SessionApi default fetch", () => {
  test("calls the global fetch with the global this binding", async () => {
    const original = globalThis.fetch;
    function fake(this: unknown, _input: RequestInfo | URL, _init?: RequestInit): Promise<Response> {
      if (this !== globalThis) throw new TypeError("Failed to execute 'fetch' on 'Window': Illegal invocation");
      return Promise.resolve(new Response(new Uint8Array([1, 2, 3]), { headers: { "X-WOPI-Baseline-Generation": "1" } }));
    }
    globalThis.fetch = fake as unknown as typeof fetch;
    try {
      const api = new SessionApi("handle-1");
      expect(await api.source()).toEqual({ bytes: Uint8Array.of(1, 2, 3), baselineGeneration: 1 });
    } finally {
      globalThis.fetch = original;
    }
  });
});

describe("SessionApi events", () => {
  test("delivers rejoinRequired once and leaves stream replacement to the page", async () => {
    let calls = 0;
    const api = new SessionApi("handle-1", async (input, init) => {
      calls += 1;
      expect(String(input)).toBe("/api/session/events");
      expect(new Headers(init?.headers).get("X-WOPI-Baseline-Generation")).toBe("1");
      expect(new Headers(init?.headers).get("Authorization")).toBe("Bearer handle-1");
      return sseResponse('event: rejoinRequired\ndata: {"reason":"sourceChanged"}\n\nevent: saved\ndata: {}\n\n');
    });
    const controller = new AbortController();
    const seen: SessionEvent[] = [];
    const timeout = setTimeout(() => controller.abort(), 50);
    try {
      await api.events(controller.signal, (event) => { seen.push(event); }, { baselineGeneration: 1, retryMs: 1 });
      expect(calls).toBe(1);
      expect(seen).toEqual([{ type: "rejoinRequired", reason: "sourceChanged" }]);
    } finally {
      clearTimeout(timeout);
      controller.abort();
    }
  });

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
    }, { baselineGeneration: 1, retryMs: 1 });
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
    await api.events(controller.signal, () => {}, { baselineGeneration: 1, retryMs: 1 });
    expect(calls).toBe(1);
  });
});

test("save posts the captured socket sequence without a state vector", async () => {
  let request: RequestInit | undefined;
  const api = new SessionApi("handle-1", (async (_input: RequestInfo | URL, init?: RequestInit) => {
    request = init;
    return Response.json({ itemVersion: "2" });
  }) as unknown as typeof fetch);
  const barrier = { connectionId: crypto.randomUUID(), seq: 7 };
  await api.save(barrier);
  expect(JSON.parse(String(request?.body))).toEqual(barrier);
});

test("save exposes a short typed synchronization failure", async () => {
  const api = new SessionApi("handle-1", (async () => Response.json({ error: { kind: "notSynced" } }, { status: 409 })) as unknown as typeof fetch);
  await expect(api.save({ connectionId: crypto.randomUUID(), seq: 0 })).rejects.toThrow("notSynced");
});

test("save forwards the request cancellation signal", async () => {
  const controller = new AbortController();
  let signal: AbortSignal | null | undefined;
  const api = new SessionApi("handle-1", (async (_input: RequestInfo | URL, init?: RequestInit) => {
    signal = init?.signal;
    return Response.json({ itemVersion: "2" });
  }) as unknown as typeof fetch);
  await api.save({ connectionId: crypto.randomUUID(), seq: 0 }, controller.signal);
  expect(signal).toBe(controller.signal);
});

test("close rejects unsuccessful responses", async () => {
  const api = new SessionApi("handle-1", (async () => new Response(null, { status: 502 })) as unknown as typeof fetch);
  await expect(api.close()).rejects.toThrow("Close failed");
});

test("close forwards the request cancellation signal", async () => {
  const controller = new AbortController();
  let request: RequestInit | undefined;
  const api = new SessionApi("handle-1", async (input, init) => {
    expect(String(input)).toBe("/api/session/close");
    request = init;
    return Response.json({ closed: true });
  });
  await api.close(controller.signal);
  expect(request?.signal).toBe(controller.signal);
  expect(request?.method).toBe("POST");
  expect(new Headers(request?.headers).get("Authorization")).toBe("Bearer handle-1");
});

test.each([401, 404, 410])("ended event streams stop retrying with status %i", async (status) => {
  let calls = 0;
  const api = new SessionApi("handle", async () => { calls += 1; return new Response(null, { status }); });
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30);
  const seen: SessionEvent[] = [];
  try {
    await api.events(controller.signal, (event) => { seen.push(event); }, { baselineGeneration: 1, retryMs: 1 });
    expect(calls).toBe(1);
    expect(seen).toEqual([{ type: "closed" }]);
  } finally { clearTimeout(timer); controller.abort(); }
});

test("recovery posts serialized bytes with a Bearer handle and returns only the copy name", async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  const api = new SessionApi("handle", async (input, init) => {
    expect(String(input)).toBe("/api/session/recovery");
    expect(init?.method).toBe("POST");
    expect(new Headers(init?.headers).get("Authorization")).toBe("Bearer handle");
    expect(new Uint8Array(await new Response(init?.body).arrayBuffer())).toEqual(bytes);
    return Response.json({ copyName: "file (conflict recovery).xlsx", access_token: "discard" });
  });
  expect(await api.recover(bytes)).toEqual({ copyName: "file (conflict recovery).xlsx" });
});

test.each([401, 404, 410, 409])("save and recovery identify an ended handle with status %i", async (status) => {
  const api = new SessionApi("handle", async () => Response.json({ error: { kind: "closed" } }, { status }));
  await expect(api.save({ connectionId: crypto.randomUUID(), seq: 1 })).rejects.toBeInstanceOf(SessionEndedError);
  await expect(api.recover(Uint8Array.of(1))).rejects.toBeInstanceOf(SessionEndedError);
});
