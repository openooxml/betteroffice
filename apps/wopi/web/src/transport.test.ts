import { describe, expect, spyOn, test } from "bun:test";

import { collabUrl, createWopiTransport, type WebSocketLike } from "./transport";
import { decodeSequencedFrame } from "../../src/relay/sequence";
import { closeDocument, saveDocument } from "./documentActions";
import { CollaborationProvider } from "../../../../packages/xlsx/src/collaboration/provider";
import type { CollaborationReplica } from "../../../../packages/xlsx/src/collaboration/types";
import { applyUpdate, Doc, encodeStateAsUpdate, encodeStateVector } from "yjs";
import { classifyFrame, encodeSyncStep1, encodeSyncStep2 } from "../../src/relay/frames";

class FakeSocket implements WebSocketLike {
  binaryType = "";
  readyState = 0;
  bufferedAmount = 0;
  onopen: (() => void) | null = null;
  onmessage: ((event: { data: unknown }) => void) | null = null;
  onerror: ((event: unknown) => void) | null = null;
  onclose: ((event: { code?: number; reason?: string }) => void) | null = null;
  sent: ArrayBuffer[] = [];
  closedWith: Array<[number | undefined, string | undefined]> = [];

  send(data: ArrayBuffer): void {
    this.sent.push(data);
  }

  close(code?: number, reason?: string): void {
    this.closedWith.push([code, reason]);
    this.readyState = 3;
  }
}

test.each([false, true])("Close waits for an offline deletion to resync after reconnect, timeout=%s", async (timeout) => {
  const local = new Doc();
  const server = new Doc();
  local.getText("body").insert(0, "before");
  applyUpdate(server, encodeStateAsUpdate(local));
  const vector = encodeStateVector(local);
  const sockets: FakeSocket[] = [];
  const transport = createWopiTransport({
    handle: "h", baselineGeneration: 1, origin: "https://office.example",
    createSocket() {
      const socket = new FakeSocket();
      const send = socket.send.bind(socket);
      socket.send = (data) => {
        send(data);
        const frame = decodeSequencedFrame(new Uint8Array(data))!;
        if (frame.seq === 0) return;
        for (const message of classifyFrame(frame.payload).messages) {
          if (message.type === "sync-step-2" || message.type === "update") applyUpdate(server, message.update);
        }
      };
      sockets.push(socket);
      return socket;
    },
  });
  const provider = new CollaborationProvider({
    clientId: local.clientID,
    encodeStateVector: () => encodeStateVector(local),
    encodeStateAsUpdate: (vector) => encodeStateAsUpdate(local, vector),
    applyUpdate: (update) => { applyUpdate(local, update, "remote"); },
    onUpdate(listener) {
      const onUpdate = (update: Uint8Array, origin: unknown) => listener(update, origin === "remote" ? "remote" : "local");
      local.on("update", onUpdate);
      return () => local.off("update", onUpdate);
    },
  }, transport);
  const blocked: boolean[] = [];
  const calls: string[] = [];
  let hostText = "before";
  const warning = spyOn(console, "error").mockImplementation(() => {});
  const receive = (socket: FakeSocket, frame: Uint8Array) => socket.onmessage?.({ data: frame });
  try {
    provider.connect();
    const first = sockets[0]!;
    first.readyState = 1;
    first.onopen?.();
    receive(first, encodeSyncStep2(encodeStateAsUpdate(server, encodeStateVector(local))));
    receive(first, encodeSyncStep1(encodeStateVector(server)));
    expect(provider.synced).toBe(true);
    first.readyState = 3;
    first.onclose?.({ code: 1006 });
    local.getText("body").delete(0, 6);
    expect(encodeStateVector(local)).toEqual(vector);
    expect(server.getText("body").toString()).toBe("before");
    await new Promise((resolve) => setTimeout(resolve, 300));
    const second = sockets[1]!;
    second.readyState = 1;
    second.onopen?.();
    receive(second, encodeSyncStep2(encodeStateAsUpdate(server, encodeStateVector(local))));
    expect(provider.synced).toBe(true);
    expect(provider.pendingBytes).toBe(0);
    const closing = closeDocument({
      flushPendingInput: async () => {}, hasPendingUpdates: () => provider.pendingBytes > 0,
    }, {
      async save(barrier) {
        expect(barrier).toEqual(transport.captureSaveBarrier());
        hostText = server.getText("body").toString();
        calls.push("save");
      },
      async close() { calls.push("close"); },
    }, {
      inputTimeoutMs: timeout ? 5 : 500,
      hasPendingResync: () => transport.hasPendingResync(),
      captureBarrier: () => { calls.push("sequence"); return transport.captureSaveBarrier(); },
      setInputBlocked: (value) => { blocked.push(value); },
    }, () => { calls.push("host-close"); });
    const outcome = closing.then(() => "closed", (error: Error) => error.message);
    await new Promise((resolve) => setTimeout(resolve, timeout ? 15 : 0));
    expect(calls).toEqual([]);
    if (timeout) {
      expect(await outcome).toBe("notSynced");
      expect(blocked).toEqual([true, false]);
    } else {
      expect(blocked).toEqual([true]);
    }
    receive(second, encodeSyncStep1(encodeStateVector(server)));
    if (timeout) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(calls).toEqual([]);
    } else {
      expect(await outcome).toBe("closed");
      expect(hostText).toBe("");
      expect(calls).toEqual(["sequence", "save", "close", "host-close"]);
    }
  } finally {
    provider.destroy();
    transport.disconnect();
    local.destroy();
    server.destroy();
    warning.mockRestore();
  }
});

test("socket errors followed by close reconnect a real provider and sync local edits", async () => {
  const sockets: FakeSocket[] = [];
  const transport = createWopiTransport({ handle: "h", baselineGeneration: 1, origin: "https://office.example", createSocket() {
    const socket = new FakeSocket();
    sockets.push(socket);
    return socket;
  } });
  const doc = new Doc();
  const provider = new CollaborationProvider({
    clientId: doc.clientID, encodeStateVector: () => encodeStateVector(doc),
    encodeStateAsUpdate: (vector) => encodeStateAsUpdate(doc, vector),
    applyUpdate: (update) => { applyUpdate(doc, update); }, onUpdate: () => () => {},
  }, transport);
  const errors: Error[] = [];
  provider.onError((error) => { errors.push(error); });
  try {
    provider.connect();
    const first = sockets[0]!;
    first.onerror?.(new Error("network unavailable"));
    first.readyState = 3;
    first.onclose?.({ code: 1006 });
    expect(provider.status).toBe("disconnected");
    expect(errors).toEqual([]);
    doc.getText("body").insert(0, "offline edit");
    await new Promise((resolve) => setTimeout(resolve, 300));
    expect(sockets).toHaveLength(2);
    const second = sockets[1]!;
    second.readyState = 1;
    second.onopen?.();
    second.onmessage?.({ data: encodeSyncStep2(Uint8Array.of(0, 0)) });
    second.onmessage?.({ data: encodeSyncStep1(Uint8Array.of(0)) });
    expect(provider.synced).toBe(true);
    expect(provider.status).toBe("connected");
    const server = new Doc();
    try {
      for (const data of second.sent.slice(1)) {
        for (const message of classifyFrame(decodeSequencedFrame(new Uint8Array(data))!.payload).messages) {
          if (message.type === "sync-step-2" || message.type === "update") applyUpdate(server, message.update);
        }
      }
      expect(server.getText("body").toString()).toBe("offline edit");
      expect(errors).toEqual([]);
    } finally { server.destroy(); }
  } finally {
    provider.destroy();
    transport.disconnect();
    doc.destroy();
  }
});

test.each([false, true])("save waits for provider backpressure, timeout=%s", async (timeout) => {
  const socket = new FakeSocket();
  const transport = createWopiTransport({ handle: "h", baselineGeneration: 1, origin: "https://office.example", createSocket: () => socket });
  let update!: Parameters<CollaborationReplica["onUpdate"]>[0];
  const provider = new CollaborationProvider({
    clientId: 7, encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0),
    applyUpdate() {}, onUpdate(listener) { update = listener; return () => {}; },
  }, transport);
  const captures: number[] = [];
  const saves: number[] = [];
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    provider.connect();
    socket.readyState = 1;
    socket.onopen?.();
    socket.onmessage?.({ data: encodeSyncStep2(Uint8Array.of(0, 0)) });
    socket.onmessage?.({ data: encodeSyncStep1(Uint8Array.of(0)) });
    const before = transport.captureSaveBarrier()!.seq;
    socket.bufferedAmount = 2 * 1024 * 1024;
    update(Uint8Array.of(1, 2, 3), "local");
    expect(provider.pendingBytes).toBeGreaterThan(0);
    const pending = saveDocument({
      flushPendingInput: async () => {},
      hasPendingUpdates: () => provider.pendingBytes > 0,
    }, { async save(barrier) { saves.push(barrier.seq); } }, {
      inputTimeoutMs: timeout ? 5 : 500,
      hasPendingResync: () => transport.hasPendingResync(),
      captureBarrier() { const barrier = transport.captureSaveBarrier(); captures.push(barrier!.seq); return barrier; },
    });
    const outcome = pending.then(() => "saved", (error: Error) => error.message);
    await new Promise((resolve) => setTimeout(resolve, timeout ? 15 : 0));
    expect(captures).toEqual([]);
    expect(saves).toEqual([]);
    socket.bufferedAmount = 0;
    if (timeout) {
      expect(await outcome).toBe("notSynced");
      await new Promise((resolve) => setTimeout(resolve, 60));
      expect(captures).toEqual([]);
      expect(saves).toEqual([]);
    } else {
      expect(await outcome).toBe("saved");
      expect(provider.pendingBytes).toBe(0);
      expect(captures).toEqual([before + 1]);
      expect(saves).toEqual(captures);
      expect(decodeSequencedFrame(new Uint8Array(socket.sent.at(-1)!))!.seq).toBe(before + 1);
    }
  } finally {
    provider.destroy();
    transport.disconnect();
    warning.mockRestore();
  }
});

describe("collabUrl", () => {
  test("maps http and https origins to the /collab endpoint", () => {
    expect(collabUrl("http://127.0.0.1:3000")).toBe("ws://127.0.0.1:3000/collab");
    expect(collabUrl("https://office.example.com")).toBe("wss://office.example.com/collab");
  });
});

describe("createWopiTransport", () => {
  test("reports 4409 and never reconnects a replaced baseline", async () => {
    const sockets: FakeSocket[] = [];
    const transport = createWopiTransport({ handle: "h", baselineGeneration: 1, createSocket: () => {
      const socket = new FakeSocket();
      sockets.push(socket);
      return socket;
    } });
    const events: unknown[] = [];
    transport.onEvent((event) => { events.push(event); });
    try {
      transport.connect();
      const socket = sockets[0]!;
      socket.readyState = 1;
      socket.onopen?.();
      socket.readyState = 3;
      socket.onclose?.({ code: 4409, reason: "rejoin from source" });
      expect(events.at(-1)).toEqual({ type: "close", code: 4409, reason: "rejoin from source" });
      expect(transport.captureSaveBarrier()).toBeNull();
      transport.connect();
      await new Promise((resolve) => setTimeout(resolve, 300));
      expect(sockets).toHaveLength(1);
    } finally {
      transport.disconnect();
    }
  });

  test.each([1006, 4401])("close code %i keeps the existing reconnect behavior", async (code) => {
    const sockets: FakeSocket[] = [];
    const transport = createWopiTransport({ handle: "h", baselineGeneration: 1, createSocket: () => {
      const socket = new FakeSocket();
      sockets.push(socket);
      return socket;
    } });
    try {
      transport.connect();
      const first = sockets[0]!;
      first.readyState = 1;
      first.onopen?.();
      const barrier = transport.captureSaveBarrier()!;
      first.readyState = 3;
      first.onclose?.({ code });
      await new Promise((resolve) => setTimeout(resolve, 300));
      expect(sockets).toHaveLength(2);
      const second = sockets[1]!;
      second.readyState = 1;
      second.onopen?.();
      expect(transport.captureSaveBarrier()?.connectionId).not.toBe(barrier.connectionId);
      expect(transport.captureSaveBarrier()?.seq).toBe(0);
    } finally {
      transport.disconnect();
    }
  });

  test("connects to /collab with the handle as the second subprotocol", () => {
    const sockets: FakeSocket[] = [];
    const seen: Array<{ url: string; protocols: string[] }> = [];
    const transport = createWopiTransport({
      handle: "handle-abc", baselineGeneration: 1,
      origin: "https://office.example.com",
      createSocket: (url, protocols) => {
        seen.push({ url, protocols });
        const socket = new FakeSocket();
        sockets.push(socket);
        return socket;
      },
    });
    transport.connect();
    expect(seen).toEqual([{ url: "wss://office.example.com/collab", protocols: ["bo-wopi.v1", "handle-abc", "bo-wopi.generation.1"] }]);
    expect(sockets[0]!.binaryType).toBe("arraybuffer");
  });

  test("forwards binary messages and sends only while open", () => {
    const socket = new FakeSocket();
    const transport = createWopiTransport({
      handle: "h", baselineGeneration: 1,
      origin: "http://127.0.0.1:3000",
      createSocket: () => socket,
    });
    const events: string[] = [];
    transport.onEvent((event) => events.push(event.type));
    expect(transport.send(Uint8Array.of(1))).toBe(false);
    transport.connect();
    socket.readyState = 1;
    socket.onopen?.();
    expect(transport.send(Uint8Array.of(1, 2))).toBe(true);
    expect(decodeSequencedFrame(new Uint8Array(socket.sent[1]!))?.payload).toEqual(Uint8Array.of(1, 2));
    socket.onmessage?.({ data: new Uint8Array([3]).buffer });
    expect(events).toEqual(["open", "message"]);
    transport.disconnect();
    expect(socket.closedWith).toEqual([[1000, "Disconnected"]]);
  });
});

test("socket barriers number accepted frames and reset on reconnect", () => {
  const sockets: FakeSocket[] = [];
  const transport = createWopiTransport({
    handle: "h", baselineGeneration: 1,
    createSocket: () => {
      const socket = new FakeSocket();
      sockets.push(socket);
      return socket;
    },
  });
  expect(transport.captureSaveBarrier()).toBeNull();
  transport.connect();
  expect(() => transport.captureSaveBarrier()).toThrow("notSynced");
  const first = sockets[0]!;
  first.readyState = 1;
  first.onopen?.();
  const initial = transport.captureSaveBarrier()!;
  expect(initial.seq).toBe(0);
  expect(decodeSequencedFrame(new Uint8Array(first.sent[0]!))).toEqual({ ...initial, payload: new Uint8Array() });
  expect(transport.send(Uint8Array.of(0, 2, 2, 0, 0))).toBe(true);
  const captured = transport.captureSaveBarrier()!;
  expect(captured).toEqual({ connectionId: initial.connectionId, seq: 1 });
  expect(decodeSequencedFrame(new Uint8Array(first.sent[1]!))).toEqual({ ...captured, payload: Uint8Array.of(0, 2, 2, 0, 0) });
  first.bufferedAmount = 1024 * 1024 + 1;
  expect(transport.send(Uint8Array.of(0, 2, 2, 0, 0))).toBe(false);
  expect(transport.captureSaveBarrier()).toEqual(captured);
  transport.disconnect();
  expect(transport.captureSaveBarrier()).toBeNull();
  transport.connect();
  const second = sockets[1]!;
  second.readyState = 1;
  second.onopen?.();
  const reconnected = transport.captureSaveBarrier()!;
  expect(reconnected.connectionId).not.toBe(captured.connectionId);
  expect(reconnected.seq).toBe(0);
  transport.send(Uint8Array.of(0, 2, 2, 0, 0));
  expect(transport.captureSaveBarrier()).toEqual({ ...reconnected, seq: 1 });
  expect(captured.seq).toBe(1);
  transport.disconnect();
});

test("backpressure polls until drained and stops polling after close or disconnect", () => {
  const callbacks = new Map<ReturnType<typeof setTimeout>, () => void>();
  const schedule = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  const timers = spyOn(globalThis, "setTimeout").mockImplementation(((callback: () => void, delay: number) => {
    const timer = schedule(() => {}, 0);
    if (delay === 50) callbacks.set(timer, callback);
    return timer;
  }) as typeof setTimeout);
  const cancellations = spyOn(globalThis, "clearTimeout").mockImplementation(((timer: ReturnType<typeof setTimeout>) => {
    callbacks.delete(timer);
    cancel(timer);
  }) as typeof clearTimeout);
  const tick = () => {
    const pending = [...callbacks.values()];
    callbacks.clear();
    for (const callback of pending) callback();
  };
  try {
    const socket = new FakeSocket();
    const transport = createWopiTransport({ handle: "h", baselineGeneration: 1, createSocket: () => socket });
    const events: string[] = [];
    transport.onEvent((event) => events.push(event.type));
    transport.connect();
    socket.readyState = 1;
    socket.onopen?.();
    expect(callbacks.size).toBe(0);
    socket.bufferedAmount = 1024 * 1024 + 1;
    expect(transport.send(Uint8Array.of(1))).toBe(false);
    expect(transport.send(Uint8Array.of(2))).toBe(false);
    expect(callbacks.size).toBe(1);
    tick();
    expect(events).toEqual(["open"]);
    expect(callbacks.size).toBe(1);
    socket.bufferedAmount = 0;
    tick();
    expect(events).toEqual(["open", "drain"]);
    expect(callbacks.size).toBe(0);
    expect(transport.send(Uint8Array.of(3))).toBe(true);
    socket.bufferedAmount = 1024 * 1024 + 1;
    transport.send(Uint8Array.of(4));
    socket.readyState = 3;
    socket.onclose?.({});
    expect(callbacks.size).toBe(0);
    transport.disconnect();
    socket.readyState = 0;
    transport.connect();
    socket.readyState = 1;
    socket.onopen?.();
    transport.send(Uint8Array.of(5));
    transport.disconnect();
    expect(callbacks.size).toBe(0);
    tick();
    expect(events.filter((event) => event === "drain")).toHaveLength(1);
  } finally {
    timers.mockRestore();
    cancellations.mockRestore();
  }
});
