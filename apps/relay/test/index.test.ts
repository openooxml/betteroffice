import { describe, expect, mock, test } from "bun:test";

mock.module("cloudflare:workers", () => ({
  DurableObject: class {
    protected readonly ctx: unknown;
    protected readonly env: unknown;

    constructor(ctx: unknown, env: unknown) {
      this.ctx = ctx;
      this.env = env;
    }
  },
}));

const { CollaborationRoom } = await import("../src/index");

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

function createRoom() {
  const sender = {
    send: mock(() => {}),
    close: mock(() => {}),
  };
  const peer = {
    send: mock(() => {}),
    close: mock(() => {}),
  };
  const storageWrites: Array<{
    key: string;
    value: Uint8Array[];
  }> = [];
  const pending: Promise<unknown>[] = [];
  let initialization = Promise.resolve();
  const state = {
    storage: {
      get: async () => [],
      put: async (key: string, value: Uint8Array[]) => {
        storageWrites.push({ key, value });
      },
    },
    blockConcurrencyWhile: (initialize: () => Promise<void>) => {
      initialization = initialize();
    },
    getWebSockets: () => [sender, peer],
    waitUntil: (promise: Promise<unknown>) => {
      pending.push(promise);
    },
  };
  const room = new CollaborationRoom(state as never, {} as never);
  return {
    initialization,
    peer,
    pending,
    room,
    sender,
    storageWrites,
  };
}

describe("CollaborationRoom.webSocketMessage", () => {
  test("broadcasts original mixed bytes and persists only sync messages", async () => {
    const document = Uint8Array.of(0, 2, 2, 10, 11);
    const awareness = Uint8Array.of(1, 1, 12);
    const mixed = frame(document, awareness);
    const harness = createRoom();
    await harness.initialization;

    harness.room.webSocketMessage(
      harness.sender as never,
      mixed.buffer as ArrayBuffer,
    );
    await Promise.all(harness.pending);

    expect(harness.peer.send).toHaveBeenCalledTimes(1);
    expect(harness.peer.send.mock.calls[0][0]).toEqual(mixed);
    expect(harness.sender.send).not.toHaveBeenCalled();
    expect(harness.storageWrites).toEqual([
      { key: "updates", value: [document] },
    ]);
  });

  test("broadcasts malformed bytes without throwing or persisting", async () => {
    const document = Uint8Array.of(0, 2, 1, 13);
    const malformed = frame(document, Uint8Array.of(1, 2, 14));
    const harness = createRoom();
    await harness.initialization;

    expect(() =>
      harness.room.webSocketMessage(
        harness.sender as never,
        malformed.buffer as ArrayBuffer,
      ),
    ).not.toThrow();

    expect(harness.peer.send).toHaveBeenCalledTimes(1);
    expect(harness.peer.send.mock.calls[0][0]).toEqual(malformed);
    expect(harness.sender.send).not.toHaveBeenCalled();
    expect(harness.pending).toEqual([]);
    expect(harness.storageWrites).toEqual([]);
  });
});
