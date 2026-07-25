import { afterEach, describe, expect, mock, setSystemTime, test } from "bun:test";

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

function createSocket() {
  return { send: mock((_: unknown) => {}), close: mock(() => {}) };
}

type FakeSocket = ReturnType<typeof createSocket>;

(globalThis as { WebSocketPair?: unknown }).WebSocketPair = function () {
  return { 0: createSocket(), 1: createSocket() };
};

const UPGRADE = new Request("https://relay.test/room/a", {
  headers: { Upgrade: "websocket" },
});

function createRoom() {
  const sender = createSocket();
  const peer = createSocket();
  const sockets: FakeSocket[] = [sender, peer];
  const storageWrites: Array<{
    key: string;
    value: Uint8Array[];
  }> = [];
  const pending: Promise<unknown>[] = [];
  const alarms: number[] = [];
  const deleteAll = mock(async () => {
    storageWrites.length = 0;
  });
  let initialization = Promise.resolve();
  const state = {
    storage: {
      get: async () => [],
      put: async (key: string, value: Uint8Array[]) => {
        storageWrites.push({ key, value });
      },
      getAlarm: async () => null,
      setAlarm: async (time: number) => {
        alarms.push(time);
      },
      deleteAll,
    },
    blockConcurrencyWhile: (initialize: () => Promise<void>) => {
      initialization = initialize();
    },
    acceptWebSocket: (socket: FakeSocket) => {
      sockets.push(socket);
    },
    getWebSockets: () => sockets,
    waitUntil: (promise: Promise<unknown>) => {
      pending.push(promise);
    },
  };
  const room = new CollaborationRoom(state as never, {} as never);
  return {
    alarms,
    deleteAll,
    initialization,
    peer,
    pending,
    room,
    sender,
    sockets,
    storageWrites,
  };
}

async function join(harness: ReturnType<typeof createRoom>) {
  await harness.room.fetch(UPGRADE);
  return harness.sockets[harness.sockets.length - 1];
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

  test("closes only the sender on a malformed frame and broadcasts nothing", async () => {
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

    expect(harness.sender.close).toHaveBeenCalledTimes(1);
    expect(harness.sender.close.mock.calls[0][0]).toBe(1002);
    expect(harness.peer.send).not.toHaveBeenCalled();
    expect(harness.peer.close).not.toHaveBeenCalled();
    expect(harness.pending).toEqual([]);
    expect(harness.storageWrites).toEqual([]);
  });

  test("closes only the sender on a client-origin auth denial", async () => {
    const auth = Uint8Array.of(2, 0, 2, 104, 105);
    const harness = createRoom();
    await harness.initialization;

    harness.room.webSocketMessage(
      harness.sender as never,
      auth.buffer as ArrayBuffer,
    );

    expect(harness.sender.close).toHaveBeenCalledTimes(1);
    expect(harness.sender.close.mock.calls[0][0]).toBe(1008);
    expect(harness.peer.send).not.toHaveBeenCalled();
    expect(harness.peer.close).not.toHaveBeenCalled();
    expect(harness.storageWrites).toEqual([]);
  });

  test("broadcasts awareness-only frames without retaining them", async () => {
    const awareness = Uint8Array.of(1, 1, 12);
    const harness = createRoom();
    await harness.initialization;

    harness.room.webSocketMessage(
      harness.sender as never,
      awareness.buffer as ArrayBuffer,
    );

    expect(harness.sender.close).not.toHaveBeenCalled();
    expect(harness.peer.send).toHaveBeenCalledTimes(1);
    expect(harness.peer.send.mock.calls[0][0]).toEqual(awareness);
    expect(harness.storageWrites).toEqual([]);
  });

  test("broadcasts sync-step-1 so live peers answer it, without retaining it", async () => {
    const query = Uint8Array.of(0, 0, 1, 15);
    const harness = createRoom();
    await harness.initialization;

    harness.room.webSocketMessage(
      harness.sender as never,
      query.buffer as ArrayBuffer,
    );

    expect(harness.sender.close).not.toHaveBeenCalled();
    expect(harness.peer.send).toHaveBeenCalledTimes(1);
    expect(harness.peer.send.mock.calls[0][0]).toEqual(query);
    expect(harness.storageWrites).toEqual([]);
  });
});

describe("CollaborationRoom expiry", () => {
  const HOUR_MS = 60 * 60 * 1000;
  const DAY_MS = 24 * HOUR_MS;
  const START_MS = Date.UTC(2026, 0, 1);
  const document = Uint8Array.of(0, 2, 2, 10, 11);

  afterEach(() => {
    setSystemTime();
  });

  function send(harness: ReturnType<typeof createRoom>): void {
    harness.room.webSocketMessage(
      harness.sender as never,
      document.buffer as ArrayBuffer,
    );
  }

  test("throttles deadline rewrites but extends on later activity", async () => {
    setSystemTime(new Date(START_MS));
    const harness = createRoom();
    await harness.initialization;

    send(harness);
    send(harness);
    await Promise.all(harness.pending);
    expect(harness.alarms).toEqual([START_MS + DAY_MS]);

    setSystemTime(new Date(START_MS + 2 * HOUR_MS));
    send(harness);
    await Promise.all(harness.pending);
    expect(harness.alarms).toEqual([
      START_MS + DAY_MS,
      START_MS + 2 * HOUR_MS + DAY_MS,
    ]);
  });

  test("replays retained frames to a joining socket", async () => {
    const harness = createRoom();
    await harness.initialization;
    send(harness);
    await Promise.all(harness.pending);

    const joiner = await join(harness);
    expect(joiner.send.mock.calls[0][0]).toEqual(document);
  });

  test("wipes an idle room when the alarm fires", async () => {
    const harness = createRoom();
    await harness.initialization;
    send(harness);
    await Promise.all(harness.pending);
    expect(harness.storageWrites).toHaveLength(1);

    harness.sockets.length = 0;
    await harness.room.alarm();

    expect(harness.deleteAll).toHaveBeenCalledTimes(1);
    expect(harness.storageWrites).toEqual([]);

    const joiner = await join(harness);
    expect(joiner.send).toHaveBeenCalledTimes(1);
    expect(joiner.send.mock.calls[0][0]).toBe(
      JSON.stringify({ type: "peers", count: 1 }),
    );
  });

  test("reschedules instead of wiping while a socket is connected", async () => {
    setSystemTime(new Date(START_MS));
    const harness = createRoom();
    await harness.initialization;
    send(harness);
    await Promise.all(harness.pending);

    setSystemTime(new Date(START_MS + DAY_MS));
    await harness.room.alarm();

    expect(harness.deleteAll).not.toHaveBeenCalled();
    expect(harness.storageWrites).toHaveLength(1);
    expect(harness.alarms).toEqual([START_MS + DAY_MS, START_MS + 2 * DAY_MS]);
  });
});
