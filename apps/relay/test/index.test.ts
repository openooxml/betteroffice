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

function updateKey(seq: number): string {
  return `update:${String(seq).padStart(16, "0")}`;
}

function createRoom(seed: Iterable<[string, unknown]> = []) {
  const sender = createSocket();
  const peer = createSocket();
  const sockets: FakeSocket[] = [sender, peer];
  const rows = new Map<string, unknown>(seed);
  const putKeys: string[][] = [];
  const deletedKeys: string[][] = [];
  const pending: Promise<unknown>[] = [];
  const alarms: number[] = [];
  const deleteAll = mock(async () => {
    rows.clear();
  });
  let initialization = Promise.resolve();
  const state = {
    storage: {
      get: async (key: string) => rows.get(key),
      list: async ({ prefix }: { prefix: string }) =>
        new Map(
          [...rows]
            .filter(([key]) => key.startsWith(prefix))
            .sort(([first], [second]) => (first < second ? -1 : 1)),
        ),
      put: async (
        keyOrBatch: string | Record<string, Uint8Array>,
        value?: Uint8Array,
      ) => {
        if (typeof keyOrBatch === "string") {
          rows.set(keyOrBatch, value);
          putKeys.push([keyOrBatch]);
          return;
        }
        for (const [key, bytes] of Object.entries(keyOrBatch)) {
          rows.set(key, bytes);
        }
        putKeys.push(Object.keys(keyOrBatch));
      },
      delete: async (keys: string | readonly string[]) => {
        const batch = typeof keys === "string" ? [keys] : [...keys];
        for (const key of batch) rows.delete(key);
        deletedKeys.push(batch);
        return batch.length;
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
    deletedKeys,
    initialization,
    peer,
    pending,
    putKeys,
    room,
    rows,
    sender,
    sockets,
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
    expect(harness.putKeys).toEqual([[updateKey(0)]]);
    expect(harness.rows.get(updateKey(0))).toEqual(document);
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
    expect(harness.rows.size).toBe(0);
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
    expect(harness.rows.size).toBe(0);
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
    expect(harness.rows.size).toBe(0);
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
    expect(harness.rows.size).toBe(0);
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
    expect(harness.rows.size).toBe(1);

    harness.sockets.length = 0;
    await harness.room.alarm();

    expect(harness.deleteAll).toHaveBeenCalledTimes(1);
    expect(harness.rows.size).toBe(0);

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
    expect(harness.rows.size).toBe(1);
    expect(harness.alarms).toEqual([START_MS + DAY_MS, START_MS + 2 * DAY_MS]);
  });
});

describe("CollaborationRoom persistence", () => {
  function documentAt(payload: number): Uint8Array {
    return Uint8Array.of(0, 2, 1, payload);
  }

  function send(
    harness: ReturnType<typeof createRoom>,
    document: Uint8Array,
  ): void {
    harness.room.webSocketMessage(
      harness.sender as never,
      document.buffer as ArrayBuffer,
    );
  }

  function replayed(socket: FakeSocket, count: number): unknown[] {
    return socket.send.mock.calls.slice(0, count).map(([value]) => value);
  }

  test("writes one key per retained frame and rewrites nothing", async () => {
    const harness = createRoom();
    await harness.initialization;

    for (const payload of [1, 2, 3]) send(harness, documentAt(payload));
    await Promise.all(harness.pending);

    expect(harness.putKeys).toEqual([
      [updateKey(0)],
      [updateKey(1)],
      [updateKey(2)],
    ]);
    expect(harness.deletedKeys).toEqual([]);
    expect(harness.rows.size).toBe(3);
  });

  test("evicts only the oldest key once the entry cap is reached", async () => {
    const harness = createRoom();
    await harness.initialization;

    for (let index = 0; index <= 512; index += 1) {
      send(harness, documentAt(index & 0xff));
    }
    await Promise.all(harness.pending);

    expect(harness.rows.size).toBe(512);
    expect(harness.deletedKeys).toEqual([[updateKey(0)]]);
    expect(harness.rows.has(updateKey(512))).toBe(true);
  });

  test("rehydrates in sequence order and appends above the highest seq", async () => {
    const first = documentAt(4);
    const second = documentAt(5);
    const harness = createRoom([
      [updateKey(9), second],
      [updateKey(2), first],
    ]);
    await harness.initialization;

    expect(harness.putKeys).toEqual([]);
    expect(harness.deletedKeys).toEqual([]);
    expect(replayed(await join(harness), 2)).toEqual([first, second]);

    send(harness, documentAt(6));
    await Promise.all(harness.pending);
    expect(harness.putKeys).toEqual([[updateKey(10)]]);
  });

  test("repairs a non-canonical entry and drops unreadable ones", async () => {
    const document = documentAt(7);
    const survivor = documentAt(9);
    const harness = createRoom([
      [updateKey(0), frame(document, Uint8Array.of(1, 1, 8))],
      [updateKey(1), Uint8Array.of(0x80)],
      ["update:nope", document],
      [updateKey(2), survivor],
    ]);
    await harness.initialization;

    expect(harness.deletedKeys).toEqual([["update:nope"], [updateKey(1)]]);
    expect(harness.putKeys).toEqual([[updateKey(0)]]);
    expect(harness.rows.get(updateKey(0))).toEqual(document);
    expect(replayed(await join(harness), 2)).toEqual([document, survivor]);
  });

  test("discards the legacy single-blob log on rehydrate", async () => {
    const harness = createRoom([["updates", [documentAt(10)]]]);
    await harness.initialization;

    expect(harness.deletedKeys).toEqual([["updates"]]);
    expect(harness.rows.size).toBe(0);
    expect((await join(harness)).send).toHaveBeenCalledTimes(1);
  });
});
