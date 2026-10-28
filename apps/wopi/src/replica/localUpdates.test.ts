import { describe, expect, test } from "bun:test";
import * as Y from "yjs";

import { Replica } from "./replica";

function harness(events = true) {
  const doc = new Y.Doc({ gc: false });
  doc.getText("t").insert(0, "abcdef");
  let repair = false;
  let unsubscribed = false;
  const engine = {
    clientId: doc.clientID,
    encodeStateVector: () => Y.encodeStateVector(doc),
    encodeStateAsUpdate: (remote?: Uint8Array) => Y.encodeStateAsUpdate(doc, remote),
    applyUpdate(update: Uint8Array) {
      Y.applyUpdate(doc, update, "remote");
      if (repair) {
        doc.transact(() => doc.getText("t").delete(0, 1), "local");
        repair = false;
      }
    },
    onLocalUpdate: events
      ? (listener: (update: Uint8Array) => void) => {
          const receive = (update: Uint8Array, origin: unknown) => {
            if (origin === "local") listener(update);
          };
          doc.on("update", receive);
          return () => {
            unsubscribed = true;
            doc.off("update", receive);
          };
        }
      : undefined,
    engineVersion: () => doc.getText("t").toString(),
    dispose: () => doc.destroy(),
  };
  return {
    doc,
    engine,
    repairOnApply: () => { repair = true; },
    unsubscribed: () => unsubscribed,
    open: (save: () => Promise<Uint8Array> = async () => new Uint8Array()) =>
      new Replica("docx", engine, save, () => ({
        format: "docx",
        initialUpdate: Y.encodeStateAsUpdate(doc),
      })),
  };
}

function expectDeleteOnly(update: Uint8Array): void {
  const decoded = Y.decodeUpdate(update);
  expect(decoded.structs).toEqual([]);
  expect(decoded.ds.clients.size).toBeGreaterThan(0);
}

describe("replica local update publication", () => {
  for (const events of [true, false]) {
    test(`delete-only apply repair is emitted and returned with events=${events}`, () => {
      const state = harness(events);
      const server = state.open();
      const peer = new Y.Doc({ gc: false });
      Y.applyUpdate(peer, server.encodeStateAsUpdate());
      const vector = server.encodeStateVector();
      peer.getText("t").delete(5, 1);
      const remote = Y.encodeStateAsUpdate(peer, vector);
      const emitted: Uint8Array[] = [];
      server.onLocalUpdate((update) => emitted.push(update));
      state.repairOnApply();
      try {
        const result = server.applyUpdate(remote);
        expect(result.accepted).toBe(true);
        expect(server.encodeStateVector()).toEqual(vector);
        expect(emitted).toHaveLength(1);
        expectDeleteOnly(emitted[0]!);
        expect(Y.decodeUpdate(emitted[0]!).ds.clients.get(state.doc.clientID))
          .toEqual([{ clock: 0, len: 1 }]);
        expect(result).toEqual({ accepted: true, followUps: emitted });
        expect(peer.getText("t").toString()).toBe("abcde");
        Y.applyUpdate(peer, emitted[0]!);
        expect(peer.getText("t").toString()).toBe(state.doc.getText("t").toString());
        expect(server.applyUpdate(remote)).toEqual({ accepted: true });
        expect(emitted).toHaveLength(1);
      } finally {
        server.destroy();
        peer.destroy();
      }
    });

    test(`delete-only save repair is emitted and returned once with events=${events}`, async () => {
      const state = harness(events);
      let repair = true;
      const server = state.open(async () => {
        if (repair) {
          state.doc.transact(() => state.doc.getText("t").delete(0, 1), "local");
          state.doc.transact(() => state.doc.getText("t").delete(0, 1), "local");
          repair = false;
        }
        return new TextEncoder().encode(state.doc.getText("t").toString());
      });
      const vector = server.encodeStateVector();
      const emitted: Uint8Array[] = [];
      server.onLocalUpdate((update) => emitted.push(update));
      try {
        const saved = await server.save();
        expect(saved.update).not.toBeNull();
        expectDeleteOnly(saved.update!);
        expect(server.encodeStateVector()).toEqual(vector);
        expect(emitted).toEqual([saved.update!]);
        expect(new TextDecoder().decode(saved.bytes)).toBe("cdef");
        expect((await server.save()).update).toBeNull();
        expect(emitted).toHaveLength(1);
      } finally {
        server.destroy();
      }
      await Promise.resolve();
      expect(state.unsubscribed()).toBe(events);
    });
  }

  test("fallback publishes only local structs and new local deletions", () => {
    const state = harness(false);
    const apply = state.engine.applyUpdate;
    state.engine.applyUpdate = (update) => {
      apply(update);
      state.doc.transact(() => {
        state.doc.getText("t").insert(0, "local-");
        state.doc.getText("t").delete(6, 1);
      }, "local");
    };
    const server = state.open();
    const peer = new Y.Doc({ gc: false });
    Y.applyUpdate(peer, server.encodeStateAsUpdate());
    const vector = server.encodeStateVector();
    peer.getText("t").insert(6, "-remote");
    peer.getText("t").delete(5, 1);
    const emitted: Uint8Array[] = [];
    server.onLocalUpdate((update) => emitted.push(update));
    try {
      const result = server.applyUpdate(Y.encodeStateAsUpdate(peer, vector));
      expect(result).toEqual({ accepted: true, followUps: emitted });
      expect(emitted).toHaveLength(1);
      expect([...Y.parseUpdateMeta(emitted[0]!).to.keys()]).toEqual([server.clientId]);
      expect(Y.decodeUpdate(emitted[0]!).ds.clients.get(server.clientId))
        .toEqual([{ clock: 0, len: 1 }]);
      Y.applyUpdate(peer, emitted[0]!);
      expect(peer.getText("t").toString()).toBe(state.doc.getText("t").toString());
      expect(peer.getText("t").toString()).toBe("local-bcde-remote");
    } finally {
      server.destroy();
      peer.destroy();
    }
  });

  test("an apply repair during a suspended save is published only by apply", async () => {
    const state = harness();
    let release!: () => void;
    const suspended = new Promise<void>((resolve) => { release = resolve; });
    const server = state.open(async () => {
      await suspended;
      state.doc.transact(() => state.doc.getText("t").delete(0, 1), "local");
      return new Uint8Array();
    });
    const emitted: Uint8Array[] = [];
    server.onLocalUpdate((update) => emitted.push(update));
    const pending = server.save();
    await Promise.resolve();
    state.repairOnApply();
    const applied = server.applyUpdate(Uint8Array.of(0, 0));
    release();
    try {
      const saved = await pending;
      expect(emitted).toHaveLength(2);
      expect(applied).toEqual({ accepted: true, followUps: [emitted[0]!] });
      expect(saved.update).toEqual(emitted[1]!);
      expect(Y.decodeUpdate(saved.update!).ds.clients.get(state.doc.clientID))
        .toEqual([{ clock: 1, len: 1 }]);
    } finally {
      server.destroy();
    }
  });

  test("destroy suppresses publication from an in-flight save", async () => {
    const state = harness();
    let release!: () => void;
    const suspended = new Promise<void>((resolve) => { release = resolve; });
    const server = state.open(async () => {
      await suspended;
      state.doc.transact(() => state.doc.getText("t").delete(0, 1), "local");
      return new Uint8Array();
    });
    const emitted: Uint8Array[] = [];
    server.onLocalUpdate((update) => emitted.push(update));
    const pending = server.save();
    await Promise.resolve();
    server.destroy();
    release();
    const saved = await pending;
    expect(saved.update).not.toBeNull();
    expect(emitted).toEqual([]);
    await Promise.resolve();
    expect(state.unsubscribed()).toBe(true);
  });
});
