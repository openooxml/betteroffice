import { expect, spyOn, test } from "bun:test";
import { applyUpdate, Doc, encodeStateAsUpdate, encodeStateVector } from "yjs";
import type { ServerReplica } from "../replica/replica";
import type { FileSession } from "../session/fileSession";
import type { SessionRegistry } from "../session/registry";
import { UserSession } from "../session/userSession";
import { requestOrderedSave } from "../sessionApi";
import { encodeUpdateFrame } from "./frames";
import { CollabRelay, type RelaySocket } from "./relay";
import { decodeSequenceAck, encodeSequencedFrame } from "./sequence";

function setup() {
  const server = new Doc();
  const client = new Doc();
  client.getText("body").insert(0, "before");
  applyUpdate(server, encodeStateAsUpdate(client));
  const user = new UserSession({
    hostId: "host", wopiSrc: "https://host.example/wopi/files/file", accessToken: "token", accessTokenTtl: 0,
    userId: "user", friendlyName: "User", canWrite: true, joinedAt: 0,
  });
  const saved: string[] = [];
  const closeCodes: Array<number | undefined> = [];
  const sent: Uint8Array[] = [];
  const socket: RelaySocket = { send(frame) { sent.push(frame); }, close(code) { closeCodes.push(code); } };
  const replica = { version: 0, encodeStateVector: () => encodeStateVector(server) };
  const session = {
    fileKey: "file", viewOnly: false, users: [user], baselineGeneration: 1,
    replica,
    on() { return () => {}; },
    connect(userId: string) {
      if (userId !== user.id) throw new Error("unknown user");
      const generation = ++user.connectionGeneration;
      this.heartbeat(userId, generation);
      return generation;
    },
    heartbeat(userId: string, generation = user.connectionGeneration) {
      if (userId === user.id && generation === user.connectionGeneration) user.connected = true;
    },
    async disconnect(userId: string, generation = user.connectionGeneration) {
      if (userId === user.id && generation === user.connectionGeneration) user.connected = false;
    },
    getUser: (userId: string) => userId === user.id ? user : null,
    isUserTokenValid: (userId: string) => userId === user.id && user.isTokenValidAt(0),
    applyUpdate(_userId: string, update: Uint8Array) {
      expect(relay.appliedSequence(session as unknown as FileSession, user.id, connectionId)).toBe(seq - 1);
      applyUpdate(server, update);
      replica.version += 1;
      return { accepted: true as const, followUps: [] };
    },
    async requestSave() {
      saved.push(server.getText("body").toString());
      return { saved: true, itemVersion: "2", by: user.id };
    },
  } satisfies Pick<FileSession,
    "fileKey" | "viewOnly" | "users" | "baselineGeneration" | "on" | "connect" | "heartbeat" | "disconnect" |
    "getUser" | "isUserTokenValid" | "applyUpdate" | "requestSave"
  > & { replica: Pick<ServerReplica, "version" | "encodeStateVector"> };
  const file = session as unknown as FileSession;
  const registry = { list: () => [file] } as unknown as SessionRegistry;
  const relay = new CollabRelay(registry);
  const connectionId = crypto.randomUUID();
  let seq = 0;
  relay.attach(socket, { session: file, user }, file.baselineGeneration);
  relay.handleMessage(socket, encodeSequencedFrame({ connectionId, seq }));
  return {
    server, client, file, user, socket, relay, saved, closeCodes, connectionId, sent,
    send(update: Uint8Array) {
      seq += 1;
      relay.handleMessage(socket, encodeSequencedFrame({ connectionId, seq }, encodeUpdateFrame(update)));
    },
    destroy() { server.destroy(); client.destroy(); },
  };
}

test.each(["deletion", "insertion"] as const)("ordered save waits for a delayed %s frame to be applied", async (kind) => {
  const h = setup();
  try {
    const vector = encodeStateVector(h.client);
    if (kind === "deletion") h.client.getText("body").delete(0, 6);
    else h.client.getText("body").insert(6, " after");
    if (kind === "deletion") expect(encodeStateVector(h.client)).toEqual(vector);
    const barrier = { connectionId: h.connectionId, seq: 1 };
    let completed = false;
    const saving = requestOrderedSave(h.file, h.user.id, barrier, { relay: h.relay, timeoutMs: 500 })
      .then(() => { completed = true; });
    await new Promise((resolve) => setTimeout(resolve, 60));
    expect(completed).toBe(false);
    expect(h.saved).toEqual([]);
    h.send(encodeStateAsUpdate(h.client, vector));
    await saving;
    expect(h.saved).toEqual([kind === "deletion" ? "" : "before after"]);
  } finally {
    h.destroy();
  }
});

test("socket saves time out without falling back to an already covered vector", async () => {
  const h = setup();
  try {
    await expect(requestOrderedSave(h.file, h.user.id, { connectionId: h.connectionId, seq: 1 }, {
      relay: h.relay, timeoutMs: 1,
    })).rejects.toMatchObject({ kind: "notSynced" });
    await expect(requestOrderedSave(h.file, h.user.id, { connectionId: crypto.randomUUID(), seq: 0 }, {
      relay: h.relay, timeoutMs: 1,
    })).rejects.toMatchObject({ kind: "notSynced" });
    expect(h.saved).toEqual([]);
  } finally {
    h.destroy();
  }
});

test("replacement sockets reset the applied sequence and cannot acknowledge an old connection", async () => {
  const h = setup();
  try {
    h.client.getText("body").insert(6, " after");
    h.send(encodeStateAsUpdate(h.client));
    expect(h.relay.appliedSequence(h.file, h.user.id, h.connectionId)).toBe(1);
    const generation = h.user.connectionGeneration;
    const socket: RelaySocket = { send() {}, close() {} };
    const connectionId = crypto.randomUUID();
    h.relay.attach(socket, { session: h.file, user: h.user }, h.file.baselineGeneration);
    expect(h.user.connectionGeneration).toBe(generation + 1);
    h.relay.detach(h.socket);
    expect(h.user.connected).toBe(true);
    h.relay.handleMessage(socket, encodeSequencedFrame({ connectionId, seq: 0 }));
    expect(h.relay.appliedSequence(h.file, h.user.id, h.connectionId)).toBeNull();
    expect(h.relay.appliedSequence(h.file, h.user.id, connectionId)).toBe(0);
    await expect(requestOrderedSave(h.file, h.user.id, { connectionId: h.connectionId, seq: 1 }, {
      relay: h.relay, timeoutMs: 1,
    })).rejects.toMatchObject({ kind: "notSynced" });
    await requestOrderedSave(h.file, h.user.id, { connectionId, seq: 0 }, { relay: h.relay, timeoutMs: 0 });
    expect(h.saved).toEqual(["before after"]);
  } finally {
    h.destroy();
  }
});

test("a rejected sequenced update forces resync and cannot be acknowledged by later frames", async () => {
  const h = setup();
  const applying = spyOn(h.file, "applyUpdate").mockReturnValue({ accepted: false, reason: "causalGap" });
  try {
    h.client.getText("body").insert(6, " after");
    h.send(encodeStateAsUpdate(h.client));
    expect(h.closeCodes).toEqual([4400]);
    expect(h.relay.appliedSequence(h.file, h.user.id, h.connectionId)).toBeNull();
    h.send(encodeStateAsUpdate(h.client));
    expect(applying).toHaveBeenCalledTimes(1);
    await expect(requestOrderedSave(h.file, h.user.id, { connectionId: h.connectionId, seq: 1 }, {
      relay: h.relay, timeoutMs: 1,
    })).rejects.toMatchObject({ kind: "notSynced" });
    expect(h.saved).toEqual([]);
  } finally {
    applying.mockRestore();
    h.destroy();
  }
});

test("the relay acknowledges accepted document frames and withholds acknowledgement on rejection", () => {
  const h = setup();
  try {
    h.client.getText("body").insert(6, " accepted");
    h.send(encodeStateAsUpdate(h.client));
    expect(h.sent.map(decodeSequenceAck).filter(Boolean)).toContainEqual({ connectionId: h.connectionId, seq: 1 });
    const rejecting = spyOn(h.file, "applyUpdate").mockReturnValue({ accepted: false, reason: "causalGap" });
    try {
      h.send(encodeStateAsUpdate(h.client));
      expect(h.sent.map(decodeSequenceAck).filter(Boolean)).not.toContainEqual({ connectionId: h.connectionId, seq: 2 });
    } finally { rejecting.mockRestore(); }
  } finally { h.destroy(); }
});
