import { beforeAll, describe, expect, test } from "bun:test";
import { initEngines } from "../replica/engines";
import { openFreshClient, openTestClient, sourceBytes, type TestClient } from "../../test/replica/clients";
import { addToken, makeHarness, SRC, ttl, type Harness } from "../../test/session/harness";
import type { FileSession } from "./fileSession";
import { WopiHostError } from "../host/errors";

beforeAll(async () => {
  await initEngines();
});

let cid = 50000;
function nextId(): number {
  cid += 11;
  return cid;
}

async function clientFor(session: FileSession, joinPayload: Parameters<typeof openTestClient>[0], source?: Uint8Array): Promise<TestClient> {
  return openTestClient(joinPayload, source ?? sourceBytes("xlsx"), nextId());
}

function edit(session: FileSession, userId: string, client: TestClient, marker: string, slot = 0): void {
  client.edit(marker, slot);
  const replica = session.replica;
  if (!replica) throw new Error("no replica");
  const result = session.applyUpdate(userId, client.encodeStateAsUpdate(replica.encodeStateVector()));
  expect(result.accepted).toBe(true);
}

function sync(session: FileSession, client: TestClient): void {
  const replica = session.replica;
  if (!replica) throw new Error("no replica");
  client.applyUpdate(replica.encodeStateAsUpdate(client.encodeStateVector()));
}

async function savedContains(h: Harness, marker: string): Promise<boolean> {
  const bytes = h.host.files.get("file1")!.bytes;
  const fresh = await openFreshClient("xlsx", bytes, nextId());
  try {
    return fresh.contains(marker);
  } finally {
    fresh.destroy();
  }
}

async function joinEditor(h: Harness, token: string, mode: "edit" | "view" = "edit") {
  return h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: token, accessTokenTtl: 0, mode });
}

describe("file sessions with two and three editors", () => {
  test("two editors converge, lock once, one autosave per debounce", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const b = await joinEditor(h, "tok-b");
    const session = h.registry.get(a.sessionKey ?? "") ?? [...h.registry.list()][0]!;
    const clients: TestClient[] = [];
    try {
      expect(a.file.name).toBe("book.xlsx");
      expect(a.user.canWrite).toBe(true);
      expect(b.user.canWrite).toBe(true);
      const ca = await clientFor(session, a.joinPayload);
      const cb = await clientFor(session, b.joinPayload);
      clients.push(ca, cb);
      edit(session, a.user.id, ca, "alpha-one ", 0);
      edit(session, b.user.id, cb, "bravo-two ", 1);
      sync(session, ca);
      sync(session, cb);
      expect(ca.contains("bravo-two")).toBe(true);
      expect(cb.contains("alpha-one")).toBe(true);
      expect(h.calls.filter((c) => c.override === "LOCK" && !c.headers.get("X-WOPI-OldLock")).length).toBe(1);
      expect(h.calls.filter((c) => c.pathname.endsWith("/contents") && c.override === null).length).toBe(1);
      await h.clock.advance(30000);
      expect(h.putCount()).toBe(1);
      expect(await savedContains(h, "alpha-one")).toBe(true);
      expect(await savedContains(h, "bravo-two")).toBe(true);
      const put = h.calls.find((c) => c.override === "PUT")!;
      expect(put.headers.get("X-COOL-WOPI-IsAutosave")).toBe("true");
      expect(put.headers.get("X-COOL-WOPI-IsModifiedByUser")).toBe("true");
      expect(put.headers.get("X-COOL-WOPI-Timestamp")).toBeTruthy();
    } finally {
      for (const c of clients) c.destroy();
      await session.leave(a.user.id).catch(() => {});
      await session.leave(b.user.id).catch(() => {});
    }
  });

  test("three editors converge and save every edit", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    addToken(h, "tok-c", "user-c");
    const joins = [await joinEditor(h, "tok-a"), await joinEditor(h, "tok-b"), await joinEditor(h, "tok-c")];
    const session = [...h.registry.list()][0]!;
    const clients: TestClient[] = [];
    try {
      for (const j of joins) clients.push(await clientFor(session, j.joinPayload));
      edit(session, joins[0]!.user.id, clients[0]!, "m-one ", 0);
      edit(session, joins[1]!.user.id, clients[1]!, "m-two ", 1);
      edit(session, joins[2]!.user.id, clients[2]!, "m-three ", 2);
      for (const c of clients) sync(session, c);
      for (const c of clients) {
        expect(c.contains("m-one")).toBe(true);
        expect(c.contains("m-two")).toBe(true);
        expect(c.contains("m-three")).toBe(true);
      }
      await h.clock.advance(30000);
      expect(h.putCount()).toBe(1);
      expect(await savedContains(h, "m-three")).toBe(true);
    } finally {
      for (const c of clients) c.destroy();
      for (const j of joins) await session.leave(j.user.id).catch(() => {});
    }
  });

  test("refresh lock runs every ten minutes with the leader token", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    try {
      for (let i = 0; i < 20; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
      }
      const refreshes = h.callsFor("REFRESH_LOCK");
      expect(refreshes.length).toBe(1);
      expect(refreshes[0]?.token).toBe("tok-a");
      for (let i = 0; i < 20; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
      }
      expect(h.callsFor("REFRESH_LOCK").length).toBe(2);
    } finally {
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("continuous updates still save at most every five minutes cap forces a save", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "tick-0 ", 0);
      for (let i = 1; i <= 11; i += 1) {
        await h.clock.advance(29000);
        edit(session, a.user.id, client, `tick-${i} `, 0);
      }
      expect(h.putCount()).toBe(1);
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });
});

describe("leader handoff and explicit save", () => {
  test("leader leaves, near expiry leader hands off, refresh and put use new token", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a", { expiresAt: ttl(h.clock, 60 * 60000) });
    addToken(h, "tok-b", "user-b", { expiresAt: ttl(h.clock, 8 * 60000) });
    addToken(h, "tok-c", "user-c", { expiresAt: ttl(h.clock, 60 * 60000) });
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: ttl(h.clock, 60 * 60000), mode: "edit" });
    const b = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: ttl(h.clock, 8 * 60000), mode: "edit" });
    const c = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-c", accessTokenTtl: ttl(h.clock, 60 * 60000), mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const events: string[] = [];
    session.on("leaderChanged", (e) => events.push(String((e as { leaderId?: unknown }).leaderId ?? "null")));
    try {
      expect(session.leader?.id).toBe(a.user.id);
      await session.leave(a.user.id);
      expect(session.leader?.id).toBe(b.user.id);
      const refreshAfterLeave = h.callsFor("REFRESH_LOCK");
      expect(refreshAfterLeave.length).toBeGreaterThanOrEqual(1);
      expect(refreshAfterLeave[refreshAfterLeave.length - 1]?.token).toBe("tok-b");
      for (let i = 0; i < 12; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(b.user.id);
        session.heartbeat(c.user.id);
      }
      await h.clock.advance(1000);
      expect(session.leader?.id).toBe(c.user.id);
      const refreshesAfterHandoff = h.callsFor("REFRESH_LOCK");
      expect(refreshesAfterHandoff.length).toBeGreaterThanOrEqual(1);
      expect(refreshesAfterHandoff[refreshesAfterHandoff.length - 1]?.token).toBe("tok-c");
      h.calls.length = 0;
      const client = await clientFor(session, c.joinPayload);
      try {
        edit(session, c.user.id, client, "handoff-edit ", 0);
        await session.requestSave(c.user.id);
        const puts = h.callsFor("PUT");
        expect(puts.length).toBe(1);
        expect(puts[0]?.token).toBe("tok-c");
      } finally {
        client.destroy();
      }
      expect(events.length).toBeGreaterThanOrEqual(1);
    } finally {
      await session.leave(b.user.id).catch(() => {});
      await session.leave(c.user.id).catch(() => {});
    }
  });

  test("explicit save credits the requesting user token", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const b = await joinEditor(h, "tok-b");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, b.joinPayload);
    const savedEvents: Array<{ by?: unknown; itemVersion?: unknown }> = [];
    session.on("saved", (e) => savedEvents.push(e as { by?: unknown }));
    try {
      edit(session, b.user.id, client, "credit-me ", 0);
      await session.requestSave(b.user.id);
      const puts = h.callsFor("PUT");
      expect(puts.length).toBe(1);
      expect(puts[0]?.token).toBe("tok-b");
      expect(puts[0]?.headers.get("X-COOL-WOPI-IsAutosave")).toBe("false");
      expect(savedEvents.length).toBe(1);
      expect(savedEvents[0]?.by).toBe(b.user.id);
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
      await session.leave(b.user.id).catch(() => {});
    }
  });

  test("token expiring event fires five minutes before expiry", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a", { expiresAt: ttl(h.clock, 20 * 60000) });
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: ttl(h.clock, 20 * 60000), mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const seen: unknown[] = [];
    session.on("tokenExpiring", (e) => seen.push(e));
    try {
      for (let i = 0; i < 28; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
      }
      expect(seen.length).toBe(0);
      await h.clock.advance(30000);
      session.heartbeat(a.user.id);
      await h.clock.advance(30000);
      expect(seen.length).toBe(1);
      expect((seen[0] as { userId?: string }).userId).toBe(a.user.id);
    } finally {
      await session.leave(a.user.id).catch(() => {});
    }
  });
});

describe("last leave, disconnect and hold", () => {
  test("stale connection generations cannot disconnect or heartbeat the current connection", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const b = await joinEditor(h, "tok-b");
    const session = a.session;
    try {
      const stale = session.connect(a.user.id);
      const current = session.connect(a.user.id);
      expect(current).toBeGreaterThan(stale);
      await session.disconnect(a.user.id, stale);
      expect(a.user.connected).toBe(true);
      const lastSeen = a.user.lastSeen;
      await h.clock.advance(1000);
      session.heartbeat(a.user.id, stale);
      expect(a.user.lastSeen).toBe(lastSeen);
      await session.disconnect(a.user.id, current);
      expect(a.user.connected).toBe(false);
      session.heartbeat(a.user.id, stale);
      expect(a.user.connected).toBe(false);
      session.heartbeat(a.user.id, current);
      expect(a.user.connected).toBe(true);
      expect(a.user.lastSeen).toBe(h.clock.t);
      expect(b.user.connected).toBe(true);
      expect(session.lockHeld).toBe(true);
    } finally {
      session.abandon();
    }
  });

  test("last leave flushes edits after the last autosave before unlock", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "early-edit ", 0);
      await h.clock.advance(30000);
      expect(h.putCount()).toBe(1);
      edit(session, a.user.id, client, "late-edit ", 1);
      await session.leave(a.user.id);
      expect(await savedContains(h, "late-edit")).toBe(true);
      const putIndex = h.calls.map((c) => c.override).lastIndexOf("PUT");
      const unlockIndex = h.calls.map((c) => c.override).lastIndexOf("UNLOCK");
      expect(putIndex).toBeGreaterThanOrEqual(0);
      expect(unlockIndex).toBeGreaterThan(putIndex);
      const exitPut = h.calls[putIndex]!;
      expect(exitPut.headers.get("X-COOL-WOPI-IsExitSave")).toBe("true");
      expect(session.state).toBe("closed");
      expect(h.registry.list().length).toBe(0);
    } finally {
      client.destroy();
    }
  });

  test("expired leaver token falls back to another valid token", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const b = await joinEditor(h, "tok-b");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "fallback-edit ", 0);
      a.user.expiresAt = h.clock.t - 1000;
      b.user.expiresAt = h.clock.t - 1000;
      await session.disconnect(b.user.id);
      b.user.expiresAt = h.clock.t + 3600000;
      await session.leave(a.user.id);
      expect(await savedContains(h, "fallback-edit")).toBe(true);
      const puts = h.callsFor("PUT");
      expect(puts[puts.length - 1]?.token).toBe("tok-b");
    } finally {
      client.destroy();
      await session.leave(b.user.id).catch(() => {});
    }
  });

  test("no valid token holds the session, a returning writer resumes and flushes", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "held-edit ", 0);
      a.user.expiresAt = h.clock.t - 1000;
      await session.leave(a.user.id);
      expect(session.state).not.toBe("closed");
      expect(h.callsFor("UNLOCK").length).toBe(0);
      const b = await joinEditor(h, "tok-b");
      expect(b.joinPayload).toBeTruthy();
      expect(await savedContains(h, "held-edit")).toBe(true);
      await session.leave(b.user.id);
      expect(session.state).toBe("closed");
    } finally {
      client.destroy();
    }
  });

  test("after thirty minutes with no valid token lostFlush fires and the session closes", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    const lost: unknown[] = [];
    session.on("lostFlush", (e) => lost.push(e));
    try {
      edit(session, a.user.id, client, "lost-edit ", 0);
      a.user.expiresAt = h.clock.t - 1000;
      await session.leave(a.user.id);
      await h.clock.advance(30 * 60000 + 1000);
      expect(lost.length).toBe(1);
      expect(JSON.stringify(lost[0])).not.toContain("lost-edit");
      expect(session.state).toBe("closed");
    } finally {
      client.destroy();
    }
  });

  test("disconnect without leave after sixty seconds without heartbeat counts as leaving", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "swept-edit ", 0);
      await h.clock.advance(61000);
      expect(a.user.connected).toBe(false);
      expect(await savedContains(h, "swept-edit")).toBe(true);
      expect(session.state).toBe("closed");
    } finally {
      client.destroy();
    }
  });

  test("heartbeat keeps a user connected", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    try {
      await h.clock.advance(50000);
      session.heartbeat(a.user.id);
      await h.clock.advance(50000);
      expect(a.user.connected).toBe(true);
      expect(session.state).toBe("open");
    } finally {
      await session.leave(a.user.id).catch(() => {});
    }
  });
});

describe("locks, conflicts and permissions", () => {
  test("lock held by another app opens view only", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    h.host.files.get("file1")!.lock = "other-app-lock";
    h.host.files.get("file1")!.lockExpiresAt = h.clock.t + 30 * 60000;
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    try {
      expect(a.user.canWrite).toBe(false);
      expect(a.file.readOnly).toBe(true);
      expect(a.viewOnlyReason).toBe("lockedElsewhere");
      expect(session.viewOnlyReason).toBe("lockedElsewhere");
      const client = await clientFor(session, a.joinPayload);
      try {
        client.edit("nope ", 0);
        const replica = session.replica!;
        expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted).toBe(false);
      } finally {
        client.destroy();
      }
    } finally {
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("stale own prefix lock is reclaimed with unlock and relock", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    h.host.files.get("file1")!.lock = "bo1:old-instance:old-session";
    h.host.files.get("file1")!.lockExpiresAt = h.clock.t + 30 * 60000;
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    try {
      expect(a.user.canWrite).toBe(true);
      const relocks = h.calls.filter((c) => c.override === "LOCK" && c.headers.get("X-WOPI-OldLock") === "bo1:old-instance:old-session");
      expect(relocks.length).toBe(1);
      expect(h.host.files.get("file1")!.lock).toBe(session.lockId);
      expect(session.lockId).toMatch(/^bo1:inst-test:/);
    } finally {
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("timestamp conflict saves a conflict copy and leaves the original untouched", async () => {
    const h = makeHarness({ quirks: { timestampConflict: true, relativeUrlToken: true } });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    const conflicts: Array<{ name: string; copyName: string }> = [];
    session.on("conflict", (e) => conflicts.push(e));
    const originalBytes = h.host.files.get("file1")!.bytes.slice();
    try {
      edit(session, a.user.id, client, "conflict-edit ", 0);
      h.host.files.get("file1")!.lastModifiedTime = "1999-01-01T00:00:00.000Z";
      const saved = await session.requestSave(a.user.id);
      expect(conflicts.length).toBe(1);
      expect(conflicts[0]?.name).toMatch(/\(conflict \d{4}-\d{2}-\d{2} \d{2}\.\d{2}\)\.xlsx$/);
      expect(saved.conflictName).toBe(conflicts[0]?.name);
      expect(saved).not.toHaveProperty("conflictUrl");
      expect(conflicts[0]).not.toHaveProperty("url");
      expect(conflicts[0]).not.toHaveProperty("copyUrl");
      expect(h.host.files.get("file1")!.bytes).toEqual(originalBytes);
      const copies = [...h.host.files.values()].filter((f) => f.name.includes("(conflict"));
      expect(copies.length).toBe(1);
      const fresh = await openFreshClient("xlsx", copies[0]!.bytes, nextId());
      try {
        expect(fresh.contains("conflict-edit")).toBe(true);
      } finally {
        fresh.destroy();
      }
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("view joins and viewer tokens cannot apply updates", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-v", "user-v", { canWrite: false });
    const a = await joinEditor(h, "tok-a");
    const v = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-v", accessTokenTtl: 0, mode: "edit" });
    const viewMode = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "view" });
    const session = [...h.registry.list()][0]!;
    try {
      expect(v.user.canWrite).toBe(false);
      expect(viewMode.user.canWrite).toBe(false);
      expect(session.applyUpdate(v.user.id, new Uint8Array([1, 2, 3])).accepted).toBe(false);
      expect(session.applyUpdate(viewMode.user.id, new Uint8Array([1, 2, 3])).accepted).toBe(false);
    } finally {
      await session.leave(a.user.id).catch(() => {});
      await session.leave(v.user.id).catch(() => {});
      await session.leave(viewMode.user.id).catch(() => {});
    }
  });

  test("maxEditorsPerFile is enforced", async () => {
    const h = makeHarness({ maxEditorsPerFile: 2 });
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    addToken(h, "tok-c", "user-c");
    await joinEditor(h, "tok-a");
    await joinEditor(h, "tok-b");
    await expect(joinEditor(h, "tok-c")).rejects.toThrow();
    const session = [...h.registry.list()][0]!;
    expect(session.users.length).toBe(2);
  });

  test("a distinctive token never appears in logs or events", async () => {
    const h = makeHarness({ quirks: { timestampConflict: true, relativeUrlToken: true } });
    const secret = "distinctive-secret-token-xyz";
    addToken(h, secret, "user-a");
    const a = await joinEditor(h, secret);
    const session = [...h.registry.list()][0]!;
    const events: string[] = [];
    for (const name of ["leaderChanged", "saved", "saveFailed", "conflict", "viewOnly", "closed", "lostFlush", "tokenExpiring", "userInvalidated"] as const) {
      session.on(name, (e) => events.push(JSON.stringify(e)));
    }
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "redact-edit ", 0);
      h.host.files.get("file1")!.lastModifiedTime = "1999-01-01T00:00:00.000Z";
      await session.requestSave(a.user.id);
      await session.leave(a.user.id);
      const combined = [...h.logs, ...events, JSON.stringify(a.user), JSON.stringify(session.users)].join("\n");
      expect(combined).not.toContain(secret);
      expect(combined).not.toContain("access_token=" + secret);
    } finally {
      client.destroy();
    }
  });
});

describe("review fixes: expiry, fallback and departure races", () => {
  test("sole editor in the last two minutes still autosaves with their valid token", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a", { expiresAt: ttl(h.clock, 10 * 60000) });
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: ttl(h.clock, 10 * 60000), mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      for (let i = 0; i < 16; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
      }
      await h.clock.advance(1000);
      session.heartbeat(a.user.id);
      expect(session.leader).toBeNull();
      edit(session, a.user.id, client, "last-minutes ", 0);
      await h.clock.advance(30000);
      expect(h.putCount()).toBe(1);
      expect(h.callsFor("PUT")[0]?.token).toBe("tok-a");
      expect(await savedContains(h, "last-minutes")).toBe(true);
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("an expired connected editor does not block flush when a valid editor leaves", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a", { expiresAt: ttl(h.clock, 5 * 60000) });
    addToken(h, "tok-b", "user-b", { expiresAt: ttl(h.clock, 10 * 60000) });
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: ttl(h.clock, 5 * 60000), mode: "edit" });
    const b = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: ttl(h.clock, 10 * 60000), mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, b.joinPayload);
    try {
      for (let i = 0; i < 12; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
        session.heartbeat(b.user.id);
      }
      expect(a.user.canWrite).toBe(false);
      edit(session, b.user.id, client, "valid-leaver ", 0);
      await session.leave(b.user.id);
      expect(await savedContains(h, "valid-leaver")).toBe(true);
      expect(h.callsFor("PUT").some((c) => c.token === "tok-b")).toBe(true);
      expect(session.state).toBe("closed");
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("a 401 on save marks the token dead and retries with another token", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const b = await joinEditor(h, "tok-b");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      await session.leave(b.user.id);
      edit(session, a.user.id, client, "revoked-edit ", 0);
      const entry = h.host.tokens.get("tok-a");
      if (typeof entry === "object" && entry !== null) entry.expiresAt = h.clock.t - 1000;
      await session.leave(a.user.id);
      expect(await savedContains(h, "revoked-edit")).toBe(true);
      const puts = h.callsFor("PUT");
      expect(puts[puts.length - 1]?.token).toBe("tok-b");
      expect(a.user.expiresAt).not.toBeNull();
      expect(a.user.expiresAt!).toBeLessThanOrEqual(h.clock.t);
      expect(session.state).toBe("closed");
      expect(session.awaitingFlushToken).toBe(false);
    } finally {
      client.destroy();
    }
  });

  test("reconnect and edit during an in flight exit put keeps the session open and flushes", async () => {
    let releasePut: () => void = () => {};
    let enteredPut: (() => void) | null = null;
    const putEntered = new Promise<void>((resolve) => {
      enteredPut = resolve;
    });
    let gated = true;
    const h = makeHarness({
      beforePut: async () => {
        if (!gated) return;
        gated = false;
        enteredPut?.();
        await new Promise<void>((resolve) => {
          releasePut = resolve;
        });
      },
    });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "first-edit ", 0);
      const leaving = session.disconnect(a.user.id);
      await putEntered;
      session.heartbeat(a.user.id);
      expect(a.user.connected).toBe(true);
      edit(session, a.user.id, client, "second-edit ", 1);
      releasePut();
      await leaving;
      expect(session.state).toBe("open");
      expect(session.lockHeld).toBe(true);
      expect(h.host.files.get("file1")!.lock).toBe(session.lockId);
      await session.requestSave(a.user.id);
      expect(await savedContains(h, "second-edit")).toBe(true);
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("a failing first lock leaves no zombie session", async () => {
    const h = makeHarness({ lockFails: true });
    addToken(h, "tok-a", "user-a");
    await expect(joinEditor(h, "tok-a")).rejects.toThrow();
    expect(h.registry.list().length).toBe(0);
  });

  test("view only viewers who time out close the session", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    h.host.files.get("file1")!.lock = "other-app-lock";
    h.host.files.get("file1")!.lockExpiresAt = h.clock.t + 30 * 60000;
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    await h.clock.advance(61000);
    expect(a.user.connected).toBe(false);
    expect(session.state).toBe("closed");
    expect(h.registry.list().length).toBe(0);
  });

  test("a concurrent first join survives another join failing", async () => {
    const h = makeHarness();
    addToken(h, "tok-b", "user-b");
    h.host.tokens.set("tok-bad", { userId: "bad", canWrite: true, expiresAt: h.clock.t - 1000 });
    const results = await Promise.allSettled([
      h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-bad", accessTokenTtl: 0, mode: "edit" }),
      h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" }),
    ]);
    expect(results[0]?.status).toBe("rejected");
    expect(results[1]?.status).toBe("fulfilled");
    expect(h.registry.list().length).toBe(1);
    const ok = results[1]?.status === "fulfilled" ? results[1].value : null;
    if (ok) await ok.session.leave(ok.user.id).catch(() => {});
  });

  test("disconnected users do not consume editor slots", async () => {
    const h = makeHarness({ maxEditorsPerFile: 2 });
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    addToken(h, "tok-c", "user-c");
    const a = await joinEditor(h, "tok-a");
    const b = await joinEditor(h, "tok-b");
    const session = [...h.registry.list()][0]!;
    await session.disconnect(b.user.id);
    const c = await joinEditor(h, "tok-c");
    expect(c.user.canWrite).toBe(true);
    await session.leave(a.user.id).catch(() => {});
    await session.leave(c.user.id).catch(() => {});
  });

  test("timer delays are clamped to the 32 bit limit for long token ttl", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a", { expiresAt: ttl(h.clock, 40 * 24 * 3600000) });
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: ttl(h.clock, 40 * 24 * 3600000), mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const seen: unknown[] = [];
    session.on("tokenExpiring", (e) => seen.push(e));
    try {
      expect(h.timerDelays.every((ms) => ms <= 2147483647)).toBe(true);
      for (let i = 0; i < 120; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
      }
      expect(seen.length).toBe(0);
    } finally {
      await session.leave(a.user.id).catch(() => {});
    }
  });
});

describe("review round 2: serialized transitions", () => {
  test("a join in flight while the last editor leaves keeps one live session and no lost edits", async () => {
    let releasePut: () => void = () => {};
    let enteredPut: (() => void) | null = null;
    const putEntered = new Promise<void>((resolve) => {
      enteredPut = resolve;
    });
    let gated = true;
    const h = makeHarness({
      beforePut: async () => {
        if (!gated) return;
        gated = false;
        enteredPut?.();
        await new Promise<void>((resolve) => {
          releasePut = resolve;
        });
      },
    });
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const firstSession = [...h.registry.list()][0]!;
    const clientA = await clientFor(firstSession, a.joinPayload);
    let clientB: TestClient | null = null;
    try {
      edit(firstSession, a.user.id, clientA, "before-leave ", 0);
      const leaving = firstSession.leave(a.user.id);
      await putEntered;
      const joining = h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
      releasePut();
      const b = await joining;
      await leaving;
      expect(h.registry.list().length).toBe(1);
      const live = [...h.registry.list()][0]!;
      expect(live.state).toBe("open");
      expect(h.host.files.get("file1")!.lock).toBe(live.lockId);
      clientB = await clientFor(live, b.joinPayload, h.host.files.get("file1")!.bytes);
      expect(clientB.contains("before-leave")).toBe(true);
      edit(live, b.user.id, clientB, "after-join ", 1);
      await live.requestSave(b.user.id);
      expect(await savedContains(h, "before-leave")).toBe(true);
      expect(await savedContains(h, "after-join")).toBe(true);
      await live.leave(b.user.id);
    } finally {
      clientA.destroy();
      clientB?.destroy();
    }
  });

  test("a reconnect plus edit while the exit unlock is in flight keeps the session and lock", async () => {
    let releaseUnlock: () => void = () => {};
    let enteredUnlock: (() => void) | null = null;
    const unlockEntered = new Promise<void>((resolve) => {
      enteredUnlock = resolve;
    });
    let gated = true;
    const h = makeHarness({
      beforeUnlock: async () => {
        if (!gated) return;
        gated = false;
        enteredUnlock?.();
        await new Promise<void>((resolve) => {
          releaseUnlock = resolve;
        });
      },
    });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "saved-first ", 0);
      await session.requestSave(a.user.id);
      expect(session.dirty).toBe(false);
      const leaving = session.disconnect(a.user.id);
      await unlockEntered;
      session.heartbeat(a.user.id);
      expect(a.user.connected).toBe(true);
      edit(session, a.user.id, client, "during-unlock ", 1);
      releaseUnlock();
      await leaving;
      expect(session.state).toBe("open");
      expect(session.lockHeld).toBe(true);
      expect(h.host.files.get("file1")!.lock).toBe(session.lockId);
      await session.requestSave(a.user.id);
      expect(await savedContains(h, "during-unlock")).toBe(true);
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("a resume from hold at twenty five minutes refreshes the lock right away", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "held-again ", 0);
      a.user.expiresAt = h.clock.t - 1000;
      await session.leave(a.user.id);
      expect(session.awaitingFlushToken).toBe(true);
      await h.clock.advance(25 * 60000);
      expect(session.state).toBe("open");
      h.calls.length = 0;
      const b = await joinEditor(h, "tok-b");
      expect(b.user.canWrite).toBe(true);
      const refreshes = h.callsFor("REFRESH_LOCK");
      expect(refreshes.length).toBeGreaterThanOrEqual(1);
      expect(refreshes[0]?.token).toBe("tok-b");
      expect(h.host.files.get("file1")!.lockExpiresAt).toBeGreaterThan(h.clock.t + 20 * 60000);
      expect(await savedContains(h, "held-again")).toBe(true);
      await session.leave(b.user.id);
      expect(session.state).toBe("closed");
    } finally {
      client.destroy();
    }
  });

  test("two concurrent first joins where the first fails in readFile still open the session", async () => {
    const h = makeHarness({ failFirstRead: true });
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const results = await Promise.allSettled([
      h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" }),
      h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" }),
    ]);
    expect(results[0]?.status).toBe("rejected");
    expect(results[1]?.status).toBe("fulfilled");
    expect(h.registry.list().length).toBe(1);
    const live = [...h.registry.list()][0]!;
    expect(live.state).toBe("open");
    expect(h.calls.filter((c) => c.pathname.endsWith("/contents") && c.override === null).length).toBe(2);
    const ok = results[1]?.status === "fulfilled" ? results[1].value : null;
    if (ok) await live.leave(ok.user.id).catch(() => {});
  });
});

describe("replica api adaptation", () => {
  test("an accepted docx edit marks dirty and autosaves after thirty seconds", async () => {
    const h = makeHarness({ format: "docx" });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await openTestClient(a.joinPayload, sourceBytes("docx"), nextId());
    const failures: unknown[] = [];
    session.on("saveFailed", (e) => failures.push(e));
    try {
      client.edit("autosave-docx ", 0);
      const replica = session.replica!;
      const result = session.applyUpdate(a.user.id, client.encodeStateAsUpdate(replica.encodeStateVector()));
      expect(result).toEqual({ accepted: true, followUps: [] });
      expect(session.replica!.version).toBe(1);
      expect(session.dirty).toBe(true);
      await h.clock.advance(30000);
      await h.clock.drain();
      expect(failures).toEqual([]);
      expect(h.putCount()).toBe(1);
      expect(session.dirty).toBe(false);
      const fresh = await openFreshClient("docx", h.host.files.get("file1")!.bytes, nextId());
      try {
        expect(fresh.contains("autosave-docx")).toBe(true);
      } finally {
        fresh.destroy();
      }
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });


  test("a docx save that publishes paragraph ids emits replicaUpdate once and stays clean", async () => {
    const h = makeHarness({ format: "docx" });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await openTestClient(a.joinPayload, sourceBytes("docx"), nextId());
    const updates: Uint8Array[] = [];
    session.on("replicaUpdate", (e) => updates.push(e.update));
    try {
      client.splitFirstParagraph();
      const replica = session.replica!;
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted).toBe(true);
      expect(session.dirty).toBe(true);
      const saved = await session.requestSave(a.user.id);
      expect(saved.saved).toBe(true);
      expect(updates.length).toBe(1);
      expect(updates[0]!.byteLength).toBeGreaterThan(0);
      expect(session.dirty).toBe(false);
      const again = await session.requestSave(a.user.id);
      expect(again.saved).toBe(false);
      expect(updates.length).toBe(1);
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });

  test("a causalGap result passes through and does not mark dirty", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    try {
      client.edit("gap-first ", 0);
      const afterFirst = client.encodeStateVector();
      const replica = session.replica!;
      const first = client.encodeStateAsUpdate(replica.encodeStateVector());
      client.edit("gap-second ", 1);
      const second = client.encodeStateAsUpdate(afterFirst);
      expect(session.applyUpdate(a.user.id, second)).toEqual({ accepted: false, reason: "causalGap" });
      expect(session.dirty).toBe(false);
      expect(session.replica!.version).toBe(0);
      expect(session.applyUpdate(a.user.id, first)).toEqual({ accepted: true, followUps: [] });
      expect(session.dirty).toBe(true);
      await session.requestSave(a.user.id);
      expect(session.dirty).toBe(false);
      expect(session.applyUpdate(a.user.id, first)).toEqual({ accepted: true, followUps: [] });
      expect(session.dirty).toBe(false);
    } finally {
      client.destroy();
      await session.leave(a.user.id).catch(() => {});
    }
  });
});

describe("session safety regressions", () => {
  test("a reconnect relock validates the source before accepting more updates", async () => {
    let releaseUnlock: () => void = () => {};
    const unlockGate = new Promise<void>((resolve) => { releaseUnlock = resolve; });
    let enteredUnlock: () => void = () => {};
    const unlockEntered = new Promise<void>((resolve) => { enteredUnlock = resolve; });
    let releaseCheck: () => void = () => {};
    const checkGate = new Promise<void>((resolve) => { releaseCheck = resolve; });
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-external", "external-app");
    const a = await joinEditor(h, "tok-a");
    const session = a.session;
    const client = await clientFor(session, a.joinPayload);
    const external = await openFreshClient("xlsx", sourceBytes("xlsx"), nextId());
    const unlock = h.hostClient.unlock;
    const checkFileInfo = h.hostClient.checkFileInfo;
    let checkRelock = false;
    let validationStarted = false;
    try {
      external.edit("external-between-locks ");
      const externalBytes = await external.save();
      h.hostClient.unlock = async (input) => {
        enteredUnlock();
        await unlockGate;
        await unlock(input);
        const externalLock = { wopiSrc: SRC, accessToken: "tok-external", lock: "external-app-lock" };
        await h.hostClient.lock(externalLock);
        await h.hostClient.putFile({ ...externalLock, body: externalBytes });
        await unlock(externalLock);
        checkRelock = true;
      };
      h.hostClient.checkFileInfo = async (input) => {
        if (checkRelock) {
          checkRelock = false;
          validationStarted = true;
          await checkGate;
        }
        return checkFileInfo(input);
      };
      const leaving = session.disconnect(a.user.id);
      await unlockEntered;
      session.heartbeat(a.user.id);
      edit(session, a.user.id, client, "reconnected-before-relock ");
      releaseUnlock();
      await h.clock.drain();
      expect(validationStarted).toBe(true);
      client.edit("blocked-during-validation ", 1);
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector())).accepted).toBe(false);
      releaseCheck();
      await leaving;
      expect(session.viewOnlyReason).toBe("externalChange");
      const result = await session.requestSave(a.user.id);
      expect(result.conflictName).toBeTruthy();
      expect(h.host.files.get("file1")!.bytes).toEqual(externalBytes);
      const copy = [...h.host.files.values()].find((file) => file.id !== "file1")!;
      const fresh = await openFreshClient("xlsx", copy.bytes, nextId());
      try {
        expect(fresh.contains("reconnected-before-relock")).toBe(true);
        expect(fresh.contains("blocked-during-validation")).toBe(false);
      } finally {
        fresh.destroy();
      }
    } finally {
      releaseUnlock();
      releaseCheck();
      client.destroy();
      external.destroy();
      await session.close();
    }
  });

  for (const save of ["autosave", "last leave"] as const) {
    test(`a reconnect relock failure recovers and ${save} writes the edit`, async () => {
      let releaseUnlock: () => void = () => {};
      const unlockGate = new Promise<void>((resolve) => { releaseUnlock = resolve; });
      let enteredUnlock: () => void = () => {};
      const unlockEntered = new Promise<void>((resolve) => { enteredUnlock = resolve; });
      const h = makeHarness({ beforeUnlock: async () => { enteredUnlock(); await unlockGate; } });
      addToken(h, "tok-a", "user-a");
      const a = await joinEditor(h, "tok-a");
      const session = a.session;
      const client = await clientFor(session, a.joinPayload);
      const notices: unknown[] = [];
      session.on("viewOnly", (event) => notices.push(event));
      const lock = h.hostClient.lock;
      let attempts = 0;
      h.hostClient.lock = async (input) => {
        attempts += 1;
        if (attempts === 1) throw new WopiHostError("server", "relock failed", { status: 500 });
        await lock(input);
      };
      try {
        const leaving = session.disconnect(a.user.id);
        await unlockEntered;
        session.heartbeat(a.user.id);
        edit(session, a.user.id, client, "retry-reconnect-edit ");
        releaseUnlock();
        await leaving;
        await h.clock.advance(1000);
        expect(notices).toContainEqual({ reason: "lockLost" });
        expect(attempts).toBe(2);
        expect(session.lockHeld).toBe(true);
        expect(session.viewOnly).toBe(false);
        if (save === "autosave") await h.clock.advance(30000);
        else await session.leave(a.user.id);
        expect(await savedContains(h, "retry-reconnect-edit")).toBe(true);
        expect(h.putCount()).toBe(1);
        expect(h.callsFor("PUT")[0]?.headers.get("X-COOL-WOPI-IsExitSave")).toBe(save === "last leave" ? "true" : "false");
      } finally {
        releaseUnlock();
        client.destroy();
        await session.close();
      }
    });
  }

  test("a delayed refresh response does not extend the lease beyond thirty minutes from its request", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const b = await joinEditor(h, "tok-b");
    const session = a.session;
    const client = await clientFor(session, b.joinPayload);
    const refreshLock = h.hostClient.refreshLock;
    let releaseRefresh: () => void = () => {};
    const refreshGate = new Promise<void>((resolve) => { releaseRefresh = resolve; });
    let enteredRefresh: () => void = () => {};
    const refreshEntered = new Promise<void>((resolve) => { enteredRefresh = resolve; });
    let requestSentAt = 0;
    h.hostClient.refreshLock = async (input) => {
      if (requestSentAt) throw new WopiHostError("server", "refresh failed", { status: 500 });
      requestSentAt = h.clock.t;
      await refreshLock(input);
      enteredRefresh();
      await refreshGate;
    };
    h.hostClient.lock = async () => { throw new WopiHostError("server", "lock failed", { status: 500 }); };
    try {
      const leaving = session.leave(a.user.id);
      await refreshEntered;
      for (let i = 0; i < 40; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(b.user.id);
      }
      releaseRefresh();
      await leaving;
      expect(h.clock.t).toBe(requestSentAt + 20 * 60000);
      expect(session.lockHeld).toBe(true);
      expect(h.host.files.get("file1")!.lockExpiresAt).toBe(requestSentAt + 30 * 60000);
      for (let i = 0; i < 20; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(b.user.id);
      }
      expect(h.clock.t).toBe(requestSentAt + 30 * 60000);
      expect(session.lockHeld).toBe(false);
      expect(session.viewOnlyReason).toBe("lockLost");
      client.edit("blocked-after-host-lease ");
      expect(session.applyUpdate(b.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector())).accepted).toBe(false);
    } finally {
      releaseRefresh();
      client.destroy();
      await session.close();
    }
  });

  for (const kind of ["unauthorized", "server"] as const) {
    test(`a post-save ${kind} error preserves the save and handles token validity`, async () => {
      const h = makeHarness();
      addToken(h, "tok-a", "user-a");
      const a = await joinEditor(h, "tok-a");
      const session = a.session;
      const client = await clientFor(session, a.joinPayload);
      const saved: unknown[] = [];
      const failed: unknown[] = [];
      session.on("saved", (event) => saved.push(event));
      session.on("saveFailed", (event) => failed.push(event));
      h.hostClient.checkFileInfo = async () => {
        throw new WopiHostError(kind, "post-save check failed", { status: kind === "unauthorized" ? 401 : 500 });
      };
      try {
        edit(session, a.user.id, client, "saved-before-check-failed ");
        const result = await session.requestSave(a.user.id);
        expect(result).toEqual({ saved: true, itemVersion: "2", by: a.user.id });
        expect(session.itemVersion).toBe("2");
        expect(session.dirty).toBe(false);
        expect(h.putCount()).toBe(1);
        expect(await savedContains(h, "saved-before-check-failed")).toBe(true);
        expect(saved).toEqual([{ itemVersion: "2", by: a.user.id }]);
        expect(failed).toEqual([]);
        expect(a.user.canWrite).toBe(kind !== "unauthorized");
        expect(a.user.isTokenValidAt(h.clock.t)).toBe(kind !== "unauthorized");
        expect(session.leader?.id ?? null).toBe(kind === "unauthorized" ? null : a.user.id);
      } finally {
        client.destroy();
        await session.close();
      }
    });
  }

  test("an unauthorized token invalidates both its edit and view handles", async () => {
    const h = makeHarness();
    addToken(h, "tok-shared", "shared-user");
    addToken(h, "tok-other", "other-user");
    const writer = await joinEditor(h, "tok-shared");
    const viewer = await joinEditor(h, "tok-shared", "view");
    const fallback = await joinEditor(h, "tok-other");
    const session = writer.session;
    const client = await clientFor(session, writer.joinPayload);
    const invalidated: unknown[] = [];
    session.on("userInvalidated", (event) => invalidated.push(event));
    try {
      edit(session, writer.user.id, client, "shared-revoked-edit ");
      h.host.tokens.delete("tok-shared");
      const result = await session.requestSave(writer.user.id);
      expect(result.saved).toBe(true);
      expect(result.by).toBe(fallback.user.id);
      expect(h.callsFor("PUT").map((call) => call.token)).toEqual(["tok-shared", "tok-other"]);
      expect(await savedContains(h, "shared-revoked-edit")).toBe(true);
      expect(invalidated).toEqual([
        { userId: writer.user.id, reason: "unauthorized" },
        { userId: viewer.user.id, reason: "unauthorized" },
      ]);
      for (const user of [writer.user, viewer.user]) {
        expect(user.canWrite).toBe(false);
        expect(user.isTokenValidAt(h.clock.t)).toBe(false);
      }
      expect(fallback.user.canWrite).toBe(true);
      expect(fallback.user.isTokenValidAt(h.clock.t)).toBe(true);
      expect(session.leader?.id).toBe(fallback.user.id);
      expect(JSON.stringify(invalidated)).not.toContain("tok-shared");
    } finally {
      client.destroy();
      await session.close();
    }
  });

  test("conflict results and events omit a token-bearing copy URL", async () => {
    const secret = "copy-only-distinctive-secret-xyz";
    const h = makeHarness({ quirks: { timestampConflict: true, relativeUrlToken: secret } });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = a.session;
    const client = await clientFor(session, a.joinPayload);
    const events: unknown[] = [];
    session.on("conflict", (event) => events.push(event));
    try {
      edit(session, a.user.id, client, "private-copy ");
      h.host.files.get("file1")!.lastModifiedTime = "1999-01-01T00:00:00.000Z";
      const result = await session.requestSave(a.user.id);
      expect(result.saved).toBe(true);
      expect(result.conflictName).toBeTruthy();
      expect(result.conflictUrl).toBeUndefined();
      expect(events).toEqual([{ name: result.conflictName, copyName: result.conflictName }]);
      expect(JSON.stringify({ events, result, logs: h.logs })).not.toContain(secret);
      expect(JSON.stringify({ events, result, logs: h.logs })).not.toContain("access_token");
    } finally {
      client.destroy();
      await session.close();
    }
  });

  for (const changed of ["version", "timestamp"] as const) {
    test(`first writer reloads a viewer's source when the ${changed} changes under the lock`, async () => {
      let onCheck = async () => {};
      const h = makeHarness({ beforeCheckFileInfo: () => onCheck() });
      addToken(h, "tok-v", "viewer", { canWrite: false });
      addToken(h, "tok-a", "writer");
      const viewer = await joinEditor(h, "tok-v", "view");
      const session = viewer.session;
      const baselineGeneration = session.baselineGeneration;
      const events: unknown[] = [];
      session.on("rejoinRequired", (event) => events.push(event));
      const external = await openFreshClient("xlsx", sourceBytes("xlsx"), nextId());
      let client: TestClient | null = null;
      try {
        external.edit("external-before-lock ");
        const bytes = await external.save();
        let changedUnderLock = false;
        onCheck = async () => {
          const file = h.host.files.get("file1")!;
          if (file.lock !== session.lockId || changedUnderLock) return;
          changedUnderLock = true;
          file.bytes = bytes;
          if (changed === "version") file.itemVersion = "2";
          else file.lastModifiedTime = "2026-10-28T12:00:00.000Z";
        };
        const writer = await joinEditor(h, "tok-a");
        expect(changedUnderLock).toBe(true);
        expect(session.baselineGeneration).toBe(baselineGeneration + 1);
        expect(events).toEqual([{ reason: "sourceChanged" }]);
        client = await clientFor(session, writer.joinPayload, bytes);
        expect(client.contains("external-before-lock")).toBe(true);
        edit(session, writer.user.id, client, "writer-after-lock ", 1);
        await session.requestSave(writer.user.id);
        expect(await savedContains(h, "external-before-lock")).toBe(true);
        expect(await savedContains(h, "writer-after-lock")).toBe(true);
        expect(h.calls.filter((call) => call.pathname.endsWith("/contents") && call.override === null)).toHaveLength(2);
      } finally {
        external.destroy();
        client?.destroy();
        await session.close();
      }
    });
  }

  test("a host change after PutFile saves all session edits as a copy without adopting the external baseline", async () => {
    let onCheck = async () => {};
    const h = makeHarness({ beforeCheckFileInfo: () => onCheck() });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = a.session;
    const client = await clientFor(session, a.joinPayload);
    const external = await openFreshClient("xlsx", sourceBytes("xlsx"), nextId());
    const notices: unknown[] = [];
    const saved: unknown[] = [];
    session.on("viewOnly", (event) => notices.push(event));
    session.on("saved", (event) => saved.push(event));
    try {
      external.edit("external-after-put ");
      const externalBytes = await external.save();
      edit(session, a.user.id, client, "session-before-put ");
      let changed = false;
      onCheck = async () => {
        if (!h.putCount() || changed) return;
        changed = true;
        edit(session, a.user.id, client, "session-during-check ", 1);
        const file = h.host.files.get("file1")!;
        file.bytes = externalBytes;
        file.version = 3;
        file.itemVersion = "3";
        file.lastModifiedTime = "2026-10-28T12:00:00.000Z";
      };
      const result = await session.requestSave(a.user.id);
      expect(result.conflictName).toBeTruthy();
      expect(session.itemVersion).toBe("1");
      expect(session.viewOnlyReason).toBe("externalChange");
      expect(notices).toEqual([{ reason: "externalChange" }]);
      expect(saved).toEqual([]);
      expect(h.host.files.get("file1")!.bytes).toEqual(externalBytes);
      const copy = [...h.host.files.values()].find((file) => file.id !== "file1")!;
      const fresh = await openFreshClient("xlsx", copy.bytes, nextId());
      try {
        expect(fresh.contains("session-before-put")).toBe(true);
        expect(fresh.contains("session-during-check")).toBe(true);
      } finally {
        fresh.destroy();
      }
      client.edit("blocked-after-conflict ");
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector())).accepted).toBe(false);
      await session.requestSave(a.user.id);
      expect(h.putCount()).toBe(1);
    } finally {
      client.destroy();
      external.destroy();
      await session.close();
    }
  });

  test("failed lock refreshes back off and recover before the lease expires", async () => {
    let failing = true;
    const h = makeHarness({ refreshStatus: () => failing ? 500 : null });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = a.session;
    try {
      for (let i = 0; i < 20; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
      }
      expect(h.callsFor("REFRESH_LOCK")).toHaveLength(1);
      await h.clock.advance(1000);
      expect(h.callsFor("REFRESH_LOCK")).toHaveLength(2);
      await h.clock.advance(1999);
      expect(h.callsFor("REFRESH_LOCK")).toHaveLength(2);
      failing = false;
      await h.clock.advance(1);
      expect(h.callsFor("REFRESH_LOCK")).toHaveLength(3);
      expect(session.lockHeld).toBe(true);
      expect(session.viewOnly).toBe(false);
      expect(h.host.files.get("file1")!.lockExpiresAt).toBe(h.clock.t + 30 * 60000);
    } finally {
      await session.close();
    }
  });

  test("the lease deadline stops updates even while a refresh request is pending, then a re-lock resumes", async () => {
    const h = makeHarness({ refreshStatus: () => 500 });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = a.session;
    const client = await clientFor(session, a.joinPayload);
    const notices: unknown[] = [];
    session.on("viewOnly", (event) => notices.push(event));
    let releaseRefresh: () => void = () => {};
    try {
      for (let i = 0; i < 20; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
      }
      h.hostClient.refreshLock = () => new Promise<void>((resolve) => { releaseRefresh = resolve; });
      for (let i = 0; i < 40; i += 1) {
        await h.clock.advance(30000);
        session.heartbeat(a.user.id);
      }
      expect(session.lockHeld).toBe(false);
      expect(session.viewOnlyReason).toBe("lockLost");
      expect(notices).toEqual([{ reason: "lockLost" }]);
      client.edit("blocked-without-lock ");
      const update = client.encodeStateAsUpdate(session.replica!.encodeStateVector());
      expect(session.applyUpdate(a.user.id, update).accepted).toBe(false);
      releaseRefresh();
      await h.clock.flush();
      await h.clock.advance(30000);
      session.heartbeat(a.user.id);
      await h.clock.advance(30000);
      session.heartbeat(a.user.id);
      expect(session.lockHeld).toBe(true);
      expect(session.viewOnly).toBe(false);
      expect(h.callsFor("LOCK").length).toBeGreaterThanOrEqual(2);
      expect(session.applyUpdate(a.user.id, update).accepted).toBe(true);
    } finally {
      releaseRefresh();
      client.destroy();
      await session.close();
    }
  });

  test("a refresh lock mismatch stops updates immediately and recovers after the other lock clears", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await joinEditor(h, "tok-a");
    const session = a.session;
    const client = await clientFor(session, a.joinPayload);
    const notices: unknown[] = [];
    session.on("viewOnly", (event) => notices.push(event));
    try {
      const file = h.host.files.get("file1")!;
      file.lock = "other-app-lock";
      file.lockExpiresAt = h.clock.t + 30 * 60000;
      await joinEditor(h, "tok-b");
      await session.leave(a.user.id);
      expect(session.viewOnlyReason).toBe("lockLost");
      expect(notices).toEqual([{ reason: "lockLost" }]);
      const writer = session.users[0]!;
      client.edit("after-lock-loss ");
      const update = client.encodeStateAsUpdate(session.replica!.encodeStateVector());
      expect(session.applyUpdate(writer.id, update).accepted).toBe(false);
      file.lock = null;
      file.lockExpiresAt = null;
      await h.clock.advance(1000);
      expect(session.lockHeld).toBe(true);
      expect(session.viewOnly).toBe(false);
      expect(session.applyUpdate(writer.id, update).accepted).toBe(true);
    } finally {
      client.destroy();
      await session.close();
    }
  });
});
