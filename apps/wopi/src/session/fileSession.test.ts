import { beforeAll, describe, expect, test } from "bun:test";
import { initEngines } from "../replica/engines";
import { openFreshClient, openTestClient, sourceBytes, type TestClient } from "../../test/replica/clients";
import { addToken, makeHarness, SRC, ttl, type Harness } from "../../test/session/harness";
import type { FileSession } from "./fileSession";

beforeAll(async () => {
  await initEngines();
});

let cid = 50000;
function nextId(): number {
  cid += 11;
  return cid;
}

async function clientFor(session: FileSession, joinPayload: Parameters<typeof openTestClient>[0]): Promise<TestClient> {
  return openTestClient(joinPayload, sourceBytes("xlsx"), nextId());
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
      expect(h.callsFor("LOCK").length + h.callsFor("LOCK").length).toBeGreaterThanOrEqual(1);
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
    const h = makeHarness({ quirks: { timestampConflict: true } });
    addToken(h, "tok-a", "user-a");
    const a = await joinEditor(h, "tok-a");
    const session = [...h.registry.list()][0]!;
    const client = await clientFor(session, a.joinPayload);
    const conflicts: Array<{ name?: string; url?: string }> = [];
    session.on("conflict", (e) => conflicts.push(e as { name?: string }));
    const originalBytes = h.host.files.get("file1")!.bytes.slice();
    try {
      edit(session, a.user.id, client, "conflict-edit ", 0);
      h.host.files.get("file1")!.lastModifiedTime = "1999-01-01T00:00:00.000Z";
      await session.requestSave(a.user.id);
      expect(conflicts.length).toBe(1);
      expect(conflicts[0]?.name).toMatch(/\(conflict \d{4}-\d{2}-\d{2} \d{2}\.\d{2}\)\.xlsx$/);
      expect(conflicts[0]?.url).toBeTruthy();
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
    const h = makeHarness();
    const secret = "distinctive-secret-token-xyz";
    addToken(h, secret, "user-a");
    const a = await joinEditor(h, secret);
    const session = [...h.registry.list()][0]!;
    const events: string[] = [];
    for (const name of ["leaderChanged", "saved", "saveFailed", "conflict", "viewOnly", "closed", "lostFlush", "tokenExpiring"] as const) {
      session.on(name, (e) => events.push(JSON.stringify(e)));
    }
    const client = await clientFor(session, a.joinPayload);
    try {
      edit(session, a.user.id, client, "redact-edit ", 0);
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
