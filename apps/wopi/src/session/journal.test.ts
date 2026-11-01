import { beforeAll, describe, expect, test } from "bun:test";
import { appendFileSync, chmodSync, existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, symlinkSync, writeSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";
import { initEngines } from "../replica/engines";
import { openFreshClient, openTestClient, sourceBytes } from "../../test/replica/clients";
import { addToken, makeHarness, SRC, type Harness } from "../../test/session/harness";
import { FileJournal } from "./journal";
import { FileSession, LOCK_LEASE_MS } from "./fileSession";
import { SessionRegistry } from "./registry";
import type { Format } from "../replica/replica";

function restartRegistry(h: Harness, journalDir: string): SessionRegistry {
  return new SessionRegistry({
    config: { ...h.config, journalDir },
    hostClient: h.hostClient,
    logger: h.logger,
    now: h.clock.now,
    scheduler: h.scheduler,
  });
}

beforeAll(async () => {
  await initEngines();
});

describe("FileJournal", () => {
  test("creates a private journal directory and file with a permissive umask", () => {
    const root = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const dir = join(root, "private");
    const journal = new FileJournal(dir, "file-key");
    const mask = process.umask(0o022);
    try {
      journal.begin(new Uint8Array([1, 2, 3]), "1");
      if (process.platform !== "win32") {
        expect(statSync(dir).mode & 0o777).toBe(0o700);
        expect(statSync(journal.path).mode & 0o777).toBe(0o600);
      }
    } finally {
      process.umask(mask);
      journal.close();
      rmSync(root, { recursive: true, force: true });
    }
  });

  for (const operation of ["read", "begin"] as const) {
    test(`tightens an existing journal directory and file before ${operation}`, () => {
      const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
      const journal = new FileJournal(dir, "file-key");
      const source = new Uint8Array([1, 2, 3]);
      const mask = process.umask(0o022);
      let reopened: FileJournal | null = null;
      try {
        journal.begin(source, "1");
        journal.close();
        chmodSync(dir, 0o755);
        chmodSync(journal.path, 0o644);
        reopened = new FileJournal(dir, "file-key");
        if (operation === "read") expect(reopened.read()?.source).toEqual(source);
        else reopened.begin(source, "1");
        if (process.platform !== "win32") {
          expect(statSync(dir).mode & 0o777).toBe(0o700);
          expect(statSync(journal.path).mode & 0o777).toBe(0o600);
        }
      } finally {
        process.umask(mask);
        journal.close();
        reopened?.close();
        rmSync(dir, { recursive: true, force: true });
      }
    });
  }

  test.skipIf(process.platform === "win32")("refuses a symlinked journal file", () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const journal = new FileJournal(dir, "target-key");
    const linked = new FileJournal(dir, "linked-key");
    try {
      journal.begin(new Uint8Array([1, 2, 3]), "1");
      journal.close();
      symlinkSync(journal.path, linked.path);
      expect(() => linked.read()).toThrow("journal path is a symlink");
      expect(() => linked.begin(new Uint8Array([4, 5, 6]), "2")).toThrow("journal path is a symlink");
    } finally {
      journal.close();
      linked.close();
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test.skipIf(process.platform === "win32")("a symlinked journal directory is followed and made private", () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const target = join(dir, "target");
    const path = join(dir, "linked");
    try {
      const journal = new FileJournal(target, "key");
      journal.begin(new Uint8Array([1, 2, 3]), "1");
      journal.close();
      chmodSync(target, 0o755);
      symlinkSync(target, path, "dir");
      const linked = new FileJournal(path, "key");
      expect(linked.read()?.source).toEqual(new Uint8Array([1, 2, 3]));
      expect(statSync(target).mode & 0o777).toBe(0o700);
      linked.close();
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test("records the source header, updates and written versions, deletes on request", () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const journal = new FileJournal(dir, "testhost:https://cloud.example.com/wopi/files/file1");
    const expectedName = `${createHash("sha256").update("testhost:https://cloud.example.com/wopi/files/file1").digest("hex")}.journal`;
    expect(journal.path.endsWith(expectedName)).toBe(true);
    const source = new TextEncoder().encode("source-bytes");
    journal.begin(source, "7", 4242, new Uint8Array([9, 9]));
    journal.appendUpdate(new Uint8Array([1, 2, 3]));
    journal.recordWrittenVersion("8");
    journal.appendUpdate(new Uint8Array([4, 5]));
    journal.flush();
    const read = new FileJournal(dir, "testhost:https://cloud.example.com/wopi/files/file1").read();
    expect(read?.itemVersion).toBe("7");
    expect(read?.clientId).toBe(4242);
    expect(read?.initialState).toEqual(new Uint8Array([9, 9]));
    expect(read?.source).toEqual(source);
    expect(read?.sha256).toBe(createHash("sha256").update(source).digest("hex"));
    expect(read?.writtenVersions).toEqual(["8"]);
    expect(read?.updates.map((u) => [...u])).toEqual([[1, 2, 3], [4, 5]]);
    journal.delete();
    expect(existsSync(journal.path)).toBe(false);
    expect(readdirSync(dir).filter((n) => n.includes("secret"))).toEqual([]);
  });

  for (const format of ["docx", "xlsx", "pptx"] as Format[]) {
    test(`edits, autosave, more edits and crash recovery saves every edit for ${format}`, async () => {
      const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
      const h = makeHarness({ journalDir: dir, format });
      addToken(h, "tok-a", "user-a");
      addToken(h, "tok-b", "user-b");
      const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
      const session = [...h.registry.list()][0]!;
      const client = await openTestClient(a.joinPayload, sourceBytes(format), 93001);
      try {
        expect(session.baselineGeneration).toBe(1);
        client.edit("saved-edit ", 0);
        session.applyUpdate(a.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector()));
        await h.clock.advance(30000);
        await h.clock.drain();
        expect(h.putCount()).toBe(1);
        client.edit("post-save-edit ", 1);
        session.applyUpdate(a.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector()));
        session.abandon();

        const registry2 = restartRegistry(h, dir);
        const b = await registry2.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
        const session2 = [...registry2.list()][0]!;
        try {
          expect(session2.baselineGeneration).toBe(2);
          expect(session2.replayStats?.rejected).toBe(0);
          expect(session2.replayStats?.accepted).toBeGreaterThan(0);
          expect(session2.dirty).toBe(true);
          await session2.requestSave(b.user.id);
          const fresh = await openFreshClient(format, h.host.files.get("file1")!.bytes, 93002);
          try {
            expect(fresh.contains("saved-edit")).toBe(true);
            expect(fresh.contains("post-save-edit")).toBe(true);
          } finally {
            fresh.destroy();
          }
        } finally {
          await session2.leave(b.user.id).catch(() => {});
        }
      } finally {
        client.destroy();
      }
    });
  }

  test("an apply time repair is journaled after its client update and replays cleanly", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const h = makeHarness({ journalDir: dir, format: "docx" });
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const source = sourceBytes("docx");
    const peer = await openTestClient(a.joinPayload, source, 95001);
    const author = await openTestClient(a.joinPayload, source, 95011);
    const emitted: Uint8Array[] = [];
    session.on("replicaUpdate", (e) => emitted.push(e.update));
    try {
      peer.splitFirstParagraph();
      session.applyUpdate(a.user.id, peer.encodeStateAsUpdate(session.replica!.encodeStateVector()));
      await session.requestSave(a.user.id);
      for (const update of emitted) peer.applyUpdate(update);
      const savedId = peer.secondParagraphOoxmlId();
      expect(savedId).toBeTruthy();
      peer.deleteSplitSecondHalf();
      session.applyUpdate(a.user.id, peer.encodeStateAsUpdate(session.replica!.encodeStateVector()));

      author.insertPilcrowClaiming(savedId!);
      const authorUpdate = author.encodeStateAsUpdate(session.replica!.encodeStateVector());
      const journalReader = new FileJournal(dir, SessionRegistry.fileKey("testhost", SRC));
      const priorUpdates = journalReader.read()!.updates;
      emitted.length = 0;
      const result = session.applyUpdate(a.user.id, authorUpdate);
      expect(result.accepted).toBe(true);
      if (!result.accepted) throw new Error("author update rejected");
      expect(result.followUps.length).toBeGreaterThan(0);
      expect(emitted.length).toBe(0);

      expect(journalReader.read()!.updates).toEqual([...priorUpdates, authorUpdate, ...result.followUps]);

      session.abandon();
      const registry2 = restartRegistry(h, dir);
      const b = await registry2.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
      const session2 = [...registry2.list()][0]!;
      try {
        expect(session2.replayStats?.rejected).toBe(0);
        expect(session2.dirty).toBe(true);
      } finally {
        await session2.leave(b.user.id).catch(() => {});
      }
    } finally {
      peer.destroy();
      author.destroy();
    }
  });

  test("accepted updates journal returned follow-ups even without local callbacks", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const h = makeHarness({ journalDir: dir, format: "docx" });
    addToken(h, "tok-a", "user-a");
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const session = a.session;
    const replica = session.replica!;
    const applyUpdate = replica.applyUpdate.bind(replica);
    const client = await openTestClient(a.joinPayload, sourceBytes("docx"), 95021);
    const emitted: Uint8Array[] = [];
    session.on("replicaUpdate", (event) => emitted.push(event.update));
    try {
      client.edit("client-update ");
      const update = client.encodeStateAsUpdate(replica.encodeStateVector());
      const afterClient = client.encodeStateVector();
      client.edit("first-follow-up ");
      const first = client.encodeStateAsUpdate(afterClient);
      const afterFirst = client.encodeStateVector();
      client.edit("second-follow-up ");
      const second = client.encodeStateAsUpdate(afterFirst);
      const followUps = [first, second];
      replica.applyUpdate = (incoming) => {
        for (const part of [incoming, ...followUps]) expect(applyUpdate(part).accepted).toBe(true);
        return { accepted: true, followUps };
      };

      expect(session.applyUpdate(a.user.id, update)).toEqual({ accepted: true, followUps });
      expect(emitted).toEqual([]);
      const journal = new FileJournal(dir, SessionRegistry.fileKey("testhost", SRC)).read();
      expect(journal!.updates).toEqual([update, first, second]);
      expect(session.dirty).toBe(true);
    } finally {
      replica.applyUpdate = applyUpdate;
      client.destroy();
      session.abandon();
    }
  });

  for (const crash of [false, true]) {
    test(`rejected follow-ups are journaled and ${crash ? "replayed after a crash" : "autosaved"}`, async () => {
      const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
      const h = makeHarness({ journalDir: dir, format: "docx" });
      addToken(h, "tok-a", "user-a");
      addToken(h, "tok-b", "user-b");
      const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
      const session = a.session;
      const replica = session.replica!;
      const applyUpdate = replica.applyUpdate.bind(replica);
      const client = await openTestClient(a.joinPayload, sourceBytes("docx"), 95031);
      const emitted: Uint8Array[] = [];
      session.on("replicaUpdate", (event) => emitted.push(event.update));
      let recovered: FileSession | null = null;
      try {
        client.edit("integrated-prefix ");
        const prefix = client.encodeStateAsUpdate(replica.encodeStateVector());
        client.edit("rejected-suffix ");
        const update = client.encodeStateAsUpdate(replica.encodeStateVector());
        const rejection = { accepted: false as const, reason: "engine rejected update (Error)", followUps: [prefix] };
        replica.applyUpdate = () => {
          expect(applyUpdate(prefix).accepted).toBe(true);
          return rejection;
        };

        expect(session.applyUpdate(a.user.id, update)).toEqual(rejection);
        replica.applyUpdate = applyUpdate;
        expect(replica.version).toBe(1);
        expect(session.dirty).toBe(true);
        expect(emitted).toEqual([]);
        const journal = new FileJournal(dir, SessionRegistry.fileKey("testhost", SRC)).read();
        expect(journal!.updates).toEqual([prefix]);
        expect(h.putCount()).toBe(0);

        let active = session;
        if (crash) {
          session.abandon();
          const registry = restartRegistry(h, dir);
          const b = await registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
          recovered = b.session;
          active = recovered;
          expect(active.replayStats).toEqual({ accepted: 2, rejected: 0 });
          expect(active.dirty).toBe(true);
        }
        await h.clock.advance(29000);
        expect(h.putCount()).toBe(0);
        await h.clock.advance(1000);
        await h.clock.drain();
        expect(h.putCount()).toBe(1);
        expect(active.dirty).toBe(false);
        const fresh = await openFreshClient("docx", h.host.files.get("file1")!.bytes, 95032);
        try {
          expect(fresh.contains("integrated-prefix")).toBe(true);
          expect(fresh.contains("rejected-suffix")).toBe(false);
        } finally {
          fresh.destroy();
        }
      } finally {
        replica.applyUpdate = applyUpdate;
        client.destroy();
        recovered?.abandon();
        session.abandon();
      }
    });
  }

  test("a viewer first after a crash does not replay or delete the journal", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const h = makeHarness({ journalDir: dir });
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const client = await openTestClient(a.joinPayload, sourceBytes("xlsx"), 94001);
    const journalPath = FileJournal.pathFor(dir, SessionRegistry.fileKey("testhost", SRC));
    try {
      client.edit("crash-held ", 0);
      session.applyUpdate(a.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector()));
      session.abandon();
      expect(existsSync(journalPath)).toBe(true);

      const registry2 = restartRegistry(h, dir);
      const viewer = await registry2.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "view" });
      const viewerSession = [...registry2.list()][0]!;
      expect(viewerSession.dirty).toBe(false);
      await viewerSession.leave(viewer.user.id);
      expect(existsSync(journalPath)).toBe(true);

      const registry3 = restartRegistry(h, dir);
      const b = await registry3.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
      const session3 = [...registry3.list()][0]!;
      try {
        expect(session3.dirty).toBe(true);
        await h.clock.advance(29000);
        expect(h.putCount()).toBe(0);
        await h.clock.advance(2000);
        expect(h.putCount()).toBe(1);
        const fresh = await openFreshClient("xlsx", h.host.files.get("file1")!.bytes, 94002);
        try {
          expect(fresh.contains("crash-held")).toBe(true);
        } finally {
          fresh.destroy();
        }
      } finally {
        await session3.leave(b.user.id).catch(() => {});
      }
    } finally {
      client.destroy();
    }
  });

  test("crash simulation replays into a new registry and the next save writes it", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const h = makeHarness({ journalDir: dir });
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const client = await openTestClient(a.joinPayload, sourceBytes("xlsx"), 91001);
    try {
      client.edit("journaled-edit ", 0);
      const replica = session.replica!;
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted).toBe(true);
      session.abandon();
      expect(h.host.files.get("file1")!.lock).toBeTruthy();

      const h2 = makeHarness({ journalDir: dir });
      h2.host.files.get("file1")!.lock = h.host.files.get("file1")!.lock;
      h2.host.files.get("file1")!.lockExpiresAt = h.clock.t + 30 * 60000;
      addToken(h2, "tok-b", "user-b");
      const b = await h2.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
      const session2 = [...h2.registry.list()][0]!;
      try {
        expect(session2.dirty).toBe(true);
        await session2.requestSave(b.user.id);
        const fresh = await openFreshClient("xlsx", h2.host.files.get("file1")!.bytes, 91002);
        try {
          expect(fresh.contains("journaled-edit")).toBe(true);
        } finally {
          fresh.destroy();
        }
      } finally {
        await session2.leave(b.user.id).catch(() => {});
      }
    } finally {
      client.destroy();
    }
  });

  test("an unmatched host version recovers the journal into a conflict copy before deleting it", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const h = makeHarness({ journalDir: dir });
    addToken(h, "tok-a", "user-a");
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const session = [...h.registry.list()][0]!;
    const client = await openTestClient(a.joinPayload, sourceBytes("xlsx"), 92001);
    try {
      client.edit("stale-edit ", 0);
      const replica = session.replica!;
      session.applyUpdate(a.user.id, client.encodeStateAsUpdate(replica.encodeStateVector()));
      await session.requestSave(a.user.id);
      client.edit("after-checkpoint-edit ", 1);
      session.applyUpdate(a.user.id, client.encodeStateAsUpdate(replica.encodeStateVector()));
      session.abandon();

      const h2 = makeHarness({ journalDir: dir });
      h2.host.files.get("file1")!.lock = h.host.files.get("file1")!.lock;
      h2.host.files.get("file1")!.lockExpiresAt = h.clock.t + 30 * 60000;
      const moved = h2.host.files.get("file1")!;
      moved.version = 9;
      moved.itemVersion = "9";
      addToken(h2, "tok-b", "user-b");
      const b = await h2.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
      const session2 = [...h2.registry.list()][0]!;
      try {
        expect(session2.dirty).toBe(false);
        expect(h2.logs.join("\n")).toContain("journal recovered as a conflict copy");
        expect(h2.putCount()).toBe(0);
        expect(h2.callsFor("PUT_RELATIVE")).toHaveLength(1);
        expect(existsSync(FileJournal.pathFor(dir, SessionRegistry.fileKey("testhost", SRC)))).toBe(false);
        const copy = [...h2.host.files.values()].find((file) => file.id !== "file1")!;
        const recovered = await openFreshClient("xlsx", copy.bytes, 92003);
        try {
          expect(recovered.contains("stale-edit")).toBe(true);
          expect(recovered.contains("after-checkpoint-edit")).toBe(true);
        } finally {
          recovered.destroy();
        }
        const fresh = await openFreshClient("xlsx", h2.host.files.get("file1")!.bytes, 92002);
        try {
          expect(fresh.contains("stale-edit")).toBe(false);
        } finally {
          fresh.destroy();
        }
      } finally {
        await session2.leave(b.user.id).catch(() => {});
      }
    } finally {
      client.destroy();
    }
  });
});

describe("journal durability regressions", () => {
  test("a fresh writer flushes a broken journal from hold without accepting new updates", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const h = makeHarness();
    const expiresAt = h.clock.t + 1000;
    addToken(h, "tok-a", "user-a", { expiresAt });
    addToken(h, "tok-b", "user-b");
    let unavailable = true;
    const fetch = h.host.fetch.bind(h.host);
    h.host.fetch = async (request) => {
      if (unavailable && request.headers.get("X-WOPI-Override") === "PUT") return new Response(null, { status: 500 });
      return fetch(request);
    };
    let failing = false;
    const journal = new FileJournal(dir, SessionRegistry.fileKey("testhost", SRC), {
      write: (fd, buffer, offset, length) => failing ? 0 : writeSync(fd, buffer, offset, length),
    });
    const session = new FileSession({
      hostClient: h.hostClient, hostId: "testhost", wopiSrc: SRC,
      fileKey: SessionRegistry.fileKey("testhost", SRC), instanceId: "journal-return-test",
      journal, logger: h.logger, now: h.clock.now, scheduler: h.scheduler,
    });
    const a = await session.join({ wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: expiresAt, mode: "edit" });
    const client = await openTestClient(a.joinPayload, sourceBytes("xlsx"), 96041);
    const failures: unknown[] = [];
    const writable: unknown[] = [];
    session.on("saveFailed", (event) => failures.push(event));
    session.on("writable", (event) => writable.push(event));
    try {
      failing = true;
      client.edit("unjournaled-held-edit ");
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector())).accepted).toBe(true);
      expect(journal.read()?.updates).toEqual([]);
      await h.clock.drain();
      expect(h.putCount()).toBe(1);
      expect(failures).toEqual([{ kind: "server" }]);
      expect(session.dirty).toBe(true);
      await h.clock.advance(1000);
      expect(a.user.isTokenValidAt(h.clock.t)).toBe(false);
      expect(session.awaitingFlushToken).toBe(true);
      expect(session.state).toBe("open");

      unavailable = false;
      const b = await session.join({ wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
      expect(b.user.canWrite).toBe(true);
      expect(b.file.canWrite).toBe(false);
      expect(b.file.readOnly).toBe(true);
      expect(b.viewOnlyReason).toBe("journalBroken");
      expect(session.journalBroken).toBe(true);
      expect(session.viewOnlyReason).toBe("journalBroken");
      expect(writable).toEqual([]);
      expect(session.awaitingFlushToken).toBe(false);
      expect(session.dirty).toBe(false);
      expect(h.callsFor("REFRESH_LOCK").map((call) => call.token)).toEqual(["tok-b"]);
      expect(h.callsFor("PUT").map((call) => call.token)).toEqual(["tok-a", "tok-b"]);
      const returning = await openTestClient(b.joinPayload, sourceBytes("xlsx"), 96042);
      try {
        returning.edit("blocked-returning-edit ", 1);
        expect(session.applyUpdate(b.user.id, returning.encodeStateAsUpdate(session.replica!.encodeStateVector()))).toEqual({ accepted: false, reason: "user is read only" });
        expect(session.dirty).toBe(false);
      } finally {
        returning.destroy();
      }
      const fresh = await openFreshClient("xlsx", h.host.files.get("file1")!.bytes, 96043);
      try {
        expect(fresh.contains("unjournaled-held-edit")).toBe(true);
        expect(fresh.contains("blocked-returning-edit")).toBe(false);
      } finally {
        fresh.destroy();
      }
      await session.leave(b.user.id);
      expect(h.callsFor("UNLOCK").map((call) => call.token)).toEqual(["tok-b"]);
      expect(session.state).toBe("closed");
    } finally {
      client.destroy();
      await session.close();
    }
  });

  test("a broken journal retains update refusal while a save crossing lease expiry recovers", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const h = makeHarness({ refreshStatus: () => 500 });
    addToken(h, "tok-a", "user-a");
    let failing = false;
    const journal = new FileJournal(dir, SessionRegistry.fileKey("testhost", SRC), {
      write: (fd, buffer, offset, length) => failing ? 0 : writeSync(fd, buffer, offset, length),
    });
    const session = new FileSession({
      hostClient: h.hostClient, hostId: "testhost", wopiSrc: SRC,
      fileKey: SessionRegistry.fileKey("testhost", SRC), instanceId: "journal-lease-test",
      journal, logger: h.logger, now: h.clock.now, scheduler: h.scheduler,
    });
    const a = await session.join({ wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const client = await openTestClient(a.joinPayload, sourceBytes("xlsx"), 96021);
    let releaseSave: () => void = () => {};
    const saveGate = new Promise<void>((resolve) => { releaseSave = resolve; });
    let enteredSave: () => void = () => {};
    const saveEntered = new Promise<void>((resolve) => { enteredSave = resolve; });
    const replica = session.replica!;
    const save = replica.save.bind(replica);
    const writable: unknown[] = [];
    session.on("writable", (event) => writable.push(event));
    try {
      const deadline = h.clock.t + LOCK_LEASE_MS;
      while (h.clock.t < deadline - 1000) {
        session.heartbeat(a.user.id);
        await h.clock.advance(Math.min(30000, deadline - 1000 - h.clock.t));
      }
      replica.save = async () => {
        enteredSave();
        await saveGate;
        return save();
      };
      failing = true;
      client.edit("unjournaled-lease-edit ");
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted).toBe(true);
      await saveEntered;
      expect(journal.read()?.updates).toEqual([]);
      session.heartbeat(a.user.id);
      await h.clock.advance(1001);
      expect(session.lockHeld).toBe(false);
      expect(session.viewOnlyReason).toBe("journalBroken");
      releaseSave();
      await h.clock.drain();
      expect(h.putCount()).toBe(0);
      expect(session.dirty).toBe(true);
      await h.clock.advance(30000);
      expect(session.lockHeld).toBe(true);
      expect(session.dirty).toBe(false);
      expect(session.journalBroken).toBe(true);
      expect(session.viewOnlyReason).toBe("journalBroken");
      expect(writable).toEqual([]);
      client.edit("blocked-after-journal-recovery ", 1);
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted).toBe(false);
      expect(h.putCount()).toBe(1);
      const fresh = await openFreshClient("xlsx", h.host.files.get("file1")!.bytes, 96022);
      try {
        expect(fresh.contains("unjournaled-lease-edit")).toBe(true);
        expect(fresh.contains("blocked-after-journal-recovery")).toBe(false);
      } finally {
        fresh.destroy();
      }
    } finally {
      releaseSave();
      client.destroy();
      await session.close();
    }
  });

  test("opening a journal truncates a torn tail before appending recoverable records", () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const journal = new FileJournal(dir, "torn-tail");
    const source = new Uint8Array([1, 2]);
    journal.begin(source, "1");
    journal.appendUpdate(new Uint8Array([3]));
    journal.close();
    const completeLength = readFileSync(journal.path).length;
    appendFileSync(journal.path, new Uint8Array([0, 0, 0, 8, 85, 4]));
    const reopened = new FileJournal(dir, "torn-tail");
    reopened.begin(source, "1");
    expect(readFileSync(journal.path).length).toBe(completeLength);
    reopened.appendUpdate(new Uint8Array([5, 6]));
    reopened.recordWrittenVersion("2");
    reopened.close();
    const recovered = new FileJournal(dir, "torn-tail").read();
    expect(recovered?.updates).toEqual([new Uint8Array([3]), new Uint8Array([5, 6])]);
    expect(recovered?.writtenVersions).toEqual(["2"]);
  });

  test("short writes are completed for source, update and version records", () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const journal = new FileJournal(dir, "short-writes", {
      write: (fd, buffer, offset, length) => writeSync(fd, buffer, offset, Math.min(2, length)),
    });
    const source = new Uint8Array([1, 2, 3, 4]);
    journal.begin(source, "1", 42, new Uint8Array([9, 8]));
    journal.appendUpdate(new Uint8Array([5, 6, 7]));
    journal.recordWrittenVersion("2");
    journal.close();
    const recovered = journal.read();
    expect(recovered?.source).toEqual(source);
    expect(recovered?.initialState).toEqual(new Uint8Array([9, 8]));
    expect(recovered?.updates).toEqual([new Uint8Array([5, 6, 7])]);
    expect(recovered?.writtenVersions).toEqual(["2"]);
  });

  test("a partial append failure marks the journal broken and promptly saves the accepted edit", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    let failing = false;
    let wrotePartial = false;
    const journal = new FileJournal(dir, SessionRegistry.fileKey("testhost", SRC), {
      write: (fd, buffer, offset, length) => {
        if (!failing) return writeSync(fd, buffer, offset, length);
        if (wrotePartial) return 0;
        wrotePartial = true;
        return writeSync(fd, buffer, offset, Math.min(2, length));
      },
    });
    const session = new FileSession({
      hostClient: h.hostClient, hostId: "testhost", wopiSrc: SRC,
      fileKey: SessionRegistry.fileKey("testhost", SRC), instanceId: "journal-test",
      journal, logger: h.logger, now: h.clock.now, scheduler: h.scheduler,
    });
    const a = await session.join({ wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const client = await openTestClient(a.joinPayload, sourceBytes("xlsx"), 96001);
    const failures: unknown[] = [];
    session.on("journalBroken", (event) => failures.push(event));
    try {
      failing = true;
      client.edit("append-failed ");
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector())).accepted).toBe(true);
      expect(session.journalBroken).toBe(true);
      expect(session.viewOnlyReason).toBe("journalBroken");
      expect(failures).toEqual([{ kind: "writeFailed" }]);
      expect(h.logs.join("\n")).toContain("error journal broken");
      client.edit("must-not-accept ", 1);
      expect(session.applyUpdate(a.user.id, client.encodeStateAsUpdate(session.replica!.encodeStateVector())).accepted).toBe(false);
      const appliedAt = h.clock.t;
      await h.clock.drain();
      expect(h.clock.t).toBe(appliedAt);
      expect(h.putCount()).toBe(1);
      expect(h.callsFor("PUT")[0]?.headers.get("X-COOL-WOPI-IsAutosave")).toBe("true");
      expect(session.dirty).toBe(false);
      const fresh = await openFreshClient("xlsx", h.host.files.get("file1")!.bytes, 96002);
      try {
        expect(fresh.contains("append-failed")).toBe(true);
        expect(fresh.contains("must-not-accept")).toBe(false);
      } finally {
        fresh.destroy();
      }
    } finally {
      client.destroy();
      await session.close();
    }
    const repaired = new FileJournal(dir, SessionRegistry.fileKey("testhost", SRC));
    repaired.begin(sourceBytes("xlsx"), "1");
    expect(repaired.read()?.updates).toEqual([]);
    repaired.close();
  });

  test("an unmatched journal survives a failed conflict copy and recovers on retry", async () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    let failCopy = true;
    const h = makeHarness({ journalDir: dir, relativeStatus: () => failCopy ? 500 : null });
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const client = await openTestClient(a.joinPayload, sourceBytes("xlsx"), 96011);
    const path = FileJournal.pathFor(dir, SessionRegistry.fileKey("testhost", SRC));
    try {
      client.edit("unmatched-preserved ");
      a.session.applyUpdate(a.user.id, client.encodeStateAsUpdate(a.session.replica!.encodeStateVector()));
      a.session.abandon();
      const original = readFileSync(path);
      const file = h.host.files.get("file1")!;
      file.version = 9;
      file.itemVersion = "9";
      const registry = restartRegistry(h, dir);
      await expect(registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" })).rejects.toThrow();
      expect(readFileSync(path)).toEqual(original);
      expect(h.putCount()).toBe(0);
      failCopy = false;
      const b = await registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
      expect(existsSync(path)).toBe(false);
      const copy = [...h.host.files.values()].find((entry) => entry.id !== "file1")!;
      const fresh = await openFreshClient("xlsx", copy.bytes, 96012);
      try {
        expect(fresh.contains("unmatched-preserved")).toBe(true);
      } finally {
        fresh.destroy();
      }
      await b.session.close();
    } finally {
      client.destroy();
    }
  });
});
