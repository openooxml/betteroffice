import { beforeAll, describe, expect, test } from "bun:test";
import { existsSync, mkdtempSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";
import { initEngines } from "../replica/engines";
import { openFreshClient, openTestClient, sourceBytes } from "../../test/replica/clients";
import { addToken, makeHarness, SRC, type Harness } from "../../test/session/harness";
import { FileJournal } from "./journal";
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

  test("host version changed discards the journal and logs it", async () => {
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
        expect(h2.logs.join("\n")).toContain("journal discarded");
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
