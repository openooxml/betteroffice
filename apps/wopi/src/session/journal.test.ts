import { beforeAll, describe, expect, test } from "bun:test";
import { existsSync, mkdtempSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";
import { initEngines } from "../replica/engines";
import { openFreshClient, openTestClient, sourceBytes } from "../../test/replica/clients";
import { addToken, makeHarness, SRC } from "../../test/session/harness";
import { FileJournal } from "./journal";

beforeAll(async () => {
  await initEngines();
});

describe("FileJournal", () => {
  test("records the base version and length prefixed updates, deletes on request", () => {
    const dir = mkdtempSync(join(tmpdir(), "wopi-journal-"));
    const journal = new FileJournal(dir, "testhost:https://cloud.example.com/wopi/files/file1");
    const expectedName = `${createHash("sha256").update("testhost:https://cloud.example.com/wopi/files/file1").digest("hex")}.journal`;
    expect(journal.path.endsWith(expectedName)).toBe(true);
    journal.begin("7");
    journal.appendUpdate(new Uint8Array([1, 2, 3]));
    journal.appendUpdate(new Uint8Array([4, 5]));
    journal.flush();
    const read = new FileJournal(dir, "testhost:https://cloud.example.com/wopi/files/file1").read();
    expect(read?.baseItemVersion).toBe("7");
    expect(read?.updates.map((u) => [...u])).toEqual([[1, 2, 3], [4, 5]]);
    journal.delete();
    expect(existsSync(journal.path)).toBe(false);
    expect(readdirSync(dir).filter((n) => n.includes("secret"))).toEqual([]);
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
