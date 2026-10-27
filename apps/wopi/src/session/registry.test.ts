import { beforeAll, describe, expect, test } from "bun:test";
import { initEngines } from "../replica/engines";
import { addToken, makeHarness, SRC } from "../../test/session/harness";
import { SessionRegistry } from "./registry";

beforeAll(async () => {
  await initEngines();
});

describe("SessionRegistry", () => {
  test("file key strips query and fragment and one session serves both spellings", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    expect(SessionRegistry.fileKey("testhost", `${SRC}?access_token=tok-a&x=1#frag`)).toBe(SessionRegistry.fileKey("testhost", SRC));
    expect(SessionRegistry.fileKey("testhost", SRC)).not.toContain("tok-a");
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: `${SRC}?foo=bar`, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const b = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
    expect(h.registry.list().length).toBe(1);
    expect(a.session).toBe(b.session);
    expect(h.registry.instanceId).toBe("inst-test");
    await a.session.leave(a.user.id);
    await a.session.leave(b.user.id);
    expect(h.registry.list().length).toBe(0);
  });

  test("rejects a wopiSrc outside the allowlist before any host call", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    await expect(h.registry.join({ hostId: "testhost", wopiSrc: "https://evil.example.com/wopi/files/file1", accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" })).rejects.toThrow();
    expect(h.calls.length).toBe(0);
  });

  test("random instance id is generated when config has none", () => {
    const h = makeHarness();
    const other = new SessionRegistry({
      config: { publicUrl: "https://office.example.com", proofKeyFile: "x", maxFileBytes: 1024, hosts: [{ id: "testhost", wopiSrcOrigins: ["https://cloud.example.com"], frameOrigins: [] }] },
      hostClient: h.hostClient,
    });
    expect(typeof other.instanceId).toBe("string");
    expect(other.instanceId.length).toBeGreaterThan(0);
  });
});
