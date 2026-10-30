import { beforeAll, describe, expect, test } from "bun:test";
import { initEngines } from "../replica/engines";
import { addToken, makeHarness, SRC } from "../../test/session/harness";
import { SessionRegistry } from "./registry";

beforeAll(async () => {
  await initEngines();
});

describe("SessionRegistry", () => {
  test("file key strips credentials and fragment and preserves routing parameters", async () => {
    const h = makeHarness();
    addToken(h, "tok-a", "user-a");
    addToken(h, "tok-b", "user-b");
    const src = `${SRC}?tenant=A`;
    const key = SessionRegistry.fileKey("testhost", src);
    expect(SessionRegistry.fileKey("testhost", `${src}&access_token=tok-a&access_token_ttl=123#frag`)).toBe(key);
    expect(SessionRegistry.fileKey("testhost", `${SRC}?access_token=tok-a#frag`)).toBe(SessionRegistry.fileKey("testhost", SRC));
    expect(SessionRegistry.fileKey("testhost", `${SRC}?tenant=B`)).not.toBe(key);
    expect(key).not.toContain("tok-a");
    const session = h.registry.create("testhost", `${src}&access_token=tok-a&access_token_ttl=123#frag`);
    expect(h.registry.create("testhost", `${src}#other`)).toBe(session);
    expect(h.registry.create("testhost", src)).toBe(session);
    expect(session.wopiSrc).toBe(src);
    const checks: URL[] = [];
    const fetch = h.host.fetch.bind(h.host);
    h.host.fetch = async (request) => {
      const url = new URL(request.url);
      if (request.method === "GET" && !url.pathname.endsWith("/contents")) checks.push(url);
      return fetch(request);
    };
    const a = await h.registry.join({ hostId: "testhost", wopiSrc: `${src}&access_token=tok-a`, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
    const b = await h.registry.join({ hostId: "testhost", wopiSrc: `${src}&access_token=tok-b&access_token_ttl=456`, accessToken: "tok-b", accessTokenTtl: 0, mode: "edit" });
    const c = await h.registry.join({ hostId: "testhost", wopiSrc: `${SRC}?tenant=B`, accessToken: "tok-b", accessTokenTtl: 0, mode: "view" });
    try {
      expect(h.registry.list().length).toBe(2);
      expect(a.session).toBe(session);
      expect(a.session).toBe(b.session);
      expect(c.session).not.toBe(session);
      expect(checks.map((url) => url.searchParams.get("tenant"))).toEqual(["A", "A", "A", "B"]);
      expect(checks.every((url) => !url.searchParams.has("access_token_ttl"))).toBe(true);
      expect(h.registry.instanceId).toBe("inst-test");
      await a.session.leave(a.user.id);
      await b.session.leave(b.user.id);
      await c.session.leave(c.user.id);
      expect(h.registry.list().length).toBe(0);
    } finally {
      await h.registry.closeAll();
    }
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
