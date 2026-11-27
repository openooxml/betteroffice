import { beforeAll, expect, test } from "bun:test";
import { initEngines } from "../replica/engines";
import { addToken, makeHarness, SRC } from "../../test/session/harness";

beforeAll(async () => {
  await initEngines();
}, 30_000);

test("an ordinary GetFile opens and nothing is written back", async () => {
  const h = makeHarness({ format: "xlsx" });
  addToken(h, "tok-a", "user-a");
  const original = new Uint8Array(h.host.files.get("file1")!.bytes);
  await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
  await h.clock.advance(120000);
  expect(h.putCount()).toBe(0);
  expect(Buffer.from(h.host.files.get("file1")!.bytes).equals(Buffer.from(original))).toBe(true);
});

for (const format of ["xlsx", "docx", "pptx"] as const) {
  test(`${format}: empty GetFile body for a non-empty file errors and never PutFiles`, async () => {
    const h = makeHarness({ format });
    addToken(h, "tok-a", "user-a");
    const original = new Uint8Array(h.host.files.get("file1")!.bytes);
    expect(original.byteLength).toBeGreaterThan(0);
    const real = h.host.fetch.bind(h.host);
    h.host.fetch = async (request: Request) => {
      const url = new URL(request.url);
      if (request.method === "GET" && url.pathname.endsWith("/contents")) {
        return new Response(new Uint8Array(0), { status: 200, headers: { "Content-Length": "0" } });
      }
      return real(request);
    };
    let opened = false;
    try {
      await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
      opened = true;
    } catch {
      opened = false;
    }
    await h.clock.advance(120000);
    const after = h.host.files.get("file1")!.bytes;
    expect({ opened, puts: h.putCount(), sameBytes: Buffer.from(after).equals(Buffer.from(original)) })
      .toEqual({ opened: false, puts: 0, sameBytes: true });
  });
}
