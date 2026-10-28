import { afterEach, beforeAll, describe, expect, test } from "bun:test";
import { mkdtempSync, writeFileSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { generateKeyPair } from "./proof";
import { createHandler, resolveStaticPath, startServer } from "./server";
import { initEngines } from "./replica/engines";
import { openTestClient } from "../test/replica/clients";
import { addToken, makeHarness, SRC } from "../test/session/harness";
import { decodeJoinPayload } from "./sessionApi";

const TOKEN = "SECRETTOKEN-static-3c2b1a";

beforeAll(async () => {
  await initEngines();
});

let servers: Array<{ stop: (close?: boolean) => void }> = [];
afterEach(() => {
  for (const server of servers) server.stop(true);
  servers = [];
});

function roots(): { webRoot: string; fontRoots: string[] } {
  const base = mkdtempSync(join(tmpdir(), "wopi-static-"));
  const webRoot = join(base, "dist");
  const fonts = join(base, "fonts");
  const fontsCjk = join(base, "fonts-cjk");
  mkdirSync(join(webRoot, "assets"), { recursive: true });
  mkdirSync(fonts, { recursive: true });
  mkdirSync(fontsCjk, { recursive: true });
  writeFileSync(join(webRoot, "main.js"), "console.log('wopi');\n");
  writeFileSync(join(webRoot, "assets", "chunk-abc123.js"), "export {};\n");
  writeFileSync(join(webRoot, "assets", "engine.wasm"), new Uint8Array([0, 97, 115, 109]));
  writeFileSync(join(webRoot, "assets", "style.css"), "body{}\n");
  writeFileSync(join(fonts, "Inter-Regular.ttf"), new Uint8Array([0, 1, 2, 3]));
  writeFileSync(join(fontsCjk, "NotoSansSC-Regular.otf"), new Uint8Array([4, 5, 6]));
  writeFileSync(join(base, "secret.txt"), "secret");
  return { webRoot, fontRoots: [fonts, fontsCjk] };
}

function handlerWith(r: { webRoot: string; fontRoots: string[] }) {
  const h = makeHarness();
  return createHandler({ config: h.config, currentKey: generateKeyPair(), registry: h.registry, webRoot: r.webRoot, fontRoots: r.fontRoots });
}

describe("static /app/ and /fonts/ serving", () => {
  test("serves main.js and hashed assets with correct types and caching", async () => {
    const handler = handlerWith(roots());
    const main = await handler(new Request("http://local/app/main.js"));
    expect(main.status).toBe(200);
    expect(main.headers.get("content-type")).toContain("text/javascript");
    expect(main.headers.get("cache-control")).toBe("no-cache");
    expect(await main.text()).toContain("wopi");

    const chunk = await handler(new Request("http://local/app/assets/chunk-abc123.js"));
    expect(chunk.status).toBe(200);
    expect(chunk.headers.get("cache-control")).toBe("public, max-age=31536000, immutable");

    const wasm = await handler(new Request("http://local/app/assets/engine.wasm"));
    expect(wasm.headers.get("content-type")).toBe("application/wasm");

    const css = await handler(new Request("http://local/app/assets/style.css"));
    expect(css.headers.get("content-type")).toContain("text/css");
  });

  test("serves bundled fonts immutably with font content types", async () => {
    const handler = handlerWith(roots());
    const ttf = await handler(new Request("http://local/fonts/Inter-Regular.ttf"));
    expect(ttf.status).toBe(200);
    expect(ttf.headers.get("content-type")).toBe("font/ttf");
    expect(ttf.headers.get("cache-control")).toBe("public, max-age=31536000, immutable");

    const otf = await handler(new Request("http://local/fonts/NotoSansSC-Regular.otf"));
    expect(otf.status).toBe(200);
    expect(otf.headers.get("content-type")).toBe("font/otf");

    expect((await handler(new Request("http://local/fonts/Missing.ttf"))).status).toBe(404);
  });

  test("path traversal cannot escape the served roots", async () => {
    const r = roots();
    const handler = handlerWith(r);
    for (const path of [
      "/app/%2e%2e/secret.txt",
      "/app/assets/%2e%2e%2fsecret.txt",
      "/app/..%5csecret.txt",
      "/fonts/%2e%2e/secret.txt",
      "/fonts/..%2fsecret.txt",
    ]) {
      const response = await handler(new Request(`http://local${path}`));
      expect(response.status).toBe(404);
      expect(await response.text()).not.toContain("secret");
    }
    expect(resolveStaticPath(r.webRoot, "/%2e%2e/secret.txt")).toBeNull();
    expect(resolveStaticPath(r.webRoot, "/assets/chunk-abc123.js")).not.toBeNull();
  });

  test("non GET methods and missing files are not served", async () => {
    const handler = handlerWith(roots());
    expect((await handler(new Request("http://local/app/main.js", { method: "POST" }))).status).toBe(404);
    expect((await handler(new Request("http://local/app/nope.js"))).status).toBe(404);
  });
});

describe("bootstrap page", () => {
  test("page html has the root element and stylesheet the boot code needs", async () => {
    const h = makeHarness();
    addToken(h, TOKEN, "user-a", { name: "Alice A" });
    const server = startServer({ config: h.config, currentKey: generateKeyPair(), registry: h.registry }, 0);
    servers.push(server);
    const base = `http://127.0.0.1:${server.port}`;
    const page = await fetch(`${base}/wopi/edit?WOPISrc=${encodeURIComponent(SRC)}`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ access_token: TOKEN, access_token_ttl: "0" }).toString(),
    });
    expect(page.status).toBe(200);
    const html = await page.text();
    expect(html).toContain('<div id="root"></div>');
    expect(html).toContain('<link rel="stylesheet" href="/app/main.css">');
    expect(html).toContain('<script type="module" src="/app/main.js"></script>');
    expect(html).not.toContain(TOKEN);
  });

  test("main.css is served as css with no-cache", async () => {
    const r = roots();
    const { writeFileSync } = await import("node:fs");
    const { join } = await import("node:path");
    writeFileSync(join(r.webRoot, "main.css"), "body{}\n");
    const handler = handlerWith(r);
    const css = await handler(new Request("http://local/app/main.css"));
    expect(css.status).toBe(200);
    expect(css.headers.get("content-type")).toContain("text/css");
    expect(css.headers.get("cache-control")).toBe("no-cache");
  });
});

describe("editnew flow", () => {
  test("an editnew blank is dirty from the start and the first save writes it without edits", async () => {
    const h = makeHarness();
    h.host.setFile("file1", new Uint8Array(0), { name: "new.xlsx" });
    addToken(h, TOKEN, "user-a", { name: "Alice A" });
    const joined = await h.registry.join({ wopiSrc: SRC, accessToken: TOKEN, accessTokenTtl: 0, mode: "edit", hostId: "testhost" });
    expect(joined.session.dirty).toBe(true);
    const result = await joined.session.requestSave(joined.user.id);
    expect(result.saved).toBe(true);
    expect(h.host.files.get("file1")!.bytes.byteLength).toBeGreaterThan(100);
    expect(joined.session.dirty).toBe(false);
  });

  test("a 0-byte host file opens as a blank, edits, saves back and reopens", async () => {
    const h = makeHarness();
    h.host.setFile("file1", new Uint8Array(0), { name: "new.xlsx" });
    addToken(h, TOKEN, "user-a", { name: "Alice A" });
    const server = startServer({ config: h.config, currentKey: generateKeyPair(), registry: h.registry }, 0);
    servers.push(server);
    const base = `http://127.0.0.1:${server.port}`;

    const page = await fetch(`${base}/wopi/edit?WOPISrc=${encodeURIComponent(SRC)}`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ access_token: TOKEN, access_token_ttl: "0" }).toString(),
    });
    expect(page.status).toBe(200);
    const html = await page.text();
    expect(html).not.toContain(TOKEN);
    const handle = /"sessionHandle":"([^"]+)"/.exec(html)![1]!;
    const auth = { Authorization: `Bearer ${handle}` };

    const source = new Uint8Array(await (await fetch(`${base}/api/session/source`, { headers: auth })).arrayBuffer());
    expect(source.byteLength).toBeGreaterThan(100);

    const payload = decodeJoinPayload((await (await fetch(`${base}/api/session/join`, { headers: auth })).json()) as never);
    expect(payload.format).toBe("xlsx");
    const client = await openTestClient(payload, source, 990301);
    try {
      client.edit("editnew-marker ", 0);
      const session = h.registry.list()[0]!;
      const applied = session.applyUpdate(handle, client.encodeStateAsUpdate(session.replica!.encodeStateVector()));
      expect(applied.accepted).toBe(true);
      const save = await fetch(`${base}/api/session/save`, { method: "POST", headers: auth, body: JSON.stringify({ stateVector: Buffer.from(client.encodeStateVector()).toString("base64") }) });
      expect(save.status).toBe(200);
      await save.text();
    } finally {
      client.destroy();
    }

    const stored = h.host.files.get("file1")!;
    expect(stored.bytes.byteLength).toBeGreaterThan(100);

    const { openFreshClient } = await import("../test/replica/clients");
    const fresh = await openFreshClient("xlsx", stored.bytes, 990303);
    try {
      expect(fresh.contains("editnew-marker")).toBe(true);
    } finally {
      fresh.destroy();
    }
  });
});
