import { afterEach, beforeAll, describe, expect, test } from "bun:test";

import { generateKeyPair } from "./proof";
import { startServer } from "./server";
import { initEngines } from "./replica/engines";
import type { Format } from "./replica/replica";
import { openFreshClient, openTestClient, sourceBytes } from "../test/replica/clients";
import { addToken, makeHarness, SRC, type Harness } from "../test/session/harness";
import { decodeJoinPayload } from "./sessionApi";
import { SessionRegistry } from "./session/registry";
import { CollabRelay } from "./relay/relay";
import { encodeSequencedFrame } from "./relay/sequence";
import { encodeUpdateFrame } from "./relay/frames";

const TOKEN = "SECRETTOKEN-routes-9f8e7d6c";
const TOKEN_B = "SECRETTOKEN-routes-b-1a2b3c";

beforeAll(async () => {
  await initEngines();
});

let servers: Array<{ stop: (close?: boolean) => void }> = [];
afterEach(() => {
  for (const server of servers) server.stop(true);
  servers = [];
});

interface Setup {
  h: Harness;
  base: string;
  bodies: string[];
  headersSeen: string[];
  relay: CollabRelay;
}

function setup(options: { fileName?: string; format?: Format; saveSyncTimeoutMs?: number; quirks?: { timestampConflict?: boolean } } = {}): Setup {
  const h = makeHarness({ fileName: options.fileName, format: options.format, quirks: options.quirks });
  if (options.fileName) h.host.setFile("file1", sourceBytes(options.format ?? "xlsx"), { name: options.fileName });
  addToken(h, TOKEN, "user-a", { name: "Alice A" });
  addToken(h, TOKEN_B, "user-b", { name: "Bob B", canWrite: true });
  const relay = new CollabRelay(h.registry);
  const server = startServer(
    { config: h.config, currentKey: generateKeyPair(), registry: h.registry, relay, saveSyncTimeoutMs: options.saveSyncTimeoutMs },
    0,
  );
  servers.push(server);
  return { h, base: `http://127.0.0.1:${server.port}`, bodies: [], headersSeen: [], relay };
}

async function tracked(setup: Setup, response: Response): Promise<string> {
  const text = await response.text();
  setup.bodies.push(text);
  setup.headersSeen.push(JSON.stringify([...response.headers.entries()]));
  return text;
}

function form(token: string, ttl = 0): string {
  return new URLSearchParams({ access_token: token, access_token_ttl: String(ttl) }).toString();
}

async function postPage(
  s: Setup,
  path: "edit" | "view",
  token = TOKEN,
  wopiSrc: string = SRC,
  extra = "",
): Promise<Response> {
  return fetch(`${s.base}/wopi/${path}?WOPISrc=${encodeURIComponent(wopiSrc)}${extra}`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: form(token),
  });
}

interface PageConfig {
  sessionHandle: string;
  fileName: string;
  format: string;
  canWrite: boolean;
  viewOnlyReason: string | null;
  postMessageOrigin: string;
  closeUrl: string;
  closePostMessage: boolean;
  locale: string | null;
  userFriendlyName: string;
  tokenExpiresAt: number | null;
}

function auth(config: PageConfig): Record<string, string> {
  return { Authorization: `Bearer ${config.sessionHandle}` };
}

function pageConfig(html: string): PageConfig {
  const match = html.match(/<script type="application\/json" id="wopi-config">([\s\S]*?)<\/script>/);
  if (!match) throw new Error("missing wopi-config script");
  return JSON.parse(match[1]!) as PageConfig;
}

function expectNoToken(s: Setup, extra: string[] = []): void {
  const all = [...s.bodies, ...s.headersSeen, ...s.h.logs, ...extra].join("\n");
  expect(all).not.toContain(TOKEN);
  expect(all).not.toContain(TOKEN_B);
}

describe("wopi edit and view routes", () => {
  test("edit POST returns bootstrap page with CSP and no token", async () => {
    const s = setup();
    const response = await postPage(s, "edit", TOKEN, SRC, "&ui=en-US&rs=en&dchat=1");
    expect(response.status).toBe(200);
    expect(response.headers.get("cache-control")).toBe("no-store");
    expect(response.headers.get("referrer-policy")).toBe("no-referrer");
    expect(response.headers.get("x-frame-options")).toBeNull();
    const csp = response.headers.get("content-security-policy") ?? "";
    expect(csp).toContain("frame-ancestors https://cloud.example.com");
    expect(csp).toContain("default-src 'self'");
    expect(csp).toContain("connect-src 'self'");
    expect(csp).toContain("img-src 'self' blob: data:");
    expect(csp).toContain("font-src 'self' data:");
    expect(csp).toContain("style-src 'self' 'unsafe-inline'");
    expect(csp).toContain("worker-src 'self' blob:");
    expect(csp).toContain("script-src 'self' 'wasm-unsafe-eval'");
    const html = await tracked(s, response);
    expect(html).toContain('<script type="module" src="/app/main.js">');
    const config = pageConfig(html);
    expect(config.sessionHandle.length).toBeGreaterThan(20);
    expect(config.fileName).toBe("book.xlsx");
    expect(config.format).toBe("xlsx");
    expect(config.canWrite).toBe(true);
    expect(config.viewOnlyReason).toBeNull();
    expect(config.postMessageOrigin).toBe("https://cloud.example.com");
    expect(config.closeUrl).toBe("");
    expect(config.locale).toBe("en-US");
    expect(config.userFriendlyName).toBe("Alice A");
    expect(config.tokenExpiresAt).toBeNull();
    expectNoToken(s);
  });

  test("view POST joins read only", async () => {
    const s = setup();
    const response = await postPage(s, "view");
    expect(response.status).toBe(200);
    const config = pageConfig(await tracked(s, response));
    expect(config.canWrite).toBe(false);
    expectNoToken(s);
  });

  test("disallowed WOPISrc is 400 before any outbound call", async () => {
    const s = setup();
    const response = await postPage(s, "edit", TOKEN, "https://evil.example.com/wopi/files/file1");
    expect(response.status).toBe(400);
    await tracked(s, response);
    expect(s.h.calls.length).toBe(0);
    expectNoToken(s);
  });

  test("missing WOPISrc and missing token are 400, GET is 405", async () => {
    const s = setup();
    expect((await fetch(`${s.base}/wopi/edit`, { method: "POST", body: form(TOKEN) })).status).toBe(400);
    const noToken = await fetch(`${s.base}/wopi/edit?WOPISrc=${encodeURIComponent(SRC)}`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: "access_token_ttl=0",
    });
    expect(noToken.status).toBe(400);
    await tracked(s, noToken);
    for (const path of ["edit", "view"]) {
      const get = await fetch(`${s.base}/wopi/${path}?WOPISrc=${encodeURIComponent(SRC)}`);
      expect(get.status).toBe(405);
      expect(get.headers.get("allow")).toContain("POST");
      await tracked(s, get);
    }
    expectNoToken(s);
  });

  test("hostile filename cannot break out of the config script", async () => {
    const hostile = "evil</script><script>alert(1)</script>.xlsx";
    const s = setup({ fileName: hostile });
    const response = await postPage(s, "edit");
    expect(response.status).toBe(200);
    const html = await tracked(s, response);
    expect(html).not.toContain("evil</script>");
    expect(html).not.toContain("<script>alert(1)</script>");
    expect(pageConfig(html).fileName).toBe(hostile);
    expectNoToken(s);
  });

  test("postMessageOrigin, closeUrl and ClosePostMessage are kept per user", async () => {
    const h = makeHarness();
    addToken(h, TOKEN, "user-a", { name: "Alice A" });
    addToken(h, TOKEN_B, "user-b", { name: "Bob B" });
    const hostClient = {
      ...h.hostClient,
      checkFileInfo: async (input: { wopiSrc: string; accessToken: string }) => {
        const info = await h.hostClient.checkFileInfo(input);
        const who = input.accessToken === TOKEN ? "alice" : "bob";
        return { ...info, ClosePostMessage: who === "alice", CloseUrl: `https://host.example/${who}`, PostMessageOrigin: `https://${who}.example.com` };
      },
    };
    const registry = new SessionRegistry({ config: h.config, hostClient: hostClient as never, logger: h.logger, now: h.clock.now, scheduler: h.scheduler });
    const server = startServer({ config: h.config, currentKey: generateKeyPair(), registry }, 0);
    servers.push(server);
    const base = `http://127.0.0.1:${server.port}`;
    const post = async (token: string): Promise<PageConfig> => {
      const response = await fetch(`${base}/wopi/edit?WOPISrc=${encodeURIComponent(SRC)}`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: form(token),
      });
      expect(response.status).toBe(200);
      return pageConfig(await response.text());
    };
    const a = await post(TOKEN);
    const b = await post(TOKEN_B);
    expect(a.closePostMessage).toBe(true);
    expect(b.closePostMessage).toBe(false);
    expect(a.closeUrl).toBe("https://host.example/alice");
    expect(a.postMessageOrigin).toBe("https://alice.example.com");
    expect(b.closeUrl).toBe("https://host.example/bob");
    expect(b.postMessageOrigin).toBe("https://bob.example.com");
    const session = registry.list()[0]!;
    expect(session.getUser(a.sessionHandle)?.closePostMessage).toBe(true);
    expect(session.getUser(b.sessionHandle)?.closePostMessage).toBe(false);
    expect(session.getUser(a.sessionHandle)?.closeUrl).toBe("https://host.example/alice");
    expect(session.getUser(a.sessionHandle)?.postMessageOrigin).toBe("https://alice.example.com");
    expect(session.getUser(b.sessionHandle)?.closeUrl).toBe("https://host.example/bob");
  });

  test("form body over 64 KiB is 413", async () => {
    const s = setup();
    const response = await fetch(`${s.base}/wopi/edit?WOPISrc=${encodeURIComponent(SRC)}`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: `access_token=${TOKEN}&pad=${"x".repeat(70 * 1024)}`,
    });
    expect(response.status).toBe(413);
    await tracked(s, response);
    expectNoToken(s);
  });
});

describe("session API", () => {
  async function joined(s: Setup): Promise<PageConfig> {
    const response = await postPage(s, "edit");
    expect(response.status).toBe(200);
    return pageConfig(await tracked(s, response));
  }

  test("auth is required and the right handle gets source and join", async () => {
    const s = setup();
    const config = await joined(s);
    expect((await fetch(`${s.base}/api/session/source`)).status).toBe(401);
    const wrong = await fetch(`${s.base}/api/session/source`, {
      headers: { Authorization: "Bearer wrong-handle-value" },
    });
    expect(wrong.status).toBe(401);
    await tracked(s, wrong);

    const source = await fetch(`${s.base}/api/session/source`, { headers: auth(config) });
    expect(source.status).toBe(200);
    expect(source.headers.get("content-type")).toContain("spreadsheetml");
    const bytes = new Uint8Array(await source.arrayBuffer());
    expect([...bytes]).toEqual([...sourceBytes("xlsx")]);

    const join = await fetch(`${s.base}/api/session/join`, { headers: auth(config) });
    expect(join.status).toBe(200);
    const payload = decodeJoinPayload((await join.json()) as never);
    const client = await openTestClient(payload, bytes, 990001);
    try {
      expect(client.format).toBe("xlsx");
    } finally {
      client.destroy();
    }
    expectNoToken(s);
  });

  test("save returns itemVersion and close ends the handle", async () => {
    const s = setup();
    const config = await joined(s);
    const session = s.h.registry.list()[0]!;
    const joinResponse = await fetch(`${s.base}/api/session/join`, { headers: auth(config) });
    const payload = decodeJoinPayload((await joinResponse.json()) as never);
    const source = new Uint8Array(
      await (await fetch(`${s.base}/api/session/source`, { headers: auth(config) })).arrayBuffer(),
    );
    const client = await openTestClient(payload, source, 990101);
    try {
      client.edit("api-save-marker ", 0);
      const replica = session.replica!;
      const result = session.applyUpdate(
        config.sessionHandle,
        client.encodeStateAsUpdate(replica.encodeStateVector()),
      );
      expect(result.accepted).toBe(true);
      const save = await fetch(`${s.base}/api/session/save`, { method: "POST", headers: auth(config), body: JSON.stringify({ stateVector: Buffer.from(client.encodeStateVector()).toString("base64") }) });
      expect(save.status).toBe(200);
      const saved = (await save.json()) as { itemVersion: string | null };
      s.bodies.push(JSON.stringify(saved));
      expect(saved.itemVersion).toBe("2");
      expect(s.h.putCount()).toBe(1);
    } finally {
      client.destroy();
    }
    const big = await fetch(`${s.base}/api/session/save`, {
      method: "POST",
      headers: { ...auth(config), "Content-Type": "application/json" },
      body: "x".repeat(70 * 1024),
    });
    expect(big.status).toBe(413);
    await tracked(s, big);

    const close = await fetch(`${s.base}/api/session/close`, { method: "POST", headers: auth(config) });
    expect(close.status).toBe(200);
    await tracked(s, close);
    expect((await fetch(`${s.base}/api/session/source`, { headers: auth(config) })).status).toBe(401);
    expectNoToken(s);
  });

  test("SSE delivers saved and conflict without tokens", async () => {
    const s = setup({ quirks: { timestampConflict: true } });
    const config = await joined(s);
    const controller = new AbortController();
    const eventsResponse = await fetch(`${s.base}/api/session/events`, {
      headers: auth(config),
      signal: controller.signal,
    });
    expect(eventsResponse.status).toBe(200);
    expect(eventsResponse.headers.get("content-type")).toContain("text/event-stream");
    const reader = eventsResponse.body!.getReader();
    const decoder = new TextDecoder();
    let text = "";
    const readUntil = async (needle: string): Promise<void> => {
      const deadline = Date.now() + 5000;
      while (!text.includes(needle)) {
        if (Date.now() > deadline) throw new Error(`SSE timeout waiting for ${needle}, got: ${text}`);
        const chunk = await reader.read();
        if (chunk.done) break;
        text += decoder.decode(chunk.value, { stream: true });
      }
      expect(text).toContain(needle);
    };

    const session = s.h.registry.list()[0]!;
    const payload = decodeJoinPayload(
      (await (await fetch(`${s.base}/api/session/join`, { headers: auth(config) })).json()) as never,
    );
    const source = new Uint8Array(
      await (await fetch(`${s.base}/api/session/source`, { headers: auth(config) })).arrayBuffer(),
    );
    const client = await openTestClient(payload, source, 990201);
    try {
      client.edit("sse-one ", 0);
      session.applyUpdate(
        config.sessionHandle,
        client.encodeStateAsUpdate(session.replica!.encodeStateVector()),
      );
      const save = await fetch(`${s.base}/api/session/save`, { method: "POST", headers: auth(config), body: JSON.stringify({ stateVector: Buffer.from(client.encodeStateVector()).toString("base64") }) });
      expect(save.status).toBe(200);
      await save.text();
      await readUntil("event: saved");

      s.h.host.setFile("file1", sourceBytes("xlsx"), {
        name: "book.xlsx",
        lastModifiedTime: new Date(s.h.clock.t + 60000).toISOString(),
      });
      client.edit("sse-two ", 1);
      session.applyUpdate(
        config.sessionHandle,
        client.encodeStateAsUpdate(session.replica!.encodeStateVector()),
      );
      const save2 = await fetch(`${s.base}/api/session/save`, { method: "POST", headers: auth(config), body: JSON.stringify({ stateVector: Buffer.from(client.encodeStateVector()).toString("base64") }) });
      expect(save2.status).toBe(200);
      await save2.text();
      await readUntil("event: conflict");
    } finally {
      client.destroy();
      controller.abort();
      await reader.cancel().catch(() => {});
    }
    expect(text).not.toContain(TOKEN);
    expectNoToken(s, [text]);
  });
});

test("ordered save waits for the requesting client's edits before PutFile", async () => {
  const s = setup();
  const config = pageConfig(await (await postPage(s, "edit")).text());
  const session = s.h.registry.list()[0]!;
  const client = await openTestClient(session.replica!.joinPayload(), session.sourceBytes!, 995001);
  try {
    client.edit("ordered-save-marker", 0);
    let completed = false;
    const saving = fetch(`${s.base}/api/session/save`, {
      method: "POST", headers: { Authorization: `Bearer ${config.sessionHandle}`, "Content-Type": "application/json" },
      body: JSON.stringify({ stateVector: Buffer.from(client.encodeStateVector()).toString("base64") }),
    }).then((response) => { completed = true; return response; });
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(completed).toBe(false);
    expect(s.h.putCount()).toBe(0);
    expect(session.applyUpdate(config.sessionHandle, client.encodeStateAsUpdate(session.replica!.encodeStateVector())).accepted).toBe(true);
    expect((await saving).status).toBe(200);
    expect(s.h.putCount()).toBe(1);
  } finally {
    client.destroy();
  }
});

test("ordered save returns 409 notSynced without writing when updates never arrive", async () => {
  const s = setup({ saveSyncTimeoutMs: 1 });
  const config = pageConfig(await (await postPage(s, "edit")).text());
  const session = s.h.registry.list()[0]!;
  const client = await openTestClient(session.replica!.joinPayload(), session.sourceBytes!, 995002);
  try {
    client.edit("missing-update", 0);
    const response = await fetch(`${s.base}/api/session/save`, {
      method: "POST", headers: { Authorization: `Bearer ${config.sessionHandle}`, "Content-Type": "application/json" },
      body: JSON.stringify({ stateVector: Buffer.from(client.encodeStateVector()).toString("base64") }),
    });
    expect(response.status).toBe(409);
    expect(await response.json()).toEqual({ error: { kind: "notSynced" } });
    expect(s.h.putCount()).toBe(0);
  } finally {
    client.destroy();
  }
});

test("save rejects a missing or malformed state vector", async () => {
  const s = setup();
  const config = pageConfig(await (await postPage(s, "edit")).text());
  for (const body of ["", "{}", '{"stateVector":"!bad"}', '{"stateVector":"/w=="}']) {
    const response = await fetch(`${s.base}/api/session/save`, {
      method: "POST", headers: { Authorization: `Bearer ${config.sessionHandle}` }, body,
    });
    expect(response.status).toBe(400);
    expect(await response.json()).toEqual({ error: { kind: "invalidStateVector" } });
  }
  expect(s.h.putCount()).toBe(0);
});

test.each(["deletion", "insertion"] as const)("HTTP save waits for a delayed %s frame through the socket barrier", async (kind) => {
  const s = setup({ format: "docx" });
  const config = pageConfig(await (await postPage(s, "edit")).text());
  const session = s.h.registry.list()[0]!;
  const user = session.getUser(config.sessionHandle)!;
  const client = await openTestClient(session.replica!.joinPayload(), session.sourceBytes!, 995100);
  const socket = { send() {}, close() {} };
  const connectionId = crypto.randomUUID();
  let seq = 0;
  s.relay.attach(socket, { session, user });
  s.relay.handleMessage(socket, encodeSequencedFrame({ connectionId, seq }));
  try {
    const marker = "socket-save-marker ";
    client.edit(marker, 0);
    if (kind === "deletion") {
      seq += 1;
      s.relay.handleMessage(socket, encodeSequencedFrame({ connectionId, seq }, encodeUpdateFrame(client.encodeStateAsUpdate())));
      await session.requestSave(user.id);
      const vector = client.encodeStateVector();
      client.remove(marker);
      expect(client.encodeStateVector()).toEqual(vector);
    }
    const before = s.h.putCount();
    seq += 1;
    let completed = false;
    const saving = fetch(`${s.base}/api/session/save`, {
      method: "POST", headers: auth(config), body: JSON.stringify({ connectionId, seq }),
    }).then((response) => { completed = true; return response; });
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(completed).toBe(false);
    expect(s.h.putCount()).toBe(before);
    s.relay.handleMessage(socket, encodeSequencedFrame({ connectionId, seq }, encodeUpdateFrame(client.encodeStateAsUpdate())));
    const response = await saving;
    expect(response.status).toBe(200);
    await response.text();
    expect(s.h.putCount()).toBe(before + 1);
    const saved = await openFreshClient("docx", s.h.host.files.get("file1")!.bytes, 995101);
    try {
      expect(saved.contains(marker.trim())).toBe(kind === "insertion");
    } finally {
      saved.destroy();
    }
  } finally {
    client.destroy();
  }
});

test("HTTP socket save times out without writing when its sequence never arrives", async () => {
  const s = setup({ saveSyncTimeoutMs: 1 });
  const config = pageConfig(await (await postPage(s, "edit")).text());
  const response = await fetch(`${s.base}/api/session/save`, {
    method: "POST", headers: auth(config), body: JSON.stringify({ connectionId: crypto.randomUUID(), seq: 1 }),
  });
  expect(response.status).toBe(409);
  expect(await response.json()).toEqual({ error: { kind: "notSynced" } });
  expect(s.h.putCount()).toBe(0);
});

test("HTTP save rejects malformed socket barriers", async () => {
  const s = setup();
  const config = pageConfig(await (await postPage(s, "edit")).text());
  for (const barrier of [{ connectionId: "bad", seq: 1 }, { connectionId: crypto.randomUUID(), seq: -1 }, { seq: 1, stateVector: "AA==" }]) {
    const response = await fetch(`${s.base}/api/session/save`, {
      method: "POST", headers: auth(config), body: JSON.stringify(barrier),
    });
    expect(response.status).toBe(400);
    expect(await response.json()).toEqual({ error: { kind: "invalidSaveBarrier" } });
  }
  expect(s.h.putCount()).toBe(0);
});
