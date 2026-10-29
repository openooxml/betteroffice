import { afterEach, beforeAll, expect, test } from "bun:test";
import { initEngines } from "./replica/engines";
import { generateKeyPair } from "./proof";
import { startServer } from "./server";
import { CollabRelay } from "./relay/relay";
import { encodeSequencedFrame } from "./relay/sequence";
import { encodeUpdateFrame } from "./relay/frames";
import { SessionRegistry } from "./session/registry";
import { addToken, makeHarness, SRC } from "../test/session/harness";
import { openFreshClient, openTestClient } from "../test/replica/clients";
import { closeDocument } from "../web/src/documentActions";

beforeAll(initEngines);
const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => { for (const dispose of cleanup.splice(0).reverse()) await dispose(); });

async function until(predicate: () => boolean) {
  const deadline = Date.now() + 2000;
  while (!predicate()) {
    if (Date.now() >= deadline) throw new Error("condition timed out");
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

async function setup(beforeUnlock?: () => Promise<void>) {
  const h = makeHarness({ beforeUnlock });
  h.config.reconnectGraceMs = 30_000;
  addToken(h, "recovery-token", "writer");
  const registry = new SessionRegistry({ config: h.config, hostClient: h.hostClient, now: h.clock.now, scheduler: h.scheduler, autosave: false });
  const relay = new CollabRelay(registry);
  const server = startServer({ config: h.config, currentKey: generateKeyPair(), registry, relay }, 0);
  cleanup.push(async () => { server.stop(true); await registry.closeAll(); });
  const base = `http://127.0.0.1:${server.port}`;
  const joined = await registry.join({ wopiSrc: SRC, accessToken: "recovery-token", accessTokenTtl: 0, mode: "edit" });
  const { session, user } = joined;
  const client = await openTestClient(joined.joinPayload, session.sourceBytes!, 998001);
  cleanup.push(async () => client.destroy());
  const headers = { Authorization: `Bearer ${user.id}` };
  async function connect(generation = session.baselineGeneration) {
    const ws = new WebSocket(`${base.replace("http:", "ws:")}/collab`, ["bo-wopi.v1", user.id, `bo-wopi.generation.${generation}`]);
    await new Promise<void>((resolve, reject) => {
      ws.addEventListener("open", () => resolve(), { once: true });
      ws.addEventListener("error", () => reject(new Error("socket failed")), { once: true });
    });
    cleanup.push(async () => {
      if (ws.readyState === WebSocket.CLOSED) return;
      await new Promise<void>((resolve) => { ws.addEventListener("close", () => resolve(), { once: true }); ws.close(); });
    });
    const connectionId = crypto.randomUUID();
    ws.send(encodeSequencedFrame({ connectionId, seq: 0 }) as unknown as ArrayBuffer);
    await until(() => relay.appliedSequence(session, user.id, connectionId) === 0);
    return { ws, connectionId };
  }
  async function drop(ws: WebSocket) {
    await new Promise<void>((resolve) => { ws.addEventListener("close", () => resolve(), { once: true }); ws.close(1001); });
    await until(() => !relay.hasSocket(session, user.id));
    await h.clock.flush();
  }
  return { h, registry, session, user, client, headers, base, relay, connect, drop };
}

test("a lone writer resumes the same handle during grace and keeps saving without departure", async () => {
  const s = await setup();
  const first = await s.connect();
  await s.drop(first.ws);
  await s.h.clock.advance(29_999);
  expect(s.session.state).toBe("open");
  expect(s.h.callsFor("UNLOCK")).toHaveLength(0);
  const second = await s.connect();
  s.client.edit("after-reconnect ");
  second.ws.send(encodeSequencedFrame({ connectionId: second.connectionId, seq: 1 }, encodeUpdateFrame(s.client.encodeStateAsUpdate())) as unknown as ArrayBuffer);
  const saved = await fetch(`${s.base}/api/session/save`, {
    method: "POST", headers: s.headers, body: JSON.stringify({ connectionId: second.connectionId, seq: 1 }),
  });
  expect(saved.status).toBe(200);
  await saved.text();
  await s.h.clock.advance(2);
  expect(s.session.getUser(s.user.id)).toBe(s.user);
  expect(s.h.callsFor("UNLOCK")).toHaveLength(0);
  expect(s.h.callsFor("LOCK")).toHaveLength(1);
  const fresh = await openFreshClient("xlsx", s.h.host.files.get("file1")!.bytes, 998002);
  try { expect(fresh.contains("after-reconnect")).toBe(true); } finally { fresh.destroy(); }
});

test("grace expiry performs one exit flush and one unlock", async () => {
  const s = await setup();
  const { ws, connectionId } = await s.connect();
  s.client.edit("departure-edit ");
  ws.send(encodeSequencedFrame({ connectionId, seq: 1 }, encodeUpdateFrame(s.client.encodeStateAsUpdate())) as unknown as ArrayBuffer);
  await until(() => s.relay.appliedSequence(s.session, s.user.id, connectionId) === 1);
  await s.drop(ws);
  await s.h.clock.advance(29_999);
  expect(s.h.putCount()).toBe(0);
  await s.h.clock.advance(1);
  expect(s.session.state).toBe("closed");
  expect(s.h.putCount()).toBe(1);
  expect(s.h.callsFor("PUT")[0]!.headers.get("X-COOL-WOPI-IsExitSave")).toBe("true");
  expect(s.h.callsFor("UNLOCK")).toHaveLength(1);
  await s.h.clock.advance(30_000);
  expect(s.h.putCount()).toBe(1);
  expect((await fetch(`${s.base}/api/session/join`, { headers: s.headers })).status).toBe(401);
});

test("final save and handle invalidation complete Close while Unlock exceeds the close deadline", async () => {
  let release!: () => void;
  let entered!: () => void;
  const unlockGate = new Promise<void>((resolve) => { release = resolve; });
  const unlocking = new Promise<void>((resolve) => { entered = resolve; });
  const s = await setup(async () => { entered(); await unlockGate; });
  const { connectionId, ws } = await s.connect();
  s.client.edit("final-close-edit ");
  ws.send(encodeSequencedFrame({ connectionId, seq: 1 }, encodeUpdateFrame(s.client.encodeStateAsUpdate())) as unknown as ArrayBuffer);
  const ended = new Promise<void>((resolve) => s.session.on("userInvalidated", () => resolve()));
  const blocked: boolean[] = [];
  let finishes = 0;
  const requests: Promise<Response>[] = [];
  const closing = closeDocument({ flushPendingInput: async () => {}, hasPendingUpdates: () => false }, {
    async save(barrier) {
      const response = await fetch(`${s.base}/api/session/save`, { method: "POST", headers: s.headers, body: JSON.stringify(barrier) });
      expect(response.status).toBe(200);
      await response.text();
    },
    async close() {
      const request = fetch(`${s.base}/api/session/close`, { method: "POST", headers: s.headers });
      requests.push(request);
      const response = await request;
      expect(response.status).toBe(200);
      await response.text();
    },
  }, { captureBarrier: () => ({ connectionId, seq: 1 }), closeTimeoutMs: 100, whenClosed: ended, setInputBlocked: (value) => { blocked.push(value); } }, () => { finishes += 1; });
  try {
    await unlocking;
    await new Promise((resolve) => setTimeout(resolve, 150));
    await closing;
    expect(finishes).toBe(1);
    expect(s.h.putCount()).toBe(1);
    expect(blocked).toEqual([true]);
    expect(s.session.getUser(s.user.id)).toBeNull();
    const again = await fetch(`${s.base}/api/session/close`, { method: "POST", headers: s.headers });
    expect(again.status).toBe(200);
    expect(await again.json()).toEqual({ closed: true });
  } finally { release(); await Promise.all(requests); await closing.catch(() => {}); }
});

test("generation rejection preserves an offline edit through the recovery conflict-copy route", async () => {
  const s = await setup();
  const generation = s.session.baselineGeneration;
  const socket = await s.connect(generation);
  addToken(s.h, "keeper-token", "keeper");
  await s.registry.join({ wopiSrc: SRC, accessToken: "keeper-token", accessTokenTtl: 0, mode: "edit" });
  await s.drop(socket.ws);
  s.client.edit("offline-recovery-edit ");
  const external = await openFreshClient("xlsx", s.session.sourceBytes!, 998010);
  try {
    external.edit("new-host-baseline ");
    const file = s.h.host.files.get("file1")!;
    file.bytes = await external.save();
    file.itemVersion = "2";
    file.lock = null;
    file.lockExpiresAt = null;
  } finally { external.destroy(); }
  await s.h.clock.advance(31_000);
  await until(() => s.session.baselineGeneration === generation + 1);
  const rejected = new WebSocket(`${s.base.replace("http:", "ws:")}/collab`, ["bo-wopi.v1", s.user.id, `bo-wopi.generation.${generation}`]);
  const code = await new Promise<number>((resolve) => rejected.addEventListener("close", (event) => resolve(event.code), { once: true }));
  expect(code).toBe(4409);
  const original = s.h.host.files.get("file1")!.bytes.slice();
  const response = await fetch(`${s.base}/api/session/recovery`, { method: "POST", headers: s.headers, body: await s.client.save() as BodyInit });
  expect(response.status).toBe(200);
  const result = await response.json() as { copyName: string };
  expect(Object.keys(result)).toEqual(["copyName"]);
  expect(result.copyName).toContain("(conflict ");
  expect(JSON.stringify(result)).not.toContain("recovery-token");
  expect(s.h.putCount()).toBe(0);
  expect(s.h.host.files.get("file1")!.bytes).toEqual(original);
  const copy = [...s.h.host.files.values()].find((file) => file.id !== "file1")!;
  const fresh = await openFreshClient("xlsx", copy.bytes, 998003);
  try { expect(fresh.contains("offline-recovery-edit")).toBe(true); } finally { fresh.destroy(); }
});

test("recovery requires a valid writer handle and caps streamed bytes", async () => {
  const s = await setup();
  s.h.config.maxFileBytes = 3;
  const denied = await fetch(`${s.base}/api/session/recovery`, { method: "POST", body: "copy" });
  expect(denied.status).toBe(401);
  const oversized = await fetch(`${s.base}/api/session/recovery`, { method: "POST", headers: s.headers, body: new ReadableStream<Uint8Array>({
    start(controller) { controller.enqueue(new TextEncoder().encode("copy")); controller.close(); },
  }) });
  expect(oversized.status).toBe(413);
  expect(s.h.callsFor("PUT_RELATIVE")).toHaveLength(0);
});


test("Close emits a terminal event before a delayed Unlock completes", async () => {
  let release!: () => void;
  let entered!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  const unlocking = new Promise<void>((resolve) => { entered = resolve; });
  const s = await setup(async () => { entered(); await gate; });
  const response = await fetch(`${s.base}/api/session/events`, { headers: { ...s.headers, "X-WOPI-Baseline-Generation": String(s.session.baselineGeneration) } });
  const reader = response.body!.getReader();
  expect(new TextDecoder().decode((await reader.read()).value)).toContain(": connected");
  const closing = fetch(`${s.base}/api/session/close`, { method: "POST", headers: s.headers });
  try {
    await unlocking;
    const next = await reader.read();
    expect(new TextDecoder().decode(next.value)).toContain("event: closed");
    expect((await reader.read()).done).toBe(true);
  } finally { release(); await reader.cancel(); await closing; }
});


test("explicit Close during reconnect grace departs immediately and repeated Close succeeds", async () => {
  const s = await setup();
  const { ws } = await s.connect();
  await s.drop(ws);
  const close = () => fetch(`${s.base}/api/session/close`, { method: "POST", headers: s.headers });
  expect((await close()).status).toBe(200);
  expect(s.session.state).toBe("closed");
  expect(s.h.callsFor("UNLOCK")).toHaveLength(1);
  expect((await close()).status).toBe(200);
  await s.h.clock.advance(30_000);
  expect(s.h.callsFor("UNLOCK")).toHaveLength(1);
});

test("host token rejection during grace starts the hold immediately", async () => {
  const s = await setup();
  const { ws } = await s.connect();
  s.client.edit("held-after-rejection ");
  expect(s.session.applyUpdate(s.user.id, s.client.encodeStateAsUpdate()).accepted).toBe(true);
  await s.drop(ws);
  s.h.host.tokens.delete("recovery-token");
  await s.session.requestSave(s.user.id);
  await s.h.clock.flush();
  expect(s.session.awaitingFlushToken).toBe(true);
  expect(s.user.connected).toBe(false);
  expect(s.relay.resolve(s.user.id)).toBeNull();
});
