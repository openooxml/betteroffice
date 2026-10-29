import { afterEach, beforeAll, describe, expect, spyOn, test } from "bun:test";

import { generateKeyPair } from "../proof";
import { startServer } from "../server";
import { initEngines } from "../replica/engines";
import type { Format, JoinPayload } from "../replica/replica";
import {
  openFreshClient,
  openTestClient,
  sourceBytes,
  type TestClient,
} from "../../test/replica/clients";
import { addToken, makeHarness, SRC, type Harness } from "../../test/session/harness";
import { decodeJoinPayload, type EncodedJoinPayload } from "../sessionApi";
import { BASELINE_GENERATION_PROTOCOL } from "../baseline";
import {
  decodeMessages,
  encodeAwarenessFrame,
  encodeQueryAwarenessFrame,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdateFrame,
} from "./frames";
import { CollabRelay, type RelaySocket } from "./relay";
import { CollaborationProvider as DocxProvider } from "../../../../packages/docx/src/collaboration/index";
import { CollaborationProvider as XlsxProvider } from "../../../../packages/xlsx/src/collaboration/index";
import { CollaborationProvider as PptxProvider } from "../../../../packages/pptx/src/collaboration/index";
import type {
  CollaborationReplica,
  CollaborationTransport,
  CollaborationTransportEvent,
} from "../../../../packages/docx/src/collaboration/index";
import { createYrsSession } from "../../../../packages/docx/src/yrs";
import { openWorkbook } from "../../../../packages/xlsx/src/wasm/loader";
import { openPresentation } from "../../../../packages/pptx/src/wasm/loader";

beforeAll(async () => {
  await initEngines();
});

let servers: Array<{ stop: (close?: boolean) => void }> = [];
let cleanups: Array<() => void | Promise<void>> = [];
afterEach(async () => {
  for (const cleanup of cleanups.reverse()) {
    try {
      await cleanup();
    } catch {}
  }
  cleanups = [];
  for (const server of servers) server.stop(true);
  servers = [];
});

const TOKENS = ["SECRETTOKEN-relay-a-111", "SECRETTOKEN-relay-b-222", "SECRETTOKEN-relay-c-333"];

interface Setup {
  h: Harness;
  base: string;
  wsBase: string;
}

function setup(format: Format = "xlsx"): Setup {
  const h = makeHarness();
  const names: Record<Format, string> = { docx: "doc.docx", xlsx: "book.xlsx", pptx: "deck.pptx" };
  h.host.setFile("file1", sourceBytes(format), { name: names[format] });
  TOKENS.forEach((token, i) => addToken(h, token, `user-${i}`, { name: `User ${i}` }));
  addToken(h, "SECRETTOKEN-relay-view-444", "user-view", { name: "Viewer", canWrite: false });
  const server = startServer({ config: h.config, currentKey: generateKeyPair(), registry: h.registry }, 0);
  servers.push(server);
  return { h, base: `http://127.0.0.1:${server.port}`, wsBase: `ws://127.0.0.1:${server.port}` };
}

async function joinPage(s: Setup, token: string, mode: "edit" | "view" = "edit", ttl = 0): Promise<string> {
  const response = await fetch(`${s.base}/wopi/${mode}?WOPISrc=${encodeURIComponent(SRC)}`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ access_token: token, access_token_ttl: String(ttl) }).toString(),
  });
  expect(response.status).toBe(200);
  const html = await response.text();
  expect(html).not.toContain(token);
  const match = html.match(/<script type="application\/json" id="wopi-config">([\s\S]*?)<\/script>/);
  return (JSON.parse(match![1]!) as { sessionHandle: string }).sessionHandle;
}

type HydratedClient = TestClient & { baselineGeneration: number };

async function hydrate(s: Setup, handle: string, clientId: number): Promise<HydratedClient> {
  const headers = { Authorization: `Bearer ${handle}` };
  const source = new Uint8Array(await (await fetch(`${s.base}/api/session/source`, { headers })).arrayBuffer());
  const join = (await (await fetch(`${s.base}/api/session/join`, { headers })).json()) as EncodedJoinPayload & { baselineGeneration: number };
  return Object.assign(await openTestClient(decodeJoinPayload(join), source, clientId), { baselineGeneration: join.baselineGeneration });
}

async function waitFor(fn: () => boolean, label: string): Promise<void> {
  const deadline = Date.now() + 8000;
  while (!fn()) {
    if (Date.now() > deadline) throw new Error(`timeout waiting for ${label}`);
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

class FrameClient {
  client: TestClient;
  ws: WebSocket;
  closeCode: number | null = null;
  closeReason = "";
  #resolveClosed: () => void = () => {};
  #closed = new Promise<void>((resolve) => {
    this.#resolveClosed = resolve;
  });

  constructor(client: TestClient, ws: WebSocket) {
    this.client = client;
    this.ws = ws;
    ws.binaryType = "arraybuffer";
    ws.addEventListener("message", (event) => {
      const data = event.data;
      if (typeof data === "string") return;
      const frame = new Uint8Array(data as ArrayBuffer);
      for (const message of decodeMessages(frame)) {
        if (message.type === "sync-step-1") {
          this.send(encodeSyncStep2(this.client.encodeStateAsUpdate(message.stateVector)));
        } else if (message.type === "sync-step-2" || message.type === "update") {
          this.client.applyUpdate(message.update);
        }
      }
    });
    ws.addEventListener("close", (event) => {
      this.closeCode = event.code;
      this.closeReason = event.reason;
      this.#resolveClosed();
    });
  }

  static async connect(s: Setup, handle: string, client: HydratedClient): Promise<FrameClient> {
    const ws = new WebSocket(`${s.wsBase}/collab`, ["bo-wopi.v1", handle, `${BASELINE_GENERATION_PROTOCOL}${client.baselineGeneration}`]);
    const frameClient = new FrameClient(client, ws);
    await new Promise<void>((resolve, reject) => {
      ws.addEventListener("open", () => resolve(), { once: true });
      ws.addEventListener("error", () => reject(new Error("websocket error")), { once: true });
    });
    frameClient.send(encodeSyncStep1(client.encodeStateVector()));
    return frameClient;
  }

  send(frame: Uint8Array): void {
    if (this.ws.readyState === WebSocket.OPEN) this.ws.send(frame as unknown as ArrayBuffer);
  }

  editAndSend(marker: string, slot = 0): void {
    this.client.edit(marker, slot);
    this.send(encodeUpdateFrame(this.client.encodeStateAsUpdate()));
  }

  get closed(): Promise<void> {
    return this.#closed;
  }

  kill(): void {
    this.ws.close();
  }

  terminate(): void {
    (this.ws as unknown as { terminate?: () => void }).terminate?.() ?? this.ws.close();
  }
}

async function expectNoTokens(s: Setup, extra: string[] = []): Promise<void> {
  const all = [...s.h.logs, ...extra].join("\n");
  for (const token of [...TOKENS, "SECRETTOKEN-relay-view-444"]) expect(all).not.toContain(token);
}

describe("collab relay", () => {
  test("a socket delayed past source replacement is refused before attaching", async () => {
    const s = setup("pptx");
    const handle = await joinPage(s, "SECRETTOKEN-relay-view-444", "view");
    const staleClient = await hydrate(s, handle, 907001);
    const session = s.h.registry.list()[0]!;
    const external = await openFreshClient("pptx", sourceBytes("pptx"), 907101);
    cleanups.push(() => staleClient.destroy(), () => external.destroy(), () => session.abandon());
    external.edit("delayed-replacement ", 0);
    const file = s.h.host.files.get("file1")!;
    file.bytes = await external.save();
    file.itemVersion = "2";
    const writerHandle = await joinPage(s, TOKENS[0]!);
    const connect = spyOn(session, "connect");
    try {
      const stale = await FrameClient.connect(s, handle, staleClient);
      await waitFor(() => stale.closeCode !== null, "delayed source rejoin close");
      expect(stale.closeCode).toBe(4409);
      expect(stale.closeReason).toBe("rejoin from source");
      expect(connect).not.toHaveBeenCalled();
      expect(staleClient.contains("delayed-replacement")).toBe(false);

      const current = await hydrate(s, handle, 907201);
      expect(current.baselineGeneration).toBe(staleClient.baselineGeneration + 1);
      const viewer = await FrameClient.connect(s, handle, current);
      const writer = await FrameClient.connect(s, writerHandle, await hydrate(s, writerHandle, 907301));
      cleanups.push(() => current.destroy(), () => writer.client.destroy());
      expect(connect).toHaveBeenCalledTimes(2);
      writer.editAndSend("delayed-live ", 1);
      await waitFor(() => viewer.client.contains("delayed-live"), "current generation live edit");
      expect(viewer.ws.readyState).toBe(WebSocket.OPEN);
      expect(viewer.closeCode).toBeNull();
      await expectNoTokens(s, [stale.closeReason]);
    } finally {
      connect.mockRestore();
    }
  });

  for (const changed of ["version", "timestamp"] as const) {
    test(`source ${changed} replacement closes existing viewers with 4409 and new sockets stay open`, async () => {
      const s = setup("xlsx");
      const viewerHandles = [
        await joinPage(s, "SECRETTOKEN-relay-view-444", "view"),
        await joinPage(s, TOKENS[1]!, "view"),
      ];
      const viewers: FrameClient[] = [];
      for (let i = 0; i < viewerHandles.length; i += 1) {
        const handle = viewerHandles[i]!;
        const viewer = await FrameClient.connect(s, handle, await hydrate(s, handle, 908001 + i * 100));
        viewers.push(viewer);
        cleanups.push(() => viewer.client.destroy());
      }
      const session = s.h.registry.list()[0]!;
      cleanups.push(() => session.abandon());
      const oldReplica = session.replica;
      const external = await openFreshClient("xlsx", sourceBytes("xlsx"), 908201);
      cleanups.push(() => external.destroy());
      external.edit("replacement-baseline ", 0);
      const file = s.h.host.files.get("file1")!;
      file.bytes = await external.save();
      if (changed === "version") file.itemVersion = "2";
      else file.lastModifiedTime = "2026-10-28T12:00:00.000Z";

      const writerHandle = await joinPage(s, TOKENS[0]!);
      expect(session.replica).not.toBe(oldReplica);
      for (const viewer of viewers) {
        await waitFor(() => viewer.closeCode !== null, "source rejoin close");
        expect(viewer.closeCode).toBe(4409);
        expect(viewer.closeReason).toBe("rejoin from source");
        expect(viewer.client.contains("replacement-baseline")).toBe(false);
      }

      const writer = await FrameClient.connect(s, writerHandle, await hydrate(s, writerHandle, 908301));
      const rejoined = await FrameClient.connect(s, viewerHandles[0]!, await hydrate(s, viewerHandles[0]!, 908401));
      cleanups.push(() => { writer.client.destroy(); rejoined.client.destroy(); });
      expect(writer.client.contains("replacement-baseline")).toBe(true);
      expect(rejoined.client.contains("replacement-baseline")).toBe(true);
      writer.editAndSend("replacement-live ", 1);
      await waitFor(() => rejoined.client.contains("replacement-live"), "live edit after source rejoin");
      expect(writer.closeCode).toBeNull();
      expect(rejoined.closeCode).toBeNull();
      expect(writer.ws.readyState).toBe(WebSocket.OPEN);
      expect(rejoined.ws.readyState).toBe(WebSocket.OPEN);
      await expectNoTokens(s, viewers.map((viewer) => viewer.closeReason));
    });
  }

  for (const failure of [0, -1, "throw"] as const) {
    for (const delivery of ["broadcast", "sync"] as const) {
      test(`${delivery} delivery ${failure} closes and detaches the dropped socket`, async () => {
        const h = makeHarness();
        addToken(h, TOKENS[0]!, "user-a");
        addToken(h, TOKENS[1]!, "user-b");
        const a = await h.registry.join({ wopiSrc: SRC, accessToken: TOKENS[0]!, accessTokenTtl: 0, mode: "edit" });
        const b = await h.registry.join({ wopiSrc: SRC, accessToken: TOKENS[1]!, accessTokenTtl: 0, mode: "edit" });
        const relay = new CollabRelay(h.registry);
        const closed: number[] = [];
        const healthyClosed: number[] = [];
        const healthyFrames: Uint8Array[] = [];
        const healthy: RelaySocket = { send: (frame) => { healthyFrames.push(frame); return frame.length; }, close: (code) => { healthyClosed.push(code!); } };
        let attempts = 0;
        const dropped: RelaySocket = {
          send: () => {
            attempts += 1;
            if (failure === "throw") throw new Error("send failed");
            return failure;
          },
          close: (code) => { closed.push(code!); },
        };
        relay.attach(healthy, relay.resolve(a.user.id)!, a.session.baselineGeneration);
        relay.attach(dropped, relay.resolve(b.user.id)!, b.session.baselineGeneration);
        relay.handleMessage(dropped, encodeAwarenessFrame([{ clientId: 909101, clock: 1, state: { name: "Bob" } }]));
        const client = await openTestClient(a.joinPayload, a.session.sourceBytes!, 909001);
        const trigger = (): void => {
          if (delivery === "broadcast") {
            client.edit(`delivery-${attempts} `, 0);
            relay.handleMessage(healthy, encodeUpdateFrame(client.encodeStateAsUpdate()));
          } else {
            relay.handleMessage(dropped, encodeSyncStep1(client.encodeStateVector()));
          }
        };
        try {
          trigger();
          expect(closed).toEqual([1013]);
          expect(attempts).toBe(1);
          await h.clock.drain();
          expect(b.user.connected).toBe(false);
          expect(healthyFrames.some((frame) => decodeMessages(frame).some((message) => message.type === "awareness" && message.entries.some((entry) => entry.clientId === 909101 && entry.isNull)))).toBe(true);
          trigger();
          expect(attempts).toBe(1);
          const replacementFrames: Uint8Array[] = [];
          const replacement: RelaySocket = { send: (frame) => { replacementFrames.push(frame); return frame.length; }, close: () => {} };
          relay.attach(replacement, relay.resolve(b.user.id)!, b.session.baselineGeneration);
          relay.handleMessage(replacement, encodeSyncStep1(client.encodeStateVector()));
          expect(replacementFrames.some((frame) => decodeMessages(frame).some((message) => message.type === "sync-step-2"))).toBe(true);
          client.edit("delivery-recovers ", 0);
          relay.handleMessage(replacement, encodeUpdateFrame(client.encodeStateAsUpdate()));
          expect(healthyFrames.some((frame) => decodeMessages(frame).some((message) => message.type === "update"))).toBe(true);
          expect(healthyClosed).toEqual([]);
          expect(a.session.state).toBe("open");
        } finally {
          client.destroy();
          a.session.abandon();
        }
      });
    }
  }

  test("two clients converge, one autosave has every edit, late joiner catches up", async () => {
    const s = setup("xlsx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 910001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 910101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    await waitFor(() => a.ws.readyState === WebSocket.OPEN, "open");

    a.editAndSend("relay-alpha ", 0);
    await waitFor(() => b.client.contains("relay-alpha"), "b sees alpha");
    b.editAndSend("relay-bravo ", 1);
    await waitFor(() => a.client.contains("relay-bravo"), "a sees bravo");

    await s.h.clock.advance(30000);
    await s.h.clock.drain();
    expect(s.h.putCount()).toBe(1);
    const fresh = await openFreshClient("xlsx", s.h.host.files.get("file1")!.bytes, 910201);
    try {
      expect(fresh.contains("relay-alpha")).toBe(true);
      expect(fresh.contains("relay-bravo")).toBe(true);
    } finally {
      fresh.destroy();
    }

    const handleC = await joinPage(s, TOKENS[2]!);
    const c = await FrameClient.connect(s, handleC, await hydrate(s, handleC, 910301));
    cleanups.push(() => c.client.destroy());
    await waitFor(() => c.client.contains("relay-alpha") && c.client.contains("relay-bravo"), "late joiner");
    await expectNoTokens(s);
  });

  test("three clients converge", async () => {
    const s = setup("xlsx");
    const clients: FrameClient[] = [];
    for (let i = 0; i < 3; i += 1) {
      const handle = await joinPage(s, TOKENS[i]!);
      const frameClient = await FrameClient.connect(s, handle, await hydrate(s, handle, 920001 + i * 100));
      clients.push(frameClient);
      cleanups.push(() => frameClient.client.destroy());
    }
    const markers = ["relay-one ", "relay-two ", "relay-three "];
    clients.forEach((client, i) => client.editAndSend(markers[i]!, i));
    for (const client of clients) {
      await waitFor(() => markers.every((m) => client.client.contains(m.trim())), "convergence");
    }
    await s.h.clock.advance(30000);
    await s.h.clock.drain();
    expect(s.h.putCount()).toBe(1);
    await expectNoTokens(s);
  });

  test("read only users sync and receive edits, changing frames close with 4403", async () => {
    const s = setup("xlsx");
    const writerHandle = await joinPage(s, TOKENS[0]!);
    const writer = await FrameClient.connect(s, writerHandle, await hydrate(s, writerHandle, 930101));
    const handle = await joinPage(s, "SECRETTOKEN-relay-view-444", "view");
    const client = await hydrate(s, handle, 930001);
    cleanups.push(() => { client.destroy(); writer.client.destroy(); });
    const frameClient = await FrameClient.connect(s, handle, client);
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(frameClient.closeCode).toBeNull();
    writer.editAndSend("viewer-live ", 0);
    await waitFor(() => frameClient.client.contains("viewer-live"), "viewer receives live edit");
    expect(frameClient.closeCode).toBeNull();
    frameClient.client.edit("viewer-edit ", 1);
    frameClient.send(encodeUpdateFrame(frameClient.client.encodeStateAsUpdate()));
    await frameClient.closed;
    expect(frameClient.closeCode).toBe(4403);
    await expectNoTokens(s);
  });

  test("empty handshake replies are not broadcast, websocket payload limit is above the frame limit", async () => {
    const s = setup("xlsx");
    expect(new CollabRelay(s.h.registry).websocket.maxPayloadLength).toBeGreaterThan(16 * 1024 * 1024);
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 931001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 931101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    await new Promise((resolve) => setTimeout(resolve, 100));
    let updates = 0;
    b.ws.addEventListener("message", (event) => {
      if (typeof event.data === "string") return;
      for (const m of decodeMessages(new Uint8Array(event.data as ArrayBuffer))) {
        if (m.type === "update" || m.type === "sync-step-2") updates += 1;
      }
    });
    a.send(encodeSyncStep1(a.client.encodeStateVector()));
    await new Promise((resolve) => setTimeout(resolve, 150));
    expect(updates).toBe(0);
    a.client.edit("counted-edit ", 0);
    const counted = encodeUpdateFrame(a.client.encodeStateAsUpdate());
    a.send(counted);
    await waitFor(() => b.client.contains("counted-edit"), "b sees counted edit");
    expect(updates).toBeGreaterThan(0);
    const afterFirst = updates;
    a.send(counted);
    await new Promise((resolve) => setTimeout(resolve, 150));
    expect(updates).toBe(afterFirst);
    expect(a.closeCode).toBeNull();
    await expectNoTokens(s);
  });

  test("closing a user closes their collab socket, session close closes the room", async () => {
    const s = setup("xlsx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 932001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 932101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    const close = await fetch(`${s.base}/api/session/close`, {
      method: "POST",
      headers: { Authorization: `Bearer ${handleA}` },
    });
    expect(close.status).toBe(200);
    await close.text();
    await a.closed;
    expect(a.closeCode).toBe(4401);
    expect(b.closeCode).toBeNull();
    b.editAndSend("after-close ", 0);
    await new Promise((resolve) => setTimeout(resolve, 50));
    const closeB = await fetch(`${s.base}/api/session/close`, {
      method: "POST",
      headers: { Authorization: `Bearer ${handleB}` },
    });
    expect(closeB.status).toBe(200);
    await closeB.text();
    await b.closed;
    expect(b.closeCode).toBe(4401);
    await expectNoTokens(s);
  });

  test("corrupt update closes only its socket with 4400", async () => {
    const s = setup("xlsx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 940001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 940101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    a.send(encodeUpdateFrame(new Uint8Array([9, 8, 7, 6, 5, 4, 3, 2, 1])));
    await a.closed;
    expect(a.closeCode).toBe(4400);
    b.editAndSend("still-working ", 0);
    const session = s.h.registry.list()[0]!;
    await waitFor(() => session.dirty, "server applied b update");
    await s.h.clock.advance(30000);
    await s.h.clock.drain();
    expect(s.h.putCount()).toBe(1);
    await expectNoTokens(s);
  });

  test("docx save replicaUpdate reaches every socket and both clients converge with the server", async () => {
    const s = setup("docx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 933001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 933101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    await new Promise((resolve) => setTimeout(resolve, 100));
    a.client.splitFirstParagraph();
    a.send(encodeUpdateFrame(a.client.encodeStateAsUpdate()));
    const session = s.h.registry.list()[0]!;
    await waitFor(
      () => [...b.client.encodeStateVector()].join(",") === [...session.replica!.encodeStateVector()].join(","),
      "b synced after split",
    );
    const save = await fetch(`${s.base}/api/session/save`, {
      method: "POST",
      headers: { Authorization: `Bearer ${handleA}` },
    });
    expect(save.status).toBe(200);
    await save.text();
    const serverVector = () => [...session.replica!.encodeStateVector()].join(",");
    await waitFor(
      () => [...a.client.encodeStateVector()].join(",") === serverVector() && [...b.client.encodeStateVector()].join(",") === serverVector(),
      "both clients converge with server after save",
    );
    expect(a.closeCode).toBeNull();
    expect(b.closeCode).toBeNull();
    await expectNoTokens(s);
  });

  test("docx identity repair follow-ups arrive after the author update and converge every client", async () => {
    const s = setup("docx");
    const handleP = await joinPage(s, TOKENS[0]!);
    const handleA = await joinPage(s, TOKENS[1]!);
    const peer = await FrameClient.connect(s, handleP, await hydrate(s, handleP, 935001));
    const author = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 935101));
    cleanups.push(() => { peer.client.destroy(); author.client.destroy(); });
    const session = s.h.registry.list()[0]!;
    const synced = (client: TestClient): boolean =>
      [...client.encodeStateVector()].join(",") === [...session.replica!.encodeStateVector()].join(",");
    await waitFor(() => synced(peer.client) && synced(author.client), "initial sync");

    peer.client.splitFirstParagraph();
    peer.send(encodeUpdateFrame(peer.client.encodeStateAsUpdate()));
    await waitFor(() => synced(peer.client) && synced(author.client), "split synced");

    const save = await fetch(`${s.base}/api/session/save`, {
      method: "POST",
      headers: { Authorization: `Bearer ${handleP}` },
    });
    expect(save.status).toBe(200);
    await save.text();
    let savedId = "";
    await waitFor(() => {
      if (!synced(peer.client)) return false;
      try {
        savedId = peer.client.secondParagraphOoxmlId();
        return savedId.length > 0;
      } catch {
        return false;
      }
    }, "saved identity synced to peer");

    peer.client.deleteSplitSecondHalf();
    peer.send(encodeUpdateFrame(peer.client.encodeStateAsUpdate()));
    await waitFor(() => synced(peer.client) && synced(author.client), "delete synced");

    author.client.insertPilcrowClaiming(savedId);
    author.send(encodeUpdateFrame(author.client.encodeStateAsUpdate()));
    await waitFor(() => synced(peer.client), "peer converges after repair");
    await waitFor(() => synced(author.client), "author converges after repair");
    expect(peer.closeCode).toBeNull();
    expect(author.closeCode).toBeNull();
    await expectNoTokens(s);
  });

  test("token validity uses the session clock, not the wall clock", async () => {
    const s = setup("xlsx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 936001));
    const handleB = await joinPage(s, TOKENS[1]!, "edit", s.h.clock.t + 60000);
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 936101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    b.editAndSend("before-expiry ", 0);
    await waitFor(() => a.client.contains("before-expiry"), "edit accepted while fake clock token is valid");
    expect(b.closeCode).toBeNull();
    await s.h.clock.advance(50000);
    await s.h.clock.drain();
    a.send(encodeSyncStep1(a.client.encodeStateVector()));
    b.send(encodeSyncStep1(b.client.encodeStateVector()));
    await s.h.clock.drain();
    await s.h.clock.advance(20000);
    await s.h.clock.drain();
    expect(a.closeCode).toBeNull();
    await b.closed;
    expect(b.closeCode).toBe(4401);
    a.editAndSend("after-expiry ", 1);
    await waitFor(() => s.h.registry.list()[0]!.dirty, "valid peer keeps editing");
    expect(b.client.contains("after-expiry")).toBe(false);
    await expectNoTokens(s);
  });

  test("xlsx causalGap answers with sync step 1, is not broadcast, and resyncs", async () => {
    const s = setup("xlsx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 934001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 934101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    await new Promise((resolve) => setTimeout(resolve, 100));
    const gappedPayloadsAtB: Uint8Array[] = [];
    let step1AtA = 0;
    b.ws.addEventListener("message", (event) => {
      if (typeof event.data === "string") return;
      for (const m of decodeMessages(new Uint8Array(event.data as ArrayBuffer))) {
        if (m.type === "update" || m.type === "sync-step-2") gappedPayloadsAtB.push(m.update);
      }
    });
    a.ws.addEventListener("message", (event) => {
      if (typeof event.data === "string") return;
      for (const m of decodeMessages(new Uint8Array(event.data as ArrayBuffer))) {
        if (m.type === "sync-step-1") step1AtA += 1;
      }
    });
    a.client.edit("gap-first ", 0);
    const afterFirst = a.client.encodeStateVector();
    a.client.edit("gap-second ", 1);
    const second = a.client.encodeStateAsUpdate(afterFirst);
    const step1Before = step1AtA;
    a.send(encodeUpdateFrame(second));
    await waitFor(() => b.client.contains("gap-first") && b.client.contains("gap-second"), "resync converges");
    expect(step1AtA).toBeGreaterThan(step1Before);
    expect(a.closeCode).toBeNull();
    expect(gappedPayloadsAtB.some((p) => [...p].join(",") === [...second].join(","))).toBe(false);
    await expectNoTokens(s);
  });

  for (const { reason, closeCode } of [
    { reason: "causalGap", closeCode: null },
    { reason: "engine rejected update (Error)", closeCode: 4400 },
    { reason: "user is read only", closeCode: 4403 },
  ]) {
    test(`rejected docx follow-ups reach every socket before ${reason} recovery`, async () => {
      const h = makeHarness({ format: "docx" });
      addToken(h, TOKENS[0]!, "user-a");
      addToken(h, TOKENS[1]!, "user-b");
      const a = await h.registry.join({ wopiSrc: SRC, accessToken: TOKENS[0]!, accessTokenTtl: 0, mode: "edit" });
      const b = await h.registry.join({ wopiSrc: SRC, accessToken: TOKENS[1]!, accessTokenTtl: 0, mode: "edit" });
      const author = await openTestClient(a.joinPayload, a.session.sourceBytes!, 934501);
      const peer = await openTestClient(b.joinPayload, b.session.sourceBytes!, 934601);
      const relay = new CollabRelay(h.registry);
      const authorFrames: Uint8Array[] = [];
      const peerFrames: Uint8Array[] = [];
      const authorClosed: number[] = [];
      const peerClosed: number[] = [];
      const events: string[] = [];
      const authorSocket: RelaySocket = {
        send: (frame) => { authorFrames.push(frame); events.push("author frame"); return frame.length; },
        close: (code) => { authorClosed.push(code!); events.push("author close"); },
      };
      const peerSocket: RelaySocket = {
        send: (frame) => { peerFrames.push(frame); events.push("peer frame"); return frame.length; },
        close: (code) => { peerClosed.push(code!); },
      };
      relay.attach(authorSocket, relay.resolve(a.user.id)!, a.session.baselineGeneration);
      relay.attach(peerSocket, relay.resolve(b.user.id)!, b.session.baselineGeneration);
      const applyUpdate = a.session.applyUpdate.bind(a.session);
      const mockedApply = spyOn(a.session, "applyUpdate");
      try {
        const initialVector = a.session.replica!.encodeStateVector();
        author.edit("integrated-first ");
        const first = author.encodeStateAsUpdate(initialVector);
        const afterFirst = author.encodeStateVector();
        author.edit("integrated-second ");
        const second = author.encodeStateAsUpdate(afterFirst);
        author.edit("unintegrated-tail ");
        const rejected = author.encodeStateAsUpdate(initialVector);
        const followUps = [first, second];
        mockedApply.mockImplementationOnce((userId, update) => {
          expect(userId).toBe(a.user.id);
          expect(update).toEqual(rejected);
          for (const followUp of followUps) expect(applyUpdate(userId, followUp).accepted).toBe(true);
          return { accepted: false, reason, followUps };
        });

        relay.handleMessage(authorSocket, encodeUpdateFrame(rejected));
        mockedApply.mockRestore();
        expect(authorFrames.slice(0, 2)).toEqual(followUps.map(encodeUpdateFrame));
        expect(peerFrames).toEqual(followUps.map(encodeUpdateFrame));
        expect(events).toEqual([
          "author frame", "peer frame", "author frame", "peer frame",
          closeCode === null ? "author frame" : "author close",
        ]);
        expect(authorClosed).toEqual(closeCode === null ? [] : [closeCode]);
        expect(peerClosed).toEqual([]);
        expect(a.session.dirty).toBe(true);
        for (const [client, frames] of [[author, authorFrames.slice(0, 2)], [peer, peerFrames]] as const) {
          for (const frame of frames) {
            for (const message of decodeMessages(frame)) {
              if (message.type === "update") client.applyUpdate(message.update);
            }
          }
        }
        expect(peer.contains("integrated-first")).toBe(true);
        expect(peer.contains("integrated-second")).toBe(true);
        expect(peer.contains("unintegrated-tail")).toBe(false);
        expect(peer.encodeStateVector()).toEqual(a.session.replica!.encodeStateVector());

        if (closeCode === null) {
          expect(authorFrames).toHaveLength(3);
          expect(authorFrames[2]).toEqual(encodeSyncStep1(a.session.replica!.encodeStateVector()));
          const [resync] = decodeMessages(authorFrames[2]!);
          if (resync?.type !== "sync-step-1") throw new Error("missing sender resync");
          relay.handleMessage(authorSocket, encodeSyncStep2(author.encodeStateAsUpdate(resync.stateVector)));
          for (const frame of peerFrames.slice(2)) {
            for (const message of decodeMessages(frame)) {
              if (message.type === "update") peer.applyUpdate(message.update);
            }
          }
          expect(peer.contains("unintegrated-tail")).toBe(true);
          expect(author.encodeStateVector()).toEqual(a.session.replica!.encodeStateVector());
          expect(peer.encodeStateVector()).toEqual(a.session.replica!.encodeStateVector());
          expect(authorClosed).toEqual([]);
          expect(peerClosed).toEqual([]);
        } else {
          expect(authorFrames).toHaveLength(2);
        }
        for (const token of TOKENS) expect(h.logs.join("\n")).not.toContain(token);
      } finally {
        mockedApply.mockRestore();
        a.session.abandon();
        author.destroy();
        peer.destroy();
      }
    });
  }

  test("oversize frame closes with 1009", async () => {
    const s = setup("xlsx");
    const handle = await joinPage(s, TOKENS[0]!);
    const client = await hydrate(s, handle, 950001);
    cleanups.push(() => client.destroy());
    const frameClient = await FrameClient.connect(s, handle, client);
    const big = new Uint8Array(16 * 1024 * 1024 + 1024);
    big.set(encodeSyncStep1(client.encodeStateVector()));
    frameClient.ws.send(big as unknown as ArrayBuffer);
    await frameClient.closed;
    expect(frameClient.closeCode).toBe(1009);
    await expectNoTokens(s);
  });

  test("a second socket for one user replaces the first", async () => {
    const s = setup("xlsx");
    const handle = await joinPage(s, TOKENS[0]!);
    const first = await FrameClient.connect(s, handle, await hydrate(s, handle, 955001));
    const secondClient = await hydrate(s, handle, 955101);
    cleanups.push(() => { first.client.destroy(); secondClient.destroy(); });
    const second = await FrameClient.connect(s, handle, secondClient);
    void second;
    await first.closed;
    expect(first.closeCode).toBe(1000);
    await expectNoTokens(s);
  });

  test("replacement during a pending save ignores the old socket's queued disconnect", async () => {
    let releasePut = (): void => {};
    let enteredPut = (): void => {};
    const pendingPut = new Promise<void>((resolve) => { releasePut = resolve; });
    const putEntered = new Promise<void>((resolve) => { enteredPut = resolve; });
    const h = makeHarness({ beforePut: async () => { enteredPut(); await pendingPut; } });
    addToken(h, TOKENS[0]!, "user-a");
    const a = await h.registry.join({ wopiSrc: SRC, accessToken: TOKENS[0]!, accessTokenTtl: 0, mode: "edit" });
    const relay = new CollabRelay(h.registry);
    const first: RelaySocket = { send: (frame) => frame.length, close: () => {} };
    const replacementClosed: number[] = [];
    const replacement: RelaySocket = { send: (frame) => frame.length, close: (code) => { replacementClosed.push(code!); } };
    relay.attach(first, relay.resolve(a.user.id)!, a.session.baselineGeneration);
    const firstGeneration = a.user.connectionGeneration;
    const client = await openTestClient(a.joinPayload, a.session.sourceBytes!, 955501);
    let save: Promise<unknown> | undefined;
    try {
      client.edit("pending-save ", 0);
      relay.handleMessage(first, encodeUpdateFrame(client.encodeStateAsUpdate()));
      save = a.session.requestSave(a.user.id);
      await putEntered;
      relay.detach(first);
      relay.attach(replacement, relay.resolve(a.user.id)!, a.session.baselineGeneration);
      expect(a.user.connectionGeneration).toBeGreaterThan(firstGeneration);
      releasePut();
      await save;
      await h.clock.drain();
      expect(a.user.connected).toBe(true);
      expect(a.session.state).toBe("open");
      expect(a.session.lockHeld).toBe(true);
      expect(h.callsFor("UNLOCK")).toHaveLength(0);
      expect(replacementClosed).toEqual([]);
      relay.websocket.close(first);
      client.edit("replacement-edit ", 1);
      relay.handleMessage(replacement, encodeUpdateFrame(client.encodeStateAsUpdate()));
      expect(a.session.dirty).toBe(true);
      await a.session.requestSave(a.user.id);
      const saved = await openFreshClient("xlsx", h.host.files.get("file1")!.bytes, 955601);
      try {
        expect(saved.contains("pending-save")).toBe(true);
        expect(saved.contains("replacement-edit")).toBe(true);
      } finally {
        saved.destroy();
      }
      expect(replacementClosed).toEqual([]);
    } finally {
      releasePut();
      await save;
      client.destroy();
      a.session.abandon();
    }
  });

  test("replacement broadcasts awareness removal for the old client", async () => {
    const s = setup("xlsx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 951001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 951101));
    const secondClient = await hydrate(s, handleA, 951201);
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); secondClient.destroy(); });
    const removals: number[] = [];
    b.ws.addEventListener("message", (event) => {
      if (typeof event.data === "string") return;
      for (const m of decodeMessages(new Uint8Array(event.data as ArrayBuffer))) {
        if (m.type === "awareness") for (const e of m.entries) if (e.isNull) removals.push(e.clientId);
      }
    });
    a.send(encodeAwarenessFrame([{ clientId: a.client.clientId, clock: 1, state: { user: { name: "User 0", color: "#000000" }, cursor: null } }]));
    await new Promise((resolve) => setTimeout(resolve, 50));
    await FrameClient.connect(s, handleA, secondClient);
    await waitFor(() => removals.includes(a.client.clientId), "removal on replacement");
    await expectNoTokens(s);
  });

  test("awareness leave clears the cache and another socket cannot clobber a client id", async () => {
    const s = setup("xlsx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const handleC = await joinPage(s, TOKENS[2]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 952001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 952101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    const removalsAtA: number[] = [];
    a.ws.addEventListener("message", (event) => {
      if (typeof event.data === "string") return;
      for (const m of decodeMessages(new Uint8Array(event.data as ArrayBuffer))) {
        if (m.type === "awareness") for (const e of m.entries) if (e.isNull) removalsAtA.push(e.clientId);
      }
    });
    a.send(encodeAwarenessFrame([{ clientId: a.client.clientId, clock: 1, state: { user: { name: "User 0", color: "#000000" }, cursor: null } }]));
    await new Promise((resolve) => setTimeout(resolve, 50));
    b.send(encodeAwarenessFrame([{ clientId: a.client.clientId, clock: 9, state: { user: { name: "Fake", color: "#000000" }, cursor: null } }]));
    await new Promise((resolve) => setTimeout(resolve, 50));
    b.ws.close();
    await b.closed;
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(removalsAtA).not.toContain(a.client.clientId);
    a.send(encodeAwarenessFrame([{ clientId: a.client.clientId, clock: 2, state: null }]));
    await new Promise((resolve) => setTimeout(resolve, 50));
    const c = await FrameClient.connect(s, handleC, await hydrate(s, handleC, 952201));
    cleanups.push(() => c.client.destroy());
    let queried: number[] = [];
    c.ws.addEventListener("message", (event) => {
      if (typeof event.data === "string") return;
      for (const m of decodeMessages(new Uint8Array(event.data as ArrayBuffer))) {
        if (m.type === "awareness") queried.push(...m.entries.map((e) => e.clientId));
      }
    });
    c.send(encodeQueryAwarenessFrame());
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(queried.filter((id) => id === a.client.clientId).length).toBe(0);
    await expectNoTokens(s);
  });

  test("awareness is relayed to other sockets", async () => {
    const s = setup("xlsx");
    const handleA = await joinPage(s, TOKENS[0]!);
    const handleB = await joinPage(s, TOKENS[1]!);
    const a = await FrameClient.connect(s, handleA, await hydrate(s, handleA, 960001));
    const b = await FrameClient.connect(s, handleB, await hydrate(s, handleB, 960101));
    cleanups.push(() => { a.client.destroy(); b.client.destroy(); });
    let saw = false;
    b.ws.addEventListener("message", (event) => {
      if (typeof event.data === "string") return;
      const messages = decodeMessages(new Uint8Array(event.data as ArrayBuffer));
      if (messages.some((m) => m.type === "awareness")) saw = true;
    });
    a.send(encodeAwarenessFrame([{ clientId: a.client.clientId, clock: 1, state: { user: { name: "User 0", color: "#000000" }, cursor: null } }]));
    await waitFor(() => saw, "awareness relay");
    await expectNoTokens(s);
  });

  test("disconnect without close triggers last leave flush on the fake clock", async () => {
    const s = setup("xlsx");
    const handle = await joinPage(s, TOKENS[0]!);
    const a = await FrameClient.connect(s, handle, await hydrate(s, handle, 970001));
    cleanups.push(() => a.client.destroy());
    a.editAndSend("flush-marker ", 0);
    await waitFor(() => s.h.registry.list()[0]?.dirty === true, "dirty");
    a.terminate();
    await a.closed.catch(() => {});
    await new Promise((resolve) => setTimeout(resolve, 50));
    await s.h.clock.advance(65000);
    await s.h.clock.drain();
    await waitFor(() => s.h.putCount() >= 1, "flush put");
    const fresh = await openFreshClient("xlsx", s.h.host.files.get("file1")!.bytes, 970101);
    try {
      expect(fresh.contains("flush-marker")).toBe(true);
    } finally {
      fresh.destroy();
    }
    expect(s.h.callsFor("UNLOCK").length).toBeGreaterThanOrEqual(1);
    await expectNoTokens(s);
  });
});

class WsTransport implements CollaborationTransport {
  ws: WebSocket | null = null;
  listeners = new Set<(event: CollaborationTransportEvent) => void>();
  constructor(private url: string, private handle: string, private baselineGeneration: number) {}
  connect(): void {
    const ws = new WebSocket(this.url, ["bo-wopi.v1", this.handle, `${BASELINE_GENERATION_PROTOCOL}${this.baselineGeneration}`]);
    this.ws = ws;
    ws.binaryType = "arraybuffer";
    ws.addEventListener("open", () => this.emit({ type: "open" }));
    ws.addEventListener("message", (event) => {
      if (typeof event.data !== "string") this.emit({ type: "message", data: new Uint8Array(event.data as ArrayBuffer) });
    });
    ws.addEventListener("close", () => this.emit({ type: "close" }));
    ws.addEventListener("error", () => this.emit({ type: "error", error: new Error("ws error") }));
  }
  disconnect(): void {
    this.ws?.close();
    this.ws = null;
  }
  send(data: Uint8Array): boolean {
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) return false;
    this.ws.send(data.slice() as unknown as ArrayBuffer);
    return true;
  }
  onEvent(listener: (event: CollaborationTransportEvent) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
  emit(event: CollaborationTransportEvent): void {
    for (const listener of [...this.listeners]) listener(event);
  }
}

type ProviderTestClient = Pick<TestClient, "edit" | "contains" | "destroy">;

async function providerClient(
  s: Setup,
  format: Format,
  handle: string,
  clientId: number,
): Promise<{ provider: { connect(): void; destroy(): void; synced: boolean }; test: ProviderTestClient }> {
  const headers = { Authorization: `Bearer ${handle}` };
  const source = new Uint8Array(await (await fetch(`${s.base}/api/session/source`, { headers })).arrayBuffer());
  const join = (await (await fetch(`${s.base}/api/session/join`, { headers })).json()) as EncodedJoinPayload & { baselineGeneration: number };
  const payload: JoinPayload = decodeJoinPayload(join);
  let replica: CollaborationReplica;
  let provider: { connect(): void; destroy(): void; synced: boolean };
  let wrapped: ProviderTestClient;
  if (format === "docx") {
    const session = await createYrsSession({ clientId });
    if (payload.format !== "docx") throw new Error("payload format");
    session.openDocx(source.slice(), false);
    session.loadState(payload.initialUpdate.slice());
    replica = session as unknown as CollaborationReplica;
    provider = new DocxProvider(replica as never, new WsTransport(`${s.wsBase}/collab`, handle, join.baselineGeneration)) as never;
    wrapped = {
      edit: (marker, slot = 0) => {
        const paragraphs = session.paragraphs("body");
        const para = slot === 0 ? paragraphs[0]! : paragraphs.at(-1)!;
        session.insertText({ story: "body", paraId: para.paraId, offset: 0 }, marker);
      },
      contains: (marker) => session.storyIds().sort().flatMap((story) => session.paragraphs(story).map((p) => p.text)).join("\n").includes(marker),
      destroy: () => { session.destroy(); },
    };
  } else if (format === "xlsx") {
    const workbook = openWorkbook(source.slice(), { collaborative: true, clientId });
    if (payload.format !== "xlsx") throw new Error("payload format");
    workbook.applyUpdate(payload.initialUpdate.slice());
    replica = workbook as unknown as CollaborationReplica;
    provider = new XlsxProvider(replica as never, new WsTransport(`${s.wsBase}/collab`, handle, join.baselineGeneration)) as never;
    wrapped = {
      edit: (marker, slot = 0) => { workbook.editCell(0, 25 + slot, 0, marker); },
      contains: (marker) => {
        for (let row = 20; row <= 32; row += 1) for (let col = 0; col <= 2; col += 1) if (workbook.cell(0, row, col).input.includes(marker)) return true;
        return false;
      },
      destroy: () => workbook.dispose(),
    };
  } else {
    const presentation = openPresentation(source.slice(), { clientId, initialUpdate: payload.format === "pptx" ? payload.initialUpdate.slice() : undefined });
    replica = presentation as unknown as CollaborationReplica;
    provider = new PptxProvider(replica as never, new WsTransport(`${s.wsBase}/collab`, handle, join.baselineGeneration)) as never;
    wrapped = {
      edit: (marker, slot = 0) => {
        const stories = presentation.snapshot().slides.flatMap((slide) => slide.shapes.flatMap((shape) => shape.textStories));
        presentation.insertText(stories[slot % stories.length]!.id, 0, marker);
      },
      contains: (marker) => JSON.stringify(presentation.snapshot()).includes(marker),
      destroy: () => presentation.dispose(),
    };
  }
  return { provider, test: wrapped };
}

describe("real CollaborationProvider clients", () => {
  for (const format of ["docx", "xlsx", "pptx"] as Format[]) {
    test(`${format} provider converges with a frame client through the relay`, async () => {
      const s = setup(format);
      const handleP = await joinPage(s, TOKENS[0]!);
      const handleF = await joinPage(s, TOKENS[1]!);
      const provided = await providerClient(s, format, handleP, 980001);
      const frame = await FrameClient.connect(s, handleF, await hydrate(s, handleF, 980101));
      cleanups.push(() => provided.provider.destroy(), () => frame.client.destroy(), () => provided.test.destroy());
      provided.provider.connect();
      await waitFor(() => provided.provider.synced, "provider synced");

      frame.editAndSend(`provider-${format}-peer `, 0);
      await waitFor(() => provided.test.contains(`provider-${format}-peer`), "provider sees peer");

      provided.test.edit(`provider-${format}-own `, 1);
      await waitFor(() => frame.client.contains(`provider-${format}-own`), "frame client sees provider");
      await expectNoTokens(s);
    });
  }
});
