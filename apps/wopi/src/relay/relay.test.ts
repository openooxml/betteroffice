import { afterEach, beforeAll, describe, expect, test } from "bun:test";

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
import {
  decodeMessages,
  encodeAwarenessFrame,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdateFrame,
} from "./frames";
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
  for (const cleanup of cleanups.reverse()) await cleanup().catch(() => {});
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

async function joinPage(s: Setup, token: string, mode: "edit" | "view" = "edit"): Promise<string> {
  const response = await fetch(`${s.base}/wopi/${mode}?WOPISrc=${encodeURIComponent(SRC)}`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ access_token: token, access_token_ttl: "0" }).toString(),
  });
  expect(response.status).toBe(200);
  const html = await response.text();
  expect(html).not.toContain(token);
  const match = html.match(/<script type="application\/json" id="wopi-config">([\s\S]*?)<\/script>/);
  return (JSON.parse(match![1]!) as { sessionHandle: string }).sessionHandle;
}

async function hydrate(s: Setup, handle: string, clientId: number): Promise<TestClient> {
  const headers = { Authorization: `Bearer ${handle}` };
  const source = new Uint8Array(await (await fetch(`${s.base}/api/session/source`, { headers })).arrayBuffer());
  const join = (await (await fetch(`${s.base}/api/session/join`, { headers })).json()) as EncodedJoinPayload;
  return openTestClient(decodeJoinPayload(join), source, clientId);
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

  static async connect(s: Setup, handle: string, client: TestClient): Promise<FrameClient> {
    const ws = new WebSocket(`${s.wsBase}/collab`, ["bo-wopi.v1", handle]);
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
    expect(s.h.putCount()).toBe(1);
    await expectNoTokens(s);
  });

  test("read only frames close with 4403", async () => {
    const s = setup("xlsx");
    const handle = await joinPage(s, "SECRETTOKEN-relay-view-444", "view");
    const client = await hydrate(s, handle, 930001);
    cleanups.push(() => client.destroy());
    const frameClient = await FrameClient.connect(s, handle, client);
    frameClient.client.edit("viewer-edit ", 0);
    frameClient.send(encodeUpdateFrame(frameClient.client.encodeStateAsUpdate()));
    await frameClient.closed;
    expect(frameClient.closeCode).toBe(4403);
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
    await s.h.clock.advance(30000);
    expect(s.h.putCount()).toBe(1);
    await expectNoTokens(s);
  });

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
  constructor(private url: string, private handle: string) {}
  connect(): void {
    const ws = new WebSocket(this.url, ["bo-wopi.v1", this.handle]);
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

async function providerClient(
  s: Setup,
  format: Format,
  handle: string,
  clientId: number,
): Promise<{ provider: { connect(): void; destroy(): void; synced: boolean }; test: TestClient }> {
  const headers = { Authorization: `Bearer ${handle}` };
  const source = new Uint8Array(await (await fetch(`${s.base}/api/session/source`, { headers })).arrayBuffer());
  const payload: JoinPayload = decodeJoinPayload(
    (await (await fetch(`${s.base}/api/session/join`, { headers })).json()) as EncodedJoinPayload,
  );
  let replica: CollaborationReplica;
  let provider: { connect(): void; destroy(): void; synced: boolean };
  let wrapped: TestClient;
  if (format === "docx") {
    const hostSession = await createYrsSession({ clientId: clientId + 100000 });
    const host = hostSession.openDocx(source.slice(), false);
    const session = await createYrsSession({ clientId });
    if (payload.format !== "docx") throw new Error("payload format");
    session.bootstrapPeer(payload.state, payload.peerMetadata, source.slice(), host);
    replica = session as unknown as CollaborationReplica;
    provider = new DocxProvider(replica as never, new WsTransport(`${s.wsBase}/collab`, handle)) as never;
    wrapped = {
      format, clientId,
      encodeStateVector: () => session.encodeStateVector(),
      encodeStateAsUpdate: (remote) => session.encodeStateAsUpdate(remote),
      applyUpdate: (update) => { session.applyUpdate(update.slice()); },
      edit: (marker, slot = 0) => {
        const paragraphs = session.paragraphs("body");
        const para = slot === 0 ? paragraphs[0]! : paragraphs.at(-1)!;
        session.insertText({ story: "body", paraId: para.paraId, offset: 0 }, marker);
      },
      remove: () => { throw new Error("unused"); },
      contains: (marker) => session.storyIds().sort().flatMap((story) => session.paragraphs(story).map((p) => p.text)).join("\n").includes(marker),
      save: async () => { throw new Error("unused"); },
      destroy: () => { session.destroy(); hostSession.destroy(); },
    };
  } else if (format === "xlsx") {
    const workbook = openWorkbook(source.slice(), { collaborative: true, clientId });
    if (payload.format !== "xlsx") throw new Error("payload format");
    workbook.applyUpdate(payload.update.slice());
    replica = workbook as unknown as CollaborationReplica;
    provider = new XlsxProvider(replica as never, new WsTransport(`${s.wsBase}/collab`, handle)) as never;
    wrapped = {
      format, clientId,
      encodeStateVector: () => workbook.encodeStateVector(),
      encodeStateAsUpdate: (remote) => workbook.encodeStateAsUpdate(remote),
      applyUpdate: (update) => { workbook.applyUpdate(update.slice()); },
      edit: (marker, slot = 0) => { workbook.editCell(0, 25 + slot, 0, marker); },
      remove: () => { throw new Error("unused"); },
      contains: (marker) => {
        for (let row = 20; row <= 32; row += 1) for (let col = 0; col <= 2; col += 1) if (workbook.cell(0, row, col).input.includes(marker)) return true;
        return false;
      },
      save: async () => workbook.save(),
      destroy: () => workbook.dispose(),
    };
  } else {
    const presentation = openPresentation(source.slice(), { clientId, initialUpdate: payload.format === "pptx" ? payload.initialUpdate.slice() : undefined });
    replica = presentation as unknown as CollaborationReplica;
    provider = new PptxProvider(replica as never, new WsTransport(`${s.wsBase}/collab`, handle)) as never;
    wrapped = {
      format, clientId,
      encodeStateVector: () => presentation.encodeStateVector(),
      encodeStateAsUpdate: (remote) => presentation.encodeStateAsUpdate(remote),
      applyUpdate: (update) => { presentation.applyUpdate(update.slice()); },
      edit: (marker, slot = 0) => {
        const stories = presentation.snapshot().slides.flatMap((slide) => slide.shapes.flatMap((shape) => shape.textStories));
        presentation.insertText(stories[slot % stories.length]!.id, 0, marker);
      },
      remove: () => { throw new Error("unused"); },
      contains: (marker) => JSON.stringify(presentation.snapshot()).includes(marker),
      save: async () => presentation.save(),
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
