import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from "bun:test";
import { useEffect } from "react";
import type { EditorApi, EditorProps, PreparedFile } from "../../../../shared/editor-host/Editor";
import { SessionApi } from "./api";
import { parseWopiConfig } from "./config";
import { PostMessageBridge } from "./postMessage";
import { decodeSequencedFrame } from "../../src/relay/sequence";
import type { WebSocketLike } from "./transport";
import type { Loaded } from "./page";
import { applyUpdate, Doc, encodeStateAsUpdate, encodeStateVector } from "yjs";
import { classifyFrame, encodeSyncStep1, encodeSyncStep2, encodeUpdateFrame } from "../../src/relay/frames";

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const mounts: EditorProps[] = [];
const lifecycle: string[] = [];
let flush = async () => {};
let publish = (_update: Uint8Array) => false;
let serverDoc: Doc;
let attachGeneration: number | null = null;
const prepareFile = mock(async (bytes: Uint8Array, format: PreparedFile["format"]) => ({ bytes, format }));
mock.module("../../../../shared/editor-host/Editor", () => ({
  prepareFile,
  default: (props: EditorProps) => {
    useEffect(() => {
      mounts.push(props);
      lifecycle.push("mount");
      const transport = props.createTransport!("wopi");
      publish = (update) => transport.send(encodeUpdateFrame(update));
      const unsubscribe = transport.onEvent((event) => {
        if (event.type === "open") transport.send(Uint8Array.of(8));
        if (event.type === "message" && classifyFrame(event.data).messages.some((message) => message.type === "sync-step-1")) {
          transport.send(encodeSyncStep2(Uint8Array.of(0, 0)));
        }
      });
      transport.connect();
      const api: EditorApi = {
        hasPendingUpdates: () => false, flushPendingInput: () => flush(), whenReady: async () => {},
        encodeStateVector: async () => Uint8Array.of(0), serialize: async () => Uint8Array.of(1),
        focus() {}, print() {},
      };
      props.onReady(api);
      return () => {
        lifecycle.push("destroy");
        unsubscribe();
        transport.disconnect();
        props.onReady(null);
      };
    }, []);
    return <div data-testid="editor" />;
  },
}));

const { act, cleanup, render } = await import("@testing-library/react");
const { default: WopiApp } = await import("./page");

const sockets: TestSocket[] = [];
class TestSocket implements WebSocketLike {
  binaryType = "";
  readyState = 0;
  bufferedAmount = 0;
  onopen: (() => void) | null = null;
  onmessage: ((event: { data: unknown }) => void) | null = null;
  onerror: ((event: unknown) => void) | null = null;
  onclose: ((event: { code?: number; reason?: string }) => void) | null = null;
  sent: ArrayBuffer[] = [];

  constructor(readonly url: string, readonly protocols: string[]) { sockets.push(this); }
  send(data: ArrayBuffer) {
    this.sent.push(data);
    const frame = decodeSequencedFrame(new Uint8Array(data));
    if (!frame || frame.seq === 0) return;
    const parsed = classifyFrame(frame.payload);
    if (parsed.kind !== "document") return;
    for (const message of parsed.messages) {
      if (message.type === "update") applyUpdate(serverDoc, message.update);
    }
  }
  close() { this.readyState = 3; }
  open() {
    if (attachGeneration !== null && this.protocols[2] !== `bo-wopi.generation.${attachGeneration}`) {
      this.drop(4409);
      return;
    }
    this.readyState = 1;
    this.onopen?.();
    this.onmessage?.({ data: encodeSyncStep1(Uint8Array.of(0)) });
  }
  drop(code: number) { this.readyState = 3; this.onclose?.({ code, reason: "rejoin from source" }); }
}

const socketDescriptor = Object.getOwnPropertyDescriptor(globalThis, "WebSocket");
Object.defineProperty(globalThis, "WebSocket", { configurable: true, writable: true, value: TestSocket });

let previousUrl: string;
beforeEach(() => {
  serverDoc = new Doc();
  previousUrl = window.location.href;
  (window as unknown as GlobalWindow).happyDOM.setURL("https://office.example/wopi/edit");
});
afterEach(() => {
  cleanup();
  serverDoc.destroy();
  attachGeneration = null;
  (window as unknown as GlobalWindow).happyDOM.setURL(previousUrl);
  sockets.length = 0;
  mounts.length = 0;
  lifecycle.length = 0;
  flush = async () => {};
  prepareFile.mockClear();
});
afterAll(async () => {
  if (socketDescriptor) Object.defineProperty(globalThis, "WebSocket", socketDescriptor);
  else Reflect.deleteProperty(globalThis, "WebSocket");
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount(replaceBeforeAttach = false) {
  const calls: string[] = [];
  const messages: Array<{ MessageId: string; SendTime: number; Values: Record<string, unknown> }> = [];
  const streams: ReadableStreamDefaultController<Uint8Array>[] = [];
  let sourceGate = Promise.resolve();
  let failSource = false;
  let generation = 1;
  let hostText = "";
  let closeRequest = async (_signal?: AbortSignal | null): Promise<Response> => Response.json({ closed: true });
  const api = new SessionApi("same-handle", async (input, init) => {
    expect(new Headers(init?.headers).get("Authorization")).toBe("Bearer same-handle");
    const path = String(input);
    calls.push(path);
    if (path.endsWith("/source")) {
      lifecycle.push("source");
      await sourceGate;
      return failSource ? new Response(null, { status: 503 }) : new Response(Uint8Array.of(2), { headers: { "X-WOPI-Baseline-Generation": String(generation) } });
    }
    if (path.endsWith("/join")) return Response.json({ format: "docx", initialUpdate: "Aw==", baselineGeneration: generation });
    if (path.endsWith("/events")) {
      return new Response(new ReadableStream<Uint8Array>({
        start(controller) {
          streams.push(controller);
          init?.signal?.addEventListener("abort", () => controller.close(), { once: true });
        },
      }));
    }
    if (path.endsWith("/save")) { hostText = serverDoc.getText("body").toString(); return Response.json({ itemVersion: "2" }); }
    if (path.endsWith("/close")) return closeRequest(init?.signal);
    throw new Error(`Unexpected request: ${path}`);
  });
  const config = parseWopiConfig({
    sessionHandle: "same-handle", format: "docx", fileName: "file.docx", canWrite: true,
    postMessageOrigin: "https://host.example", userFriendlyName: "Peer",
  });
  const bridge = new PostMessageBridge({
    parent: { postMessage(message) { messages.push(JSON.parse(message)); } },
    addEventListener() {}, removeEventListener() {},
  }, config.postMessageOrigin, {
    onSave: (options) => loaded.actions.save!(options), onClose() {}, onBlurFocus() {}, onGrabFocus() {}, onGetViews: () => [],
  });
  const loaded: Loaded = {
    config, api, bridge, file: { bytes: Uint8Array.of(0), format: "docx" },
    join: { format: "docx", initialUpdate: Uint8Array.of(1), baselineGeneration: 1 }, clientId: 7, actions: { views: [] },
  };
  let view!: ReturnType<typeof render>;
  bridge.frameReady();
  await act(async () => { view = render(<WopiApp loaded={loaded} />); });
  if (replaceBeforeAttach) { generation = 2; attachGeneration = generation; }
  await act(async () => sockets[0]!.open());
  return {
    calls, messages, loaded, view,
    get hostText() { return hostText; },
    setCloseRequest(request: typeof closeRequest) { closeRequest = request; },
    blockSource() { let release!: () => void; sourceGate = new Promise((resolve) => { release = resolve; }); return release; },
    failSource() { failSource = true; },
    event() { streams.at(-1)!.enqueue(new TextEncoder().encode('event: rejoinRequired\ndata: {"reason":"sourceChanged"}\n\n')); },
  };
}

test.each(["socket", "event", "socket and event", "event and socket"])("%s replaces the baseline exactly once", async (signal) => {
  const h = await mount();
  const first = sockets[0]!;
  const oldClose = first.onclose!;
  const barrier = decodeSequencedFrame(new Uint8Array(first.sent[0]!))!;
  const release = h.blockSource();
  const oldMount = mounts[0]!;
  await act(async () => {
    oldMount.onChange();
    if (signal === "socket") first.drop(4409);
    else if (signal === "event") h.event();
    else if (signal === "socket and event") { h.event(); first.drop(4409); }
    else { h.event(); await Promise.resolve(); first.drop(4409); }
  });
  expect(h.view.queryByTestId("editor")).toBeNull();
  expect(lifecycle).toEqual(["mount", "destroy", "source"]);
  expect(sockets).toHaveLength(1);
  expect(h.calls.filter((path) => path.endsWith("/source"))).toHaveLength(1);
  await expect(h.loaded.actions.save!()).rejects.toThrow("rejoinRequired");
  await act(async () => release());
  expect(mounts).toHaveLength(2);
  expect(h.calls.filter((path) => path.endsWith("/source"))).toHaveLength(1);
  expect(h.calls.filter((path) => path.endsWith("/join"))).toHaveLength(1);
  expect(prepareFile).toHaveBeenCalledWith(Uint8Array.of(2), "docx", { fontBaseUrl: "/fonts/" });
  expect(mounts[1]!.file.bytes).toEqual(Uint8Array.of(2));
  expect(mounts[1]!.initialUpdate).toEqual(Uint8Array.of(3));
  expect(sockets).toHaveLength(2);
  await act(async () => sockets[1]!.open());
  const newBarrier = decodeSequencedFrame(new Uint8Array(sockets[1]!.sent[0]!))!;
  expect(newBarrier.connectionId).not.toBe(barrier.connectionId);
  expect(newBarrier.seq).toBe(0);
  expect(sockets[1]!.protocols).toEqual(["bo-wopi.v1", "same-handle", "bo-wopi.generation.1"]);
  expect(h.messages.filter((message) => message.Values.Status === "Document_Loaded")).toHaveLength(2);
  expect(h.messages.filter((message) => message.Values.Status === "Frame_Ready")).toHaveLength(2);
  await act(async () => {
    oldClose({ code: 4409 });
    oldMount.onChange();
    await new Promise((resolve) => setTimeout(resolve, 300));
  });
  expect(sockets).toHaveLength(2);
  expect(mounts).toHaveLength(2);
  expect(h.calls.filter((path) => path.endsWith("/source"))).toHaveLength(1);
  expect(h.messages.filter((message) => message.Values.Modified === true)).toHaveLength(1);
});

test("rejoin rejects an in-flight host save and a late flush cannot save the replacement", async () => {
  const h = await mount();
  let release!: () => void;
  flush = () => new Promise<void>((resolve) => { release = resolve; });
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    await act(async () => {
      h.loaded.bridge.handleMessage({ origin: "https://host.example", data: { MessageId: "Action_Save", Values: { Notify: true } } });
    });
    expect(release).toBeDefined();
    await act(async () => sockets[0]!.drop(4409));
    expect(h.messages.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([
      { MessageId: "Action_Save_Resp", SendTime: expect.any(Number), Values: { success: false, result: "rejoinRequired" } },
    ]);
    await act(async () => release());
    expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(0);
    expect(mounts).toHaveLength(2);
  } finally {
    warning.mockRestore();
  }
});

test.each([1006, 4401])("close code %i reconnects the existing editor without refetching", async (code) => {
  const h = await mount();
  await act(async () => {
    sockets[0]!.drop(code);
    await new Promise((resolve) => setTimeout(resolve, 300));
  });
  expect(sockets).toHaveLength(2);
  expect(mounts).toHaveLength(1);
  expect(lifecycle).toEqual(["mount"]);
  expect(h.calls.filter((path) => path.endsWith("/source") || path.endsWith("/join"))).toHaveLength(0);
});

test("a failed refetch reports failure and leaves the old editor destroyed", async () => {
  const h = await mount();
  h.failSource();
  await act(async () => sockets[0]!.drop(4409));
  expect(h.view.queryByTestId("editor")).toBeNull();
  expect(h.view.getByRole("alert").textContent).toContain("503");
  expect(h.messages.filter((message) => message.Values.Status === "Failed")).toHaveLength(1);
  expect(mounts).toHaveLength(1);
});

test("deleting while disconnected then Close reports notSynced and keeps the editor open", async () => {
  const h = await mount();
  const doc = new Doc();
  doc.getText("body").insert(0, "before");
  const vector = encodeStateVector(doc);
  flush = async () => { doc.getText("body").delete(0, 6); };
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    await act(async () => {
      sockets[0]!.drop(1006);
      await h.loaded.actions.close!();
    });
    expect(encodeStateVector(doc)).toEqual(vector);
    expect(doc.getText("body").toString()).toBe("");
    expect(h.calls.filter((path) => path.endsWith("/save") || path.endsWith("/close"))).toEqual([]);
    expect(h.view.getByRole("alert").textContent).toBe("notSynced");
    expect(h.view.getByTestId("editor")).toBeDefined();
    expect(h.view.container.querySelector("main")!.inert).toBe(false);
  } finally { doc.destroy(); warning.mockRestore(); }
});

test("a stalled Close restores input and clears the shared promise so another Close retries", async () => {
  const h = await mount();
  const signals: Array<AbortSignal | null | undefined> = [];
  h.setCloseRequest((signal) => {
    signals.push(signal);
    return signals.length === 1 ? new Promise<Response>(() => {}) : Promise.resolve(Response.json({ closed: true }));
  });
  const schedule = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  const deadlines = new Map<ReturnType<typeof setTimeout>, () => void>();
  const timers = spyOn(globalThis, "setTimeout").mockImplementation(((callback: () => void, delay: number) => {
    if (delay !== 10_000) return schedule(callback, delay);
    const timer = schedule(() => {}, delay);
    deadlines.set(timer, callback);
    return timer;
  }) as typeof setTimeout);
  const cancellations = spyOn(globalThis, "clearTimeout").mockImplementation(((timer: ReturnType<typeof setTimeout>) => {
    deadlines.delete(timer);
    cancel(timer);
  }) as typeof clearTimeout);
  try {
    let first!: Promise<void>;
    await act(async () => { first = h.loaded.actions.close!(); });
    expect(h.loaded.actions.close!()).toBe(first);
    expect(h.view.container.querySelector("main")!.inert).toBe(true);
    expect(deadlines.size).toBe(1);
    await act(async () => { for (const deadline of [...deadlines.values()]) deadline(); });
    expect(signals).toHaveLength(1);
    expect(signals[0]?.aborted).toBe(true);
    expect(h.view.container.querySelector("main")!.inert).toBe(false);
    expect(h.view.getByRole("alert").textContent).toBe("closeTimeout");
    expect(h.view.getByTestId("editor")).toBeDefined();
    await act(async () => { await h.loaded.actions.close!(); });
    expect(signals).toHaveLength(2);
    expect(signals[1]).not.toBe(signals[0]);
    expect(h.view.queryByTestId("editor")).toBeNull();
  } finally {
    for (const timer of deadlines.keys()) cancel(timer);
    timers.mockRestore();
    cancellations.mockRestore();
  }
});

test("Action_Save with DontSaveIfUnmodified saves before idle change notifications", async () => {
  const h = await mount();
  const doc = new Doc();
  flush = async () => {
    doc.getText("body").insert(0, "pending edit");
    expect(publish(encodeStateAsUpdate(doc))).toBe(true);
  };
  try {
    await act(async () => {
      h.loaded.bridge.handleMessage({ origin: "https://host.example", data: {
        MessageId: "Action_Save", Values: { DontSaveIfUnmodified: true, Notify: true },
      } });
    });
    expect(h.messages.filter((message) => message.Values.Modified === true)).toEqual([]);
    expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(1);
    expect(h.hostText).toBe("pending edit");
    expect(h.messages.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([
      { MessageId: "Action_Save_Resp", SendTime: expect.any(Number), Values: { success: true } },
    ]);
  } finally { doc.destroy(); }
});

test("replacement after join but before socket attach refetches exactly once", async () => {
  const h = await mount(true);
  expect(h.calls.filter((path) => path.endsWith("/source"))).toHaveLength(1);
  expect(h.calls.filter((path) => path.endsWith("/join"))).toHaveLength(1);
  expect(mounts).toHaveLength(2);
  expect(mounts[1]!.file.bytes).toEqual(Uint8Array.of(2));
  expect(sockets).toHaveLength(2);
  await act(async () => sockets[1]!.open());
  expect(sockets[1]!.readyState).toBe(1);
  expect(sockets[1]!.protocols[2]).toBe("bo-wopi.generation.2");
  expect(h.calls.filter((path) => path.endsWith("/source"))).toHaveLength(1);
  expect(h.messages.filter((message) => message.Values.Status === "Document_Loaded")).toHaveLength(2);
});

test("source and join from different generations retry before preparing an editor", async () => {
  const calls: string[] = [];
  let sources = 0;
  const api = new SessionApi("same-handle", async (input) => {
    const path = String(input);
    calls.push(path);
    if (path.endsWith("/source")) {
      sources += 1;
      return new Response(Uint8Array.of(sources), { headers: { "X-WOPI-Baseline-Generation": String(sources) } });
    }
    return Response.json({ format: "docx", initialUpdate: "Aw==", baselineGeneration: 2 });
  });
  const { loadBaseline } = await import("./page");
  const baseline = await loadBaseline(api);
  expect(calls).toEqual(["/api/session/source", "/api/session/join", "/api/session/source", "/api/session/join"]);
  expect(prepareFile).toHaveBeenCalledTimes(1);
  expect(baseline.file.bytes).toEqual(Uint8Array.of(2));
  expect(baseline.join.baselineGeneration).toBe(2);
});
