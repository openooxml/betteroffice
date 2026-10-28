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

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const mounts: EditorProps[] = [];
const lifecycle: string[] = [];
let flush = async () => {};
const prepareFile = mock(async (bytes: Uint8Array, format: PreparedFile["format"]) => ({ bytes, format }));
mock.module("../../../../shared/editor-host/Editor", () => ({
  prepareFile,
  default: (props: EditorProps) => {
    useEffect(() => {
      mounts.push(props);
      lifecycle.push("mount");
      const transport = props.createTransport!("wopi");
      const unsubscribe = transport.onEvent((event) => {
        if (event.type === "open") transport.send(Uint8Array.of(8));
      });
      transport.connect();
      const api: EditorApi = {
        flushPendingInput: () => flush(), whenReady: async () => {},
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
  send(data: ArrayBuffer) { this.sent.push(data); }
  close() { this.readyState = 3; }
  open() { this.readyState = 1; this.onopen?.(); }
  drop(code: number) { this.readyState = 3; this.onclose?.({ code, reason: "rejoin from source" }); }
}

const socketDescriptor = Object.getOwnPropertyDescriptor(globalThis, "WebSocket");
Object.defineProperty(globalThis, "WebSocket", { configurable: true, writable: true, value: TestSocket });

let previousUrl: string;
beforeEach(() => {
  previousUrl = window.location.href;
  (window as unknown as GlobalWindow).happyDOM.setURL("https://office.example/wopi/edit");
});
afterEach(() => {
  cleanup();
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

async function mount() {
  const calls: string[] = [];
  const messages: Array<{ MessageId: string; SendTime: number; Values: Record<string, unknown> }> = [];
  const streams: ReadableStreamDefaultController<Uint8Array>[] = [];
  let sourceGate = Promise.resolve();
  let failSource = false;
  const api = new SessionApi("same-handle", async (input, init) => {
    expect(new Headers(init?.headers).get("Authorization")).toBe("Bearer same-handle");
    const path = String(input);
    calls.push(path);
    if (path.endsWith("/source")) {
      lifecycle.push("source");
      await sourceGate;
      return failSource ? new Response(null, { status: 503 }) : new Response(Uint8Array.of(2));
    }
    if (path.endsWith("/join")) return Response.json({ format: "docx", initialUpdate: "Aw==" });
    if (path.endsWith("/events")) {
      return new Response(new ReadableStream<Uint8Array>({
        start(controller) {
          streams.push(controller);
          init?.signal?.addEventListener("abort", () => controller.close(), { once: true });
        },
      }));
    }
    if (path.endsWith("/save")) return Response.json({ itemVersion: "2" });
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
    join: { format: "docx", initialUpdate: Uint8Array.of(1) }, clientId: 7, actions: { views: [] },
  };
  let view!: ReturnType<typeof render>;
  bridge.frameReady();
  await act(async () => { view = render(<WopiApp loaded={loaded} />); });
  await act(async () => sockets[0]!.open());
  return {
    calls, messages, loaded, view,
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
  expect(sockets[1]!.protocols).toEqual(["bo-wopi.v1", "same-handle"]);
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
