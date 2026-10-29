import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from "bun:test";
import { useEffect } from "react";
import * as ReactDOMClient from "react-dom/client";
import type { EditorApi, EditorProps, PreparedFile } from "../../../../shared/editor-host/Editor";
import { SessionApi } from "./api";
import { parseWopiConfig } from "./config";
import { PostMessageBridge } from "./postMessage";
import { decodeSequencedFrame, encodeSequenceAck } from "../../src/relay/sequence";
import type { WebSocketLike } from "./transport";
import type { Loaded } from "./page";
import { applyUpdate, Doc, encodeStateAsUpdate, encodeStateVector } from "yjs";
import { classifyFrame, encodeSyncStep1, encodeSyncStep2, encodeUpdateFrame } from "../../src/relay/frames";

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const mounts: EditorProps[] = [];
const lifecycle: string[] = [];
let flush: EditorApi["flushPendingInput"] = async () => {};
let hasPendingUpdates = () => false;
let serialize: EditorApi["serialize"] = async () => Uint8Array.of(1);
let publish = (_update: Uint8Array) => false;
let serverDoc: Doc;
let attachGeneration: number | null = null;
let acknowledge = true;
let readOnlyRelay = false;
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
        hasPendingUpdates: () => hasPendingUpdates(), flushPendingInput: (options) => flush(options), whenReady: async () => {},
        encodeStateVector: async () => Uint8Array.of(0), serialize: () => serialize(),
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
    return <div data-testid="editor" data-readonly={String(props.readOnly ?? false)} />;
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
    if (acknowledge) queueMicrotask(() => this.onmessage?.({ data: encodeSequenceAck(frame) }));
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
    if (!readOnlyRelay) this.onmessage?.({ data: encodeSyncStep1(Uint8Array.of(0)) });
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
  acknowledge = true;
  readOnlyRelay = false;
  (window as unknown as GlobalWindow).happyDOM.setURL(previousUrl);
  sockets.length = 0;
  mounts.length = 0;
  lifecycle.length = 0;
  flush = async () => {};
  hasPendingUpdates = () => false;
  serialize = async () => Uint8Array.of(1);
  prepareFile.mockClear();
});
afterAll(async () => {
  if (socketDescriptor) Object.defineProperty(globalThis, "WebSocket", socketDescriptor);
  else Reflect.deleteProperty(globalThis, "WebSocket");
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount(replaceBeforeAttach = false, format: PreparedFile["format"] = "docx", options: {
  canWrite?: boolean;
  credentialCanWrite?: boolean;
  readOnlyRelay?: boolean;
  snapshot?: { canWrite: boolean; viewOnlyReason: string | null };
} = {}) {
  let canWrite = options.canWrite ?? true;
  let snapshot = options.snapshot;
  readOnlyRelay = options.readOnlyRelay ?? false;
  const calls: string[] = [];
  const messages: Array<{ MessageId: string; SendTime: number; Values: Record<string, unknown> }> = [];
  const streams: ReadableStreamDefaultController<Uint8Array>[] = [];
  let sourceGate = Promise.resolve();
  let failSource = false;
  let active = true;
  let failRecovery = false;
  const copies: Uint8Array[] = [];
  let generation = 1;
  let hostText = "";
  let closeRequest = async (_signal?: AbortSignal | null): Promise<Response> => Response.json({ closed: true });
  let recoveryRequest = async (_signal?: AbortSignal | null): Promise<Response> =>
    failRecovery ? new Response(null, { status: 502 }) : Response.json({ copyName: "file (conflict recovered).docx" });
  const api = new SessionApi("same-handle", async (input, init) => {
    expect(new Headers(init?.headers).get("Authorization")).toBe("Bearer same-handle");
    const path = String(input);
    calls.push(path);
    if (!active && path.endsWith("/join")) return new Response(null, { status: 401 });
    if (path.endsWith("/source")) {
      lifecycle.push("source");
      await sourceGate;
      return failSource ? new Response(null, { status: 503 }) : new Response(Uint8Array.of(2), { headers: { "X-WOPI-Baseline-Generation": String(generation) } });
    }
    if (path.endsWith("/join")) return Response.json({ format, initialUpdate: "Aw==", baselineGeneration: generation, canWrite });
    if (path.endsWith("/events")) {
      return new Response(new ReadableStream<Uint8Array>({
        start(controller) {
          streams.push(controller);
          if (snapshot) controller.enqueue(new TextEncoder().encode(`event: snapshot\ndata: ${JSON.stringify(snapshot)}\n\n`));
          init?.signal?.addEventListener("abort", () => { try { controller.close(); } catch {} }, { once: true });
        },
      }));
    }
    if (path.endsWith("/recovery")) {
      copies.push(new Uint8Array(await new Response(init?.body).arrayBuffer()));
      return recoveryRequest(init?.signal);
    }
    if (path.endsWith("/save")) { hostText = serverDoc.getText("body").toString(); return Response.json({ itemVersion: "2" }); }
    if (path.endsWith("/close")) return closeRequest(init?.signal);
    throw new Error(`Unexpected request: ${path}`);
  });
  const config = parseWopiConfig({
    sessionHandle: "same-handle", format, fileName: `file.${format}`, canWrite: options.credentialCanWrite ?? true,
    postMessageOrigin: "https://host.example", userFriendlyName: "Peer", closePostMessage: true,
  });
  const bridge = new PostMessageBridge({
    parent: { postMessage(message) { messages.push(JSON.parse(message)); } },
    addEventListener() {}, removeEventListener() {},
  }, config.postMessageOrigin, {
    onSave: (options) => loaded.actions.save!(options), onClose() {}, onBlurFocus() {}, onGrabFocus() {}, onGetViews: () => [],
  });
  const loaded: Loaded = {
    config, api, bridge, file: { bytes: Uint8Array.of(0), format },
    join: { format, initialUpdate: Uint8Array.of(1), baselineGeneration: 1, canWrite }, clientId: 7, actions: { views: [] },
  };
  let view!: ReturnType<typeof render>;
  bridge.frameReady();
  await act(async () => { view = render(<WopiApp loaded={loaded} />); });
  if (replaceBeforeAttach) { generation = 2; attachGeneration = generation; }
  await act(async () => sockets[0]!.open());
  return {
    calls, messages, loaded, view, copies,
    endSession() { active = false; },
    failRecovery() { failRecovery = true; },
    get hostText() { return hostText; },
    setCloseRequest(request: typeof closeRequest) { closeRequest = request; },
    setRecoveryRequest(request: typeof recoveryRequest) { recoveryRequest = request; },
    blockSource() { let release!: () => void; sourceGate = new Promise((resolve) => { release = resolve; }); return release; },
    failSource() { failSource = true; },
    closedEvent() { streams.at(-1)!.enqueue(new TextEncoder().encode("event: closed\ndata: {}\n\n")); },
    viewOnlyEvent() {
      canWrite = false;
      readOnlyRelay = true;
      streams.at(-1)!.enqueue(new TextEncoder().encode('event: viewOnly\ndata: {"reason":"lockLost"}\n\n'));
    },
    writableEvent(writable = true) {
      canWrite = writable;
      readOnlyRelay = !writable;
      streams.at(-1)!.enqueue(new TextEncoder().encode(`event: writable\ndata: ${JSON.stringify({ canWrite })}\n\n`));
    },
    disconnectEvents(permission: NonNullable<typeof snapshot>) { snapshot = permission; streams.at(-1)!.close(); },
    snapshotEvent(writable: boolean, viewOnlyReason: string | null) {
      streams.at(-1)!.enqueue(new TextEncoder().encode(`event: snapshot\ndata: ${JSON.stringify({ canWrite: writable, viewOnlyReason })}\n\n`));
    },
    event() { streams.at(-1)!.enqueue(new TextEncoder().encode('event: rejoinRequired\ndata: {"reason":"sourceChanged"}\n\n')); },
  };
}

const compositionSaves = (["docx", "xlsx", "pptx"] as const).flatMap((format) =>
  (["Action_Save", "Save button", "Close"] as const).map((action) => [format, action] as const));

test.each(["docx", "xlsx", "pptx"] as const)("%s initially view-only joins reject Save and complete Close without outbound resync", async (format) => {
  const h = await mount(false, format, { canWrite: false, readOnlyRelay: true });
  const flushing = mock(async () => {});
  flush = flushing;
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    expect(h.loaded.config.canWrite).toBe(true);
    expect(h.view.getByTestId("editor").getAttribute("data-readonly")).toBe("true");
    await act(async () => {
      await expect(Promise.resolve(mounts[0]!.onSaveRequest!())).rejects.toThrow("notAllowed");
      h.loaded.bridge.handleMessage({ origin: "https://host.example", data: {
        MessageId: "Action_Save", Values: { Notify: true },
      } });
    });
    expect(h.messages.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([
      { MessageId: "Action_Save_Resp", SendTime: expect.any(Number), Values: { success: false, result: "notAllowed" } },
    ]);
    expect(flushing).not.toHaveBeenCalled();
    expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
    await act(async () => { await h.loaded.actions.close!(); });
    expect(flushing).not.toHaveBeenCalled();
    expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
    expect(h.calls.filter((path) => path.endsWith("/close"))).toHaveLength(1);
    expect(h.view.queryByTestId("editor")).toBeNull();
    expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
  } finally { warning.mockRestore(); }
});

test("session permission events update editing and subsequent Close decisions", async () => {
  const h = await mount(false, "xlsx", { canWrite: false, readOnlyRelay: true });
  expect(h.view.getByTestId("editor").getAttribute("data-readonly")).toBe("true");
  await act(async () => h.writableEvent());
  expect(h.view.getByTestId("editor").getAttribute("data-readonly")).toBe("false");
  await act(async () => h.viewOnlyEvent());
  expect(h.view.getByTestId("editor").getAttribute("data-readonly")).toBe("true");
  const flushing = mock(async () => {});
  flush = flushing;
  await act(async () => { await h.loaded.actions.close!(); });
  expect(flushing).not.toHaveBeenCalled();
  expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(0);
  expect(h.calls.filter((path) => path.endsWith("/close"))).toHaveLength(1);
  expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
});

test.each((["docx", "xlsx", "pptx"] as const).flatMap((format) =>
  (["provider", "unacknowledged"] as const).map((pending) => [format, pending] as const),
))("%s read-only Close preserves %s work before notifying the host", async (format, pending) => {
  const h = await mount(false, format);
  const local = new Doc();
  const recovered = new Doc();
  let blob: Blob | undefined;
  let downloadName = "";
  const createUrl = spyOn(URL, "createObjectURL").mockImplementation((value) => {
    blob = value as Blob;
    return "blob:recovery";
  });
  const revokeUrl = spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
  const click = spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
    downloadName = this.download;
  });
  try {
    local.getText("body").insert(0, "queued before permission loss");
    sockets[0]!.bufferedAmount = 2 * 1024 * 1024;
    expect(publish(encodeStateAsUpdate(local))).toBe(false);
    if (pending === "provider") hasPendingUpdates = () => true;
    else mounts[0]!.onLocalUpdate!();
    const serializing = mock(async () => encodeStateAsUpdate(local));
    serialize = serializing;
    await act(async () => h.viewOnlyEvent());
    await act(async () => { await h.loaded.actions.close!(); });
    expect(serializing).toHaveBeenCalledTimes(1);
    expect(h.view.getByRole("button", { name: "Download recovered document" })).toBeDefined();
    expect(h.view.getByTestId("editor")).toBeDefined();
    expect(lifecycle).not.toContain("destroy");
    expect(h.calls.filter((path) => path.endsWith("/save") || path.endsWith("/close"))).toEqual([]);
    expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toEqual([]);
    await act(async () => h.view.getByRole("button", { name: "Download recovered document" }).click());
    applyUpdate(recovered, new Uint8Array(await blob!.arrayBuffer()));
    expect(recovered.getText("body").toString()).toBe("queued before permission loss");
    expect(downloadName).toMatch(new RegExp(`^file \\(recovered .+\\)\\.${format}$`));
    expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
    expect(h.calls.filter((path) => path.endsWith("/close"))).toHaveLength(1);
    expect(h.copies).toEqual([]);
    expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
    expect(h.view.queryByTestId("editor")).toBeNull();
  } finally {
    local.destroy(); recovered.destroy();
    createUrl.mockRestore(); revokeUrl.mockRestore(); click.mockRestore();
  }
});

test("a failed read-only Close recovery keeps unsent work open for another Close", async () => {
  const h = await mount();
  hasPendingUpdates = () => true;
  serialize = async () => { throw new Error("serialization failed"); };
  await act(async () => h.viewOnlyEvent());
  await act(async () => { await h.loaded.actions.close!(); });
  expect(h.view.getByRole("alert").textContent).toContain("serialization failed");
  expect(h.view.getByTestId("editor")).toBeDefined();
  expect(h.view.container.querySelector("main")!.inert).toBe(false);
  expect(h.calls.filter((path) => path.endsWith("/close"))).toEqual([]);
  expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toEqual([]);
  serialize = async () => Uint8Array.of(7);
  await act(async () => { await h.loaded.actions.close!(); });
  expect(h.view.getByRole("button", { name: "Download recovered document" })).toBeDefined();
  expect(lifecycle).not.toContain("destroy");
});

test.each(["first attachment", "reconnection"])("permission snapshots apply on %s", async (attachment) => {
  const permission = { canWrite: false, viewOnlyReason: "lockedElsewhere" };
  const h = await mount(false, "docx", { snapshot: attachment === "first attachment" ? permission : undefined });
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    if (attachment === "reconnection") {
      await act(async () => {
        h.disconnectEvents(permission);
        await new Promise((resolve) => setTimeout(resolve, 1100));
      });
    }
    expect(h.view.getByTestId("editor").getAttribute("data-readonly")).toBe("true");
    expect(h.view.getByRole("alert").textContent).toContain("being edited elsewhere");
    await act(async () => { await expect(h.loaded.actions.save!()).rejects.toThrow("notAllowed"); });
    expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
    await act(async () => h.snapshotEvent(true, null));
    expect(h.view.getByTestId("editor").getAttribute("data-readonly")).toBe("false");
    expect(h.view.queryByRole("alert")).toBeNull();
  } finally { warning.mockRestore(); }
});

test("a writable session event keeps credential viewers read-only", async () => {
  const h = await mount(false, "docx", { canWrite: false, credentialCanWrite: false, readOnlyRelay: true });
  await act(async () => h.writableEvent(false));
  expect(h.view.getByTestId("editor").getAttribute("data-readonly")).toBe("true");
  await act(async () => { await h.loaded.actions.close!(); });
  expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
  expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
});

test.each(["docx", "pptx"] as const)("%s DontTerminateEdit waits for the worker commit after compositionend", async (format) => {
  const h = await mount(false, format);
  const local = new Doc();
  const input = h.view.getByTestId("editor");
  let release!: () => void;
  const worker = new Promise<void>((resolve) => { release = resolve; });
  let commit = Promise.resolve();
  let draft = "unfinished draft";
  const flushing = mock(async (options?: { terminateEdit?: boolean }) => {
    if (options?.terminateEdit !== false) draft = "";
    await commit;
  });
  flush = flushing;
  input.tabIndex = 0;
  input.focus();
  input.addEventListener("compositionend", () => {
    commit = worker.then(() => {
      local.getText("body").insert(0, "漢字");
      expect(publish(encodeStateAsUpdate(local))).toBe(true);
      mounts[0]!.onChange();
    });
  });
  try {
    await act(async () => {
      input.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      h.loaded.bridge.handleMessage({ origin: "https://host.example", data: {
        MessageId: "Action_Save", Values: { Notify: true, DontTerminateEdit: true },
      } });
    });
    expect(flushing).not.toHaveBeenCalled();
    await act(async () => {
      input.dispatchEvent(new Event("compositionend", { bubbles: true }));
      await new Promise((resolve) => setTimeout(resolve, 60));
    });
    expect(flushing).toHaveBeenCalledTimes(1);
    expect(draft).toBe("unfinished draft");
    expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
    expect(h.messages.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([]);
    expect(document.activeElement).toBe(input);
    await act(async () => { release(); await commit; });
    expect(h.hostText).toBe("漢字");
    expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(1);
    expect(h.messages.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([
      { MessageId: "Action_Save_Resp", SendTime: expect.any(Number), Values: { success: true } },
    ]);
  } finally { release(); await commit; local.destroy(); }
});

test.each(compositionSaves)("%s %s waits for compositionend and saves the composed edit", async (format, action) => {
  const h = await mount(false, format);
  const local = new Doc();
  const input = h.view.getByTestId("editor");
  const flushing = mock(async () => {});
  flush = flushing;
  input.addEventListener("compositionstart", (event) => event.stopPropagation());
  input.addEventListener("compositionend", (event) => {
    event.stopPropagation();
    local.getText("body").insert(0, "漢字");
    expect(publish(encodeStateAsUpdate(local))).toBe(true);
    mounts[0]!.onChange();
  });
  try {
    let saving: Promise<unknown> | undefined;
    await act(async () => {
      input.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      if (action === "Action_Save") {
        h.loaded.bridge.handleMessage({ origin: "https://host.example", data: { MessageId: action, Values: { Notify: true } } });
      } else {
        saving = action === "Close" ? h.loaded.actions.close!() : Promise.resolve(mounts[0]!.onSaveRequest!());
      }
    });
    expect(flushing).not.toHaveBeenCalled();
    expect(h.calls.filter((path) => path.endsWith("/save") || path.endsWith("/close"))).toEqual([]);
    expect(h.messages.filter((message) => message.MessageId === "Action_Save_Resp" || message.MessageId === "UI_Close")).toEqual([]);
    expect(local.getText("body").toString()).toBe("");
    await act(async () => {
      input.dispatchEvent(new Event("compositionend", { bubbles: true }));
      await new Promise((resolve) => setTimeout(resolve, 60));
      await saving;
    });
    expect(flushing).toHaveBeenCalledTimes(1);
    expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(1);
    expect(h.hostText).toBe("漢字");
    if (action === "Action_Save") {
      expect(h.messages.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([
        { MessageId: "Action_Save_Resp", SendTime: expect.any(Number), Values: { success: true } },
      ]);
    } else if (action === "Close") {
      expect(h.calls.filter((path) => path.endsWith("/close"))).toHaveLength(1);
      expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
    }
  } finally { local.destroy(); }
});

test.each(compositionSaves)("%s %s fails with compositionActive at the deadline and can be retried", async (format, action) => {
  const h = await mount(false, format);
  const input = h.view.getByTestId("editor");
  const flushing = mock(async () => {});
  flush = flushing;
  const schedule = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  const deadlines = new Map<ReturnType<typeof setTimeout>, () => void>();
  const timers = spyOn(globalThis, "setTimeout").mockImplementation(((callback: () => void, delay: number) => {
    if (delay < 100 || delay > 3000) return schedule(callback, delay);
    const timer = schedule(() => {}, delay);
    deadlines.set(timer, callback);
    return timer;
  }) as typeof setTimeout);
  const cancellations = spyOn(globalThis, "clearTimeout").mockImplementation(((timer: ReturnType<typeof setTimeout>) => {
    deadlines.delete(timer);
    cancel(timer);
  }) as typeof clearTimeout);
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    let outcome: Promise<unknown> | undefined;
    await act(async () => {
      input.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      if (action === "Action_Save") {
        h.loaded.bridge.handleMessage({ origin: "https://host.example", data: {
          MessageId: action, Values: { Notify: true, DontTerminateEdit: true },
        } });
      } else {
        const saving = action === "Close" ? h.loaded.actions.close!() : Promise.resolve(mounts[0]!.onSaveRequest!());
        outcome = saving.then(() => "saved", (error: Error) => error.message);
      }
    });
    if (action === "Close") expect(h.view.container.querySelector("main")!.inert).toBe(true);
    expect(deadlines.size).toBe(1);
    await act(async () => {
      for (const deadline of [...deadlines.values()]) deadline();
      await outcome;
    });
    expect(flushing).not.toHaveBeenCalled();
    expect(h.calls.filter((path) => path.endsWith("/save") || path.endsWith("/close"))).toEqual([]);
    expect(h.view.getByRole("alert").textContent).toMatch(/finish typing/i);
    expect(h.view.getByTestId("editor")).toBe(input);
    if (action === "Close") expect(h.view.container.querySelector("main")!.inert).toBe(false);
    else expect(h.view.container.querySelector("main")!.inert).not.toBe(true);
    expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toEqual([]);
    if (action === "Action_Save") {
      expect(h.messages.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([
        { MessageId: "Action_Save_Resp", SendTime: expect.any(Number), Values: { success: false, result: "compositionActive" } },
      ]);
    } else if (action === "Save button") {
      expect(await outcome).toBe("compositionActive");
    }
    await act(async () => input.dispatchEvent(new Event("compositionend", { bubbles: true })));
    expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
    await act(async () => {
      if (action === "Close") await h.loaded.actions.close!();
      else await h.loaded.actions.save!();
    });
    expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(1);
    expect(h.view.queryByRole("alert")).toBeNull();
    if (action === "Close") expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
  } finally {
    for (const timer of deadlines.keys()) cancel(timer);
    timers.mockRestore(); cancellations.mockRestore(); warning.mockRestore();
  }
});

test.each(["remount", "teardown"])("%s cancels composition waiting and clears the old editor's composition", async (action) => {
  const h = await mount();
  const oldInput = h.view.getByTestId("editor");
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    let outcome!: Promise<string>;
    await act(async () => {
      oldInput.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      outcome = h.loaded.actions.save!().then(() => "saved", (error: Error) => error.message);
    });
    expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
    await act(async () => {
      if (action === "remount") sockets[0]!.drop(4409);
      else h.view.unmount();
      await outcome;
    });
    expect(await outcome).toBe(action === "remount" ? "rejoinRequired" : "editorDisposed");
    if (action === "remount") {
      expect(mounts).toHaveLength(1);
      await act(async () => { oldInput.dispatchEvent(new Event("compositionend", { bubbles: true })); });
      for (let i = 0; i < 40 && mounts.length < 2; i += 1) {
        await act(async () => { await new Promise((resolve) => setTimeout(resolve, 25)); });
      }
      expect(mounts).toHaveLength(2);
      oldInput.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      await act(async () => { sockets[1]!.open(); await h.loaded.actions.save!(); });
      expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(1);
    } else {
      oldInput.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      expect(h.calls.filter((path) => path.endsWith("/save"))).toEqual([]);
    }
  } finally { warning.mockRestore(); }
});

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
  const pendingFlush = new Promise<void>((resolve) => { release = resolve; });
  flush = () => pendingFlush;
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

test.each([1006])("close code %i reconnects the existing editor without refetching", async (code) => {
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

test("a stalled Close completes after timeout and a second Close is a no-op", async () => {
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
    const container = h.view.container.querySelector("main")!;
    expect(container.inert).toBe(true);
    expect(deadlines.size).toBe(1);
    await act(async () => { for (const deadline of [...deadlines.values()]) deadline(); await first; });
    expect(signals).toHaveLength(1);
    expect(signals[0]?.aborted).toBe(true);
    expect(container.inert).toBe(true);
    expect(h.view.queryByRole("alert")).toBeNull();
    expect(h.view.queryByTestId("editor")).toBeNull();
    expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
    await act(async () => { await h.loaded.actions.close!(); });
    expect(signals).toHaveLength(1);
    expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(1);
    expect(h.view.queryByTestId("editor")).toBeNull();
    expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
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
    return Response.json({ format: "docx", initialUpdate: "Aw==", baselineGeneration: 2, canWrite: true });
  });
  const { loadBaseline } = await import("./page");
  const baseline = await loadBaseline(api);
  expect(calls).toEqual(["/api/session/source", "/api/session/join", "/api/session/source", "/api/session/join"]);
  expect(prepareFile).toHaveBeenCalledTimes(1);
  expect(baseline.file.bytes).toEqual(Uint8Array.of(2));
  expect(baseline.join.baselineGeneration).toBe(2);
});

test.each(["unauthorized", "rejected upgrade", "closed event"])("an ended session blocks input and stops retries after %s", async (reason) => {
  const h = await mount();
  await act(async () => {
    if (reason === "closed event") {
      h.endSession();
      h.closedEvent();
    } else if (reason === "unauthorized") sockets[0]!.drop(4401);
    else {
      sockets[0]!.drop(1006);
      await new Promise((resolve) => setTimeout(resolve, 300));
      h.endSession();
      sockets[1]!.drop(1006);
    }
  });
  const count = sockets.length;
  await act(async () => { await new Promise((resolve) => setTimeout(resolve, 600)); });
  expect(sockets).toHaveLength(count);
  expect(h.view.container.querySelector("main")!.inert).toBe(true);
  expect(h.view.getByRole("alert").textContent).toMatch(/session.*ended.*reload/i);
  expect(mounts).toHaveLength(1);
});

test.each(["provider", "unacknowledged", "none"])("expired reconnect preserves %s local work as a download", async (pending) => {
  const h = await mount();
  const local = new Doc();
  const recovered = new Doc();
  let blob: Blob | undefined;
  let downloadName = "";
  const createUrl = spyOn(URL, "createObjectURL").mockImplementation((value) => {
    blob = value as Blob;
    return "blob:recovery";
  });
  const revokeUrl = spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
  const click = spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
    downloadName = this.download;
  });
  try {
    await act(async () => {
      sockets[0]!.drop(1006);
      if (pending !== "none") local.getText("body").insert(0, "edit made offline");
      serialize = async () => encodeStateAsUpdate(local);
      if (pending === "provider") hasPendingUpdates = () => true;
      if (pending === "unacknowledged") mounts[0]!.onLocalUpdate!();
      expect(publish(encodeStateAsUpdate(local))).toBe(false);
      await new Promise((resolve) => setTimeout(resolve, 300));
      h.endSession();
      sockets[1]!.drop(1006);
    });
    expect(h.view.container.querySelector("main")!.inert).toBe(true);
    expect(mounts).toHaveLength(1);
    expect(createUrl).not.toHaveBeenCalled();
    if (pending === "none") {
      expect(h.view.queryByRole("button", { name: "Download recovered document" })).toBeNull();
      expect(h.view.getByRole("alert").textContent).toMatch(/session.*ended.*reload/i);
      return;
    }
    expect(h.view.getByRole("alert").textContent).toMatch(/download.*before.*reload/i);
    act(() => h.view.getByRole("button", { name: "Download recovered document" }).click());
    expect(downloadName).toMatch(/^file \(recovered .+\)\.docx$/);
    applyUpdate(recovered, new Uint8Array(await blob!.arrayBuffer()));
    expect(recovered.getText("body").toString()).toBe("edit made offline");
    expect(h.copies).toHaveLength(0);
    expect(lifecycle).not.toContain("destroy");
    expect(h.view.getByRole("button", { name: "Download recovered document" })).toBeDefined();
  } finally {
    local.destroy(); recovered.destroy();
    createUrl.mockRestore(); revokeUrl.mockRestore(); click.mockRestore();
  }
});

test.each((["docx", "xlsx", "pptx"] as const).flatMap((format) =>
  [false, true].map((closing) => [format, closing] as const),
))("%s baseline recovery waits for composition and preserves its worker commit after Close=%s", async (format, closing) => {
  const h = await mount(false, format);
  const input = h.view.getByTestId("editor");
  const local = new Doc();
  const recovered = new Doc();
  let release!: () => void;
  const worker = new Promise<void>((resolve) => { release = resolve; });
  let commit = Promise.resolve();
  const flushing = mock(async () => { await commit; });
  const warning = spyOn(console, "error").mockImplementation(() => {});
  flush = flushing;
  serialize = async () => encodeStateAsUpdate(local);
  input.addEventListener("compositionend", () => {
    if (input.dataset.readonly === "true" || h.view.container.querySelector("main")!.inert) return;
    commit = worker.then(() => {
      local.getText("body").insert(0, "漢字");
      mounts[0]!.onLocalUpdate!();
    });
  });
  try {
    await act(async () => {
      input.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      if (closing) void h.loaded.actions.close!();
      sockets[0]!.drop(4409);
    });
    expect(h.view.getByTestId("editor")).toBe(input);
    expect(input.dataset.readonly).toBe("false");
    expect(h.view.container.querySelector("main")!.inert).not.toBe(true);
    expect(flushing).not.toHaveBeenCalled();
    expect(h.copies).toHaveLength(0);
    expect(lifecycle).not.toContain("destroy");
    await act(async () => {
      input.dispatchEvent(new Event("compositionend", { bubbles: true }));
      await new Promise((resolve) => setTimeout(resolve, 60));
    });
    expect(flushing).toHaveBeenCalledTimes(1);
    expect(h.view.container.querySelector("main")!.inert).toBe(true);
    expect(h.copies).toHaveLength(0);
    expect(lifecycle).not.toContain("destroy");
    await act(async () => { release(); await commit; });
    expect(h.copies).toHaveLength(1);
    applyUpdate(recovered, h.copies[0]!);
    expect(recovered.getText("body").toString()).toBe("漢字");
    expect(mounts).toHaveLength(2);
    expect(h.calls.filter((path) => path.endsWith("/save") || path.endsWith("/close"))).toEqual([]);
    expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toEqual([]);
  } finally { release(); await commit; local.destroy(); recovered.destroy(); warning.mockRestore(); }
});

test.each([false, true])("baseline recovery retains unfinished composition after its deadline and Retry with Close=%s", async (closing) => {
  const h = await mount(false, "pptx");
  const input = h.view.getByTestId("editor");
  const local = new Doc();
  const recovered = new Doc();
  const flushing = mock(async () => {});
  const warning = spyOn(console, "error").mockImplementation(() => {});
  flush = flushing;
  serialize = async () => encodeStateAsUpdate(local);
  input.addEventListener("compositionend", () => {
    if (input.dataset.readonly === "true" || h.view.container.querySelector("main")!.inert) return;
    local.getText("body").insert(0, "漢字");
    mounts[0]!.onLocalUpdate!();
  });
  const schedule = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  const deadlines = new Map<ReturnType<typeof setTimeout>, () => void>();
  const timers = spyOn(globalThis, "setTimeout").mockImplementation(((callback: () => void, delay: number) => {
    if (delay < 100 || delay > 3000) return schedule(callback, delay);
    const timer = schedule(() => {}, delay);
    deadlines.set(timer, callback);
    return timer;
  }) as typeof setTimeout);
  const cancellations = spyOn(globalThis, "clearTimeout").mockImplementation(((timer: ReturnType<typeof setTimeout>) => {
    deadlines.delete(timer);
    cancel(timer);
  }) as typeof clearTimeout);
  try {
    await act(async () => {
      input.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      if (closing) void h.loaded.actions.close!();
      sockets[0]!.drop(4409);
    });
    expect(deadlines.size).toBe(1);
    await act(async () => { for (const deadline of [...deadlines.values()]) deadline(); });
    expect(h.view.getByRole("alert").textContent).toMatch(/composition.*timed out/i);
    expect(h.view.getByTestId("editor")).toBe(input);
    expect(input.dataset.readonly).toBe("false");
    expect(h.view.container.querySelector("main")!.inert).not.toBe(true);
    expect(flushing).not.toHaveBeenCalled();
    expect(h.copies).toHaveLength(0);
    expect(h.calls.filter((path) => path.endsWith("/source"))).toHaveLength(0);
    expect(lifecycle).not.toContain("destroy");
    await act(async () => h.view.getByRole("button", { name: "Retry" }).click());
    expect(input.dataset.readonly).toBe("false");
    expect(h.view.container.querySelector("main")!.inert).not.toBe(true);
    expect(flushing).not.toHaveBeenCalled();
    await act(async () => {
      input.dispatchEvent(new Event("compositionend", { bubbles: true }));
      await new Promise((resolve) => setTimeout(resolve, 60));
    });
    expect(h.copies).toHaveLength(1);
    applyUpdate(recovered, h.copies[0]!);
    expect(recovered.getText("body").toString()).toBe("漢字");
    expect(mounts).toHaveLength(2);
  } finally {
    for (const timer of deadlines.keys()) cancel(timer);
    timers.mockRestore(); cancellations.mockRestore(); warning.mockRestore();
    local.destroy(); recovered.destroy();
  }
});

test.each(["flush", "serialization", "upload"])("a stalled recovery %s times out and can be retried", async (stage) => {
  const h = await mount();
  hasPendingUpdates = () => true;
  const signals: Array<AbortSignal | null | undefined> = [];
  if (stage === "flush") flush = () => new Promise<void>(() => {});
  if (stage === "serialization") serialize = () => new Promise<Uint8Array>(() => {});
  if (stage === "upload") h.setRecoveryRequest((signal) => {
    signals.push(signal);
    return new Promise<Response>(() => {});
  });
  const schedule = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  const deadlines = new Map<ReturnType<typeof setTimeout>, () => void>();
  const timers = spyOn(globalThis, "setTimeout").mockImplementation(((callback: () => void, delay: number) => {
    if ((delay < 100 || delay > 3000) && delay !== 10_000) return schedule(callback, delay);
    const timer = schedule(() => {}, delay);
    deadlines.set(timer, callback);
    return timer;
  }) as typeof setTimeout);
  const cancellations = spyOn(globalThis, "clearTimeout").mockImplementation(((timer: ReturnType<typeof setTimeout>) => {
    deadlines.delete(timer);
    cancel(timer);
  }) as typeof clearTimeout);
  try {
    await act(async () => sockets[0]!.drop(4409));
    expect(deadlines.size).toBe(1);
    await expect(h.loaded.actions.save!()).rejects.toThrow(/recovery/i);
    await expect(h.loaded.actions.close!()).rejects.toThrow(/recovery/i);
    await act(async () => { for (const deadline of [...deadlines.values()]) deadline(); });
    expect(h.view.getByRole("alert").textContent).toMatch(/recovery.*timed out/i);
    expect(h.view.container.querySelector("main")!.inert).toBe(true);
    expect(mounts).toHaveLength(1);
    expect(lifecycle).not.toContain("destroy");
    expect(h.calls.filter((path) => path.endsWith("/source"))).toHaveLength(0);
    if (stage === "upload") expect(signals[0]?.aborted).toBe(true);
    flush = async () => {};
    serialize = async () => Uint8Array.of(7);
    h.setRecoveryRequest(async () => Response.json({ copyName: "file (retried).docx" }));
    await act(async () => h.view.getByRole("button", { name: "Retry" }).click());
    expect(mounts).toHaveLength(2);
    expect(h.copies.at(-1)).toEqual(Uint8Array.of(7));
    expect(h.view.getByRole("alert").textContent).toContain("file (retried).docx");
  } finally {
    for (const timer of deadlines.keys()) cancel(timer);
    timers.mockRestore(); cancellations.mockRestore();
  }
});

test("an ended handle during recovery offers the serialized document as a local download", async () => {
  const h = await mount();
  hasPendingUpdates = () => true;
  h.setRecoveryRequest(async () => Response.json({ error: { kind: "unauthorized" } }, { status: 401 }));
  await act(async () => sockets[0]!.drop(4409));
  expect(h.view.getByRole("button", { name: "Download recovered document" })).toBeDefined();
  expect(h.view.container.querySelector("main")!.inert).toBe(true);
  expect(mounts).toHaveLength(1);
  expect(lifecycle).not.toContain("destroy");
  expect(h.view.queryByRole("button", { name: "Retry" })).toBeNull();
});

test("a rejected save detects an ended handle and preserves unacknowledged work", async () => {
  const h = await mount();
  const local = new Doc();
  local.getText("body").insert(0, "unacknowledged edit");
  serialize = async () => encodeStateAsUpdate(local);
  mounts[0]!.onLocalUpdate!();
  const save = spyOn(h.loaded.api, "save").mockImplementation(() =>
    new SessionApi("expired", async () => new Response(null, { status: 401 }))
      .save({ connectionId: crypto.randomUUID(), seq: 1 }));
  const warning = spyOn(console, "error").mockImplementation(() => {});
  try {
    await act(async () => { await expect(h.loaded.actions.save!()).rejects.toThrow("sessionEnded"); });
    expect(h.view.getByRole("button", { name: "Download recovered document" })).toBeDefined();
    expect(h.view.container.querySelector("main")!.inert).toBe(true);
    expect(lifecycle).not.toContain("destroy");
  } finally { local.destroy(); save.mockRestore(); warning.mockRestore(); }
});

test("own Close completes on handle invalidation while the close response is delayed", async () => {
  const h = await mount();
  let invalidate!: () => void;
  const requested = new Promise<void>((resolve) => { invalidate = resolve; });
  h.setCloseRequest(() => { invalidate(); return new Promise<Response>(() => {}); });
  let closing!: Promise<void>;
  await act(async () => { closing = h.loaded.actions.close!(); await requested; });
  const container = h.view.container.querySelector("main")!;
  expect(container.inert).toBe(true);
  await act(async () => { sockets[0]!.drop(4401); await closing; });
  expect(container.inert).toBe(true);
  expect(h.view.queryByTestId("editor")).toBeNull();
  expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
  await act(async () => { await h.loaded.actions.close!(); });
  expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(1);
  expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
});

test.each(["socket", "event"])("%s baseline rejoin uploads pending offline work once before destroying the editor", async (reason) => {
  const h = await mount();
  const local = new Doc();
  local.getText("body").insert(0, "offline edit");
  mounts[0]!.onLocalUpdate!();
  serialize = async () => { lifecycle.push("serialize"); return encodeStateAsUpdate(local); };
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  flush = () => gate;
  await act(async () => {
    if (reason === "socket") sockets[0]!.drop(4409);
    else h.event();
  });
  expect(h.view.getByTestId("editor")).toBeDefined();
  expect(h.view.container.querySelector("main")!.inert).toBe(true);
  expect(lifecycle).not.toContain("destroy");
  await act(async () => { sockets[0]!.drop(4409); release(); });
  expect(h.copies).toHaveLength(1);
  const recovered = new Doc();
  try {
    applyUpdate(recovered, h.copies[0]!);
    expect(recovered.getText("body").toString()).toBe("offline edit");
    expect(lifecycle.indexOf("serialize")).toBeLessThan(lifecycle.indexOf("destroy"));
    expect(mounts).toHaveLength(2);
    expect(h.view.getByRole("alert").textContent).toContain("file (conflict recovered).docx");
  } finally { local.destroy(); recovered.destroy(); }
});

test("baseline rejoin with no pending updates does not upload a copy", async () => {
  const h = await mount();
  await act(async () => sockets[0]!.drop(4409));
  expect(h.copies).toHaveLength(0);
  expect(mounts).toHaveLength(2);
});

test("failed recovery upload preserves the old editor read-only", async () => {
  const h = await mount();
  h.failRecovery();
  hasPendingUpdates = () => true;
  await act(async () => sockets[0]!.drop(4409));
  expect(h.copies).toHaveLength(1);
  expect(mounts).toHaveLength(1);
  expect(lifecycle).not.toContain("destroy");
  expect(h.view.getByTestId("editor")).toBeDefined();
  expect(h.view.container.querySelector("main")!.inert).toBe(true);
  expect(h.view.getByRole("alert").textContent).toMatch(/recovery.*502/i);
  expect(h.calls.filter((path) => path.endsWith("/source"))).toHaveLength(0);
});

test("a sent but unacknowledged edit is recovered even with an empty provider queue", async () => {
  const h = await mount();
  const local = new Doc();
  local.getText("body").insert(0, "sent edit");
  serialize = async () => encodeStateAsUpdate(local);
  acknowledge = false;
  try {
    expect(publish(encodeStateAsUpdate(local))).toBe(true);
    await act(async () => sockets[0]!.drop(4409));
    expect(h.copies).toHaveLength(1);
    expect(mounts).toHaveLength(2);
  } finally { local.destroy(); }
});

test("an unacknowledged initial resync without local edits does not upload a copy", async () => {
  acknowledge = false;
  const h = await mount();
  await act(async () => sockets[0]!.drop(4409));
  expect(h.copies).toHaveLength(0);
  expect(mounts).toHaveLength(2);
});

test("an invalid handle response after the final save completes Close", async () => {
  const h = await mount();
  h.setCloseRequest(async () => new Response(null, { status: 401 }));
  await act(async () => { await h.loaded.actions.close!(); });
  expect(h.calls.filter((path) => path.endsWith("/save"))).toHaveLength(1);
  expect(h.view.queryByTestId("editor")).toBeNull();
  expect(h.messages.filter((message) => message.MessageId === "UI_Close")).toHaveLength(1);
});

test("host messages during deferred boot receive normal replies through the window listener", async () => {
  const previousRoot = document.getElementById("root");
  const element = document.createElement("script");
  element.id = "wopi-config";
  element.textContent = JSON.stringify({ sessionHandle: "boot-handle", format: "docx", fileName: "file.docx",
    postMessageOrigin: "https://host.example", userFriendlyName: "Peer" });
  document.body.appendChild(element);
  let release!: (value: { bytes: Uint8Array; baselineGeneration: number }) => void;
  const source = spyOn(SessionApi.prototype, "source").mockImplementation(() => new Promise((resolve) => { release = resolve; }));
  const join = spyOn(SessionApi.prototype, "join").mockResolvedValue({ format: "docx", initialUpdate: Uint8Array.of(0), baselineGeneration: 1, canWrite: true });
  const root = { render: mock(() => {}), unmount() {} };
  const createRoot = spyOn(ReactDOMClient, "createRoot").mockReturnValue(root);
  const send = (MessageId: string, Values = {}) => window.dispatchEvent(new MessageEvent("message", {
    origin: "https://host.example", data: JSON.stringify({ MessageId, Values }),
  }));
  const messages: Array<{ MessageId: string; Values: unknown }> = [];
  const post = spyOn(window.parent, "postMessage").mockImplementation((data: unknown) => {
    const message = JSON.parse(String(data));
    messages.push(message);
    if (message.Values.Status === "Frame_Ready") {
      send("Action_Save", { Notify: true });
      send("Get_Views");
    }
  });
  let bridge: PostMessageBridge | undefined;
  const frameReady = PostMessageBridge.prototype.frameReady;
  const ready = spyOn(PostMessageBridge.prototype, "frameReady").mockImplementation(function (this: PostMessageBridge) {
    bridge = this;
    frameReady.call(this);
  });
  try {
    await act(async () => { await import("./main"); });
    expect(source).toHaveBeenCalledTimes(1);
    expect(root.render).not.toHaveBeenCalled();
    send("Action_Save", { Notify: true });
    send("Get_Views");
    await Promise.resolve();
    expect(messages.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([
      expect.objectContaining({ Values: { success: false, result: "editorNotReady" } }),
      expect.objectContaining({ Values: { success: false, result: "editorNotReady" } }),
    ]);
    expect(messages.filter((message) => message.MessageId === "Get_Views_Resp")).toEqual([
      expect.objectContaining({ Values: [{ ViewId: 0, UserName: "Peer" }] }),
      expect.objectContaining({ Values: [{ ViewId: 0, UserName: "Peer" }] }),
    ]);
  } finally {
    release?.({ bytes: Uint8Array.of(2), baselineGeneration: 1 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 0)); });
    bridge?.dispose(); element.remove();
    if (!previousRoot) document.getElementById("root")?.remove();
    source.mockRestore(); join.mockRestore(); createRoot.mockRestore(); post.mockRestore(); ready.mockRestore();
  }
});
