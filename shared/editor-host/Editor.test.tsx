import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, expect, mock, spyOn, test } from "bun:test";
import { forwardRef, useEffect, useImperativeHandle } from "react";
import type { XlsxAsyncEditorApi, XlsxAsyncEditorProps, XlsxCommandResult } from "@betteroffice/xlsx-react";
import type { PptxAsyncEditorApi, PptxAsyncEditorProps } from "@betteroffice/pptx-react";
import type { DocxErrorInfo } from "@betteroffice/docx-react";
import type { AsyncReplicaAdapter } from "@betteroffice/docx/collaboration";
import type { CollaborationReplica } from "./collaboration/types";
import type { EditorApi } from "./Editor";
import { Doc } from "yjs";
import { SessionApi } from "../../apps/wopi/web/src/api";
import { parseWopiConfig } from "../../apps/wopi/web/src/config";
import type { Loaded } from "../../apps/wopi/web/src/page";
import { PostMessageBridge } from "../../apps/wopi/web/src/postMessage";
import * as wopiTransport from "../../apps/wopi/web/src/transport";

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

interface DocxProps {
  documentBuffer: Uint8Array;
  onFirstPagePainted(): void;
  onSave(buffer: ArrayBuffer): void;
  onSaveRequest?: () => boolean | void | Promise<boolean | void>;
  onError(error: Error, info?: DocxErrorInfo): void;
  mode?: "editing" | "viewing";
  readOnly?: boolean;
  legacySyncApi?: boolean;
  onChange?: () => void;
  onDocumentChange?: () => void;
  collaboration?: { onReplica?(replica: CollaborationReplica | null): void };
}
let docxProps: DocxProps;
let docxReplica: CollaborationReplica | null = null;
let deferDocxPaint = false;
const parseDocx = mock(async () => ({}));
const flushDocxInput = mock(async () => {});
const flushWorkbookInput = mock(async () => "version");
const flushPresentationInput = mock(async () => {});
let hydrateWorkbook = async () => {};
let hydratePresentation = async () => {};
let save: (props: DocxProps) => Promise<ArrayBuffer | null> = async () =>
  null;
const workerOpens: Uint8Array[] = [];
let workersDestroyed = 0;
class ResidentEngineWorkerClient {
  async openSource({ bytes }: { bytes: Uint8Array }) {
    workerOpens.push(bytes);
    return {};
  }
  destroy() {
    workersDestroyed += 1;
  }
}
let workbookProps: XlsxAsyncEditorProps;
let workbookSave: (props: XlsxAsyncEditorProps) => Promise<XlsxCommandResult>;
let deferWorkbookReady = false;
let readyWorkbook: (() => void) | undefined;
let disposeWorkbookApi: (() => void) | undefined;
const workbookReplica = {
  clientId: 7,
  encodeStateVector: () => Uint8Array.of(0),
  encodeStateAsUpdate: () => Uint8Array.of(0, 0),
  applyUpdate() {},
  onUpdate: () => () => {},
};
mock.module("@betteroffice/xlsx-react", () => ({
  XlsxEditor: (props: XlsxAsyncEditorProps) => {
    workbookProps = props;
    useEffect(() => {
      let dispose: void | (() => void);
      readyWorkbook = () => {
        dispose = props.onReady?.({
          focus() {},
          whenHydrated: () => hydrateWorkbook(),
          version: flushWorkbookInput,
          collaboration: { encodeStateVector: async () => Uint8Array.of(8) },
          commands: { execute: () => workbookSave(workbookProps) },
        } as unknown as XlsxAsyncEditorApi);
        disposeWorkbookApi = typeof dispose === "function" ? dispose : undefined;
      };
      if (!deferWorkbookReady) readyWorkbook();
      return () => { if (typeof dispose === "function") dispose(); };
    }, []);
    useEffect(() => {
      props.collaboration?.onReplica?.(workbookReplica);
      return () => props.collaboration?.onReplica?.(null);
    }, [props.collaboration?.onReplica]);
    return <input data-testid="xlsx-input" />;
  },
}));
const disposeWorkbook = mock(async () => {});
const openSession = mock(async (_bytes: Uint8Array) => ({ dispose: disposeWorkbook }));
const initWorkbook = mock(async () => {});
mock.module("@betteroffice/xlsx", () => ({ openWorkbookSession: openSession, initWasm: initWorkbook }));
const presentationReplica = {
  clientId: 7,
  encodeStateVector: async () => Uint8Array.of(1),
  encodeStateAsUpdate: async () => Uint8Array.of(2),
  applyUpdate: async () => {},
  onUpdate: () => () => {},
};
const providerReplicas: unknown[] = [];
let presentationProps: PptxAsyncEditorProps;
const presentationOptions: PptxAsyncEditorProps['collaboration'][] = [];
let savePresentation: () => Promise<Uint8Array> = async () => Uint8Array.of(3);
let deferPresentationReady = false;
let readyPresentation: (() => void) | undefined;
let disposePresentation: (() => void) | undefined;
mock.module("@betteroffice/pptx-react", () => ({
  PptxEditor: (props: PptxAsyncEditorProps) => {
    presentationProps = props;
    presentationOptions.push(props.collaboration);
    useEffect(() => {
      let dispose: void | (() => void);
      readyPresentation = () => {
        dispose = props.onReady?.({ saveAsync: () => savePresentation(), focus() {}, flushPendingInput: flushPresentationInput, whenHydrated: () => hydratePresentation() } as unknown as PptxAsyncEditorApi);
        disposePresentation = typeof dispose === "function" ? dispose : undefined;
      };
      if (!deferPresentationReady) readyPresentation();
      return () => { if (typeof dispose === "function") dispose(); };
    }, [props.onReady]);
    useEffect(() => {
      props.collaboration?.onReplica?.(presentationReplica);
      return () => props.collaboration?.onReplica?.(null);
    }, [props.collaboration?.onReplica]);
    return null;
  },
}));
let providerPendingBytes = 0;
class Provider {
  get pendingBytes() { return providerPendingBytes; }
  constructor(replica: unknown) { providerReplicas.push(replica); }
  onStatus() { return () => {}; }
  onError() { return () => {}; }
  connect() {}
  destroy() {}
}

test.each(["docx", "xlsx", "pptx"] as const)("%s save readiness reads the live provider queue", async (format) => {
  docxReplica = {
    clientId: 42, encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0),
    applyUpdate() {}, onUpdate: () => () => {},
  };
  const { api } = await mount(format, true);
  expect(api.hasPendingUpdates()).toBe(false);
  providerPendingBytes = 7;
  try { expect(api.hasPendingUpdates()).toBe(true); }
  finally { providerPendingBytes = 0; }
});
mock.module("@betteroffice/docx-react", () => ({
  DocxEditor: forwardRef<unknown, DocxProps>(function DocxEditor(props, ref) {
    docxProps = props;
    useImperativeHandle(ref, () => ({ save: () => save(props), flushPendingInput: flushDocxInput }));
    useEffect(() => {
      props.collaboration?.onReplica?.(docxReplica);
      return () => props.collaboration?.onReplica?.(null);
    }, [props.collaboration?.onReplica]);
    useEffect(() => {
      if (deferDocxPaint) return;
      if (props.legacySyncApi) { props.onFirstPagePainted(); return; }
      const client = new ResidentEngineWorkerClient();
      let current = true;
      void client.openSource({ bytes: props.documentBuffer }).then(() => {
        if (current) props.onFirstPagePainted();
      }, (cause: Error) => { if (current) props.onError(cause); });
      return () => { current = false; client.destroy(); };
    }, [props.documentBuffer]);
    return <span data-testid="docx-mode" data-mode={props.mode} data-readonly={String(props.readOnly ?? false)} />;
  }),
}));
const loadFont = mock(async (_face: unknown, _options?: { baseUrl?: string }): Promise<Uint8Array> => Uint8Array.of(1));
const openPresentationSession = mock(async (_bytes: Uint8Array, _options: unknown) => ({ dispose: async () => {} }));
mock.module("@betteroffice/pptx/worker", () => ({ openPresentationSession }));

const roomTransportCalls: [string, string][] = [];
const createRoomTransport = mock((relayOrigin: string, roomId: string) => {
  roomTransportCalls.push([relayOrigin, roomId]);
  return {};
});
mock.module("./collaboration/createRoomTransport", () => ({ createRoomTransport }));

for (const [specifier, exports] of Object.entries({
  "@betteroffice/docx-react/styles.css": {},
  "@betteroffice/docx/docx": { parseDocx },
  "@betteroffice/docx/yrs": { ResidentEngineWorkerClient },
  "@betteroffice/docx/layout": { configureDefaultFonts: () => {} },
  "@betteroffice/docx/utils": { setGoogleFontsEnabled: () => {} },
  "@betteroffice/docx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/xlsx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/pptx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/fonts": {
    loadBundledFontBytes: loadFont,
    resolveLastResortFace: () => {},
    resolveMetricCompatFace: () => {},
  },
}))
  mock.module(specifier, () => exports);

const { act, cleanup, render } = await import("@testing-library/react");
const { default: Editor, prepareFile } = await import("./Editor");
const { default: WopiApp } = await import("../../apps/wopi/web/src/page");

afterEach(() => {
  cleanup();
  docxReplica = null;
  deferDocxPaint = false;
  hydrateWorkbook = async () => {};
  hydratePresentation = async () => {};
  deferWorkbookReady = false;
  readyWorkbook = undefined;
  disposeWorkbookApi = undefined;
  deferPresentationReady = false;
  readyPresentation = undefined;
  disposePresentation = undefined;
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount(format: "docx" | "xlsx" | "pptx" = "docx", collaborative = false, onChange = () => {}, readOnly = false, onLocalUpdate = () => {}, onSaveRequest?: () => Promise<boolean>) {
  const ready: { api?: EditorApi } = {};
  const errors: string[] = [];
  const saved: Uint8Array[] = [];
  let view!: ReturnType<typeof render>;
  await act(async () => {
    view = render(
      <Editor
        file={{ bytes: new Uint8Array(), format }}
        name="saved.docx"
        room={collaborative ? { id: "test", name: "Peer" } : null}
        relayOrigin="https://relay.example.test"
        readOnly={readOnly}
        onSave={(bytes) => saved.push(bytes)}
        onChange={onChange}
        onLocalUpdate={onLocalUpdate}
        onSaveRequest={onSaveRequest}
        onError={(error) => errors.push(error.message)}
        onReady={(api) => {
          if (api) ready.api = api;
        }}
        onOpen={async () => {}}
        onStatus={() => {}}
      />
    );
  });
  return { api: ready.api!, errors, saved, view };
}

function readinessClock() {
  let now = 0;
  const timers = new Map<unknown, { at: number; run(): void }>();
  const schedule = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  const scheduling = spyOn(globalThis, "setTimeout").mockImplementation(((...parameters: Parameters<typeof schedule>) => {
    const [callback, delay, ...args] = parameters;
    if (delay !== 30_000) return schedule(...parameters);
    const timer = schedule(() => {}, 0);
    timers.set(timer, { at: now + delay, run() {
      if (typeof callback !== "function") throw new Error("Expected a timeout callback");
      callback(...args);
    } });
    return timer;
  }) as typeof schedule);
  const cancelling = spyOn(globalThis, "clearTimeout").mockImplementation(((timer: Parameters<typeof cancel>[0]) => {
    timers.delete(timer);
    cancel(timer);
  }) as typeof cancel);
  return {
    advance(milliseconds: number) {
      now += milliseconds;
      for (const [timer, deadline] of timers) {
        if (deadline.at > now) continue;
        timers.delete(timer);
        deadline.run();
      }
    },
    get pending() { return timers.size; },
    restore() { scheduling.mockRestore(); cancelling.mockRestore(); },
  };
}

test("docx input flush reaches the mounted editor barrier", async () => {
  flushDocxInput.mockClear();
  const { api } = await mount();
  expect(api.flushPendingInput).toBeDefined();
  await api.flushPendingInput!();
  expect(flushDocxInput).toHaveBeenCalledTimes(1);
});

test.each(["xlsx", "pptx"] as const)("%s commits pending input through its public API", async (format) => {
  const flush = format === "xlsx" ? flushWorkbookInput : flushPresentationInput;
  flush.mockClear();
  const { api } = await mount(format);
  await api.flushPendingInput();
  expect(flush).toHaveBeenCalledTimes(1);
});

test.each([false, true])("xlsx Action_Save commits deferred text through version with DontTerminateEdit=%s", async (DontTerminateEdit) => {
  const document = new Doc();
  let deferredText = "";
  let release!: () => void;
  const deferred = new Promise<void>((resolve) => { release = resolve; });
  flushWorkbookInput.mockClear();
  flushWorkbookInput.mockImplementationOnce(async () => {
    await deferred;
    document.getText("body").insert(0, deferredText);
    workbookProps.onChange?.();
    return "committed";
  });
  const saved: string[] = [];
  const replies: Array<{ MessageId: string; SendTime?: number; Values: Record<string, unknown> }> = [];
  const barrier = { connectionId: crypto.randomUUID(), seq: 1 };
  const captureSaveBarrier = mock(() => barrier);
  const connection = spyOn(wopiTransport, "createWopiTransport").mockReturnValue({
    connect() {}, disconnect() {}, send: () => true, onEvent: () => () => {},
    hasPendingResync: () => false, hasUnacknowledgedUpdates: () => false, noteLocalUpdate() {},
    captureSaveBarrier,
  });
  const api = new SessionApi("xlsx-handle", async (input, init) => {
    if (String(input).endsWith("/save")) {
      expect(JSON.parse(String(init?.body))).toEqual(barrier);
      saved.push(document.getText("body").toString());
      return Response.json({ itemVersion: "2" });
    }
    if (String(input).endsWith("/events")) {
      return new Response(new ReadableStream({
        start(controller) { init?.signal?.addEventListener("abort", () => controller.close(), { once: true }); },
      }));
    }
    throw new Error(`Unexpected request: ${input}`);
  });
  let saving: Promise<void> | undefined;
  const bridge = new PostMessageBridge({
    parent: { postMessage(message) { replies.push(JSON.parse(message)); } },
    addEventListener() {}, removeEventListener() {},
  }, "https://host.example", {
    onSave(options) {
      saving = loaded.actions.save!(options);
      return saving;
    },
    onClose() {}, onBlurFocus() {}, onGrabFocus() {}, onGetViews: () => [],
  });
  const loaded: Loaded = {
    config: parseWopiConfig({ sessionHandle: "xlsx-handle", format: "xlsx", fileName: "book.xlsx", canWrite: true }),
    api, bridge, file: { bytes: Uint8Array.of(0), format: "xlsx" },
    join: { format: "xlsx", initialUpdate: Uint8Array.of(0, 0), baselineGeneration: 1, canWrite: true },
    clientId: 7, actions: { views: [] },
  };
  let view: ReturnType<typeof render> | undefined;
  try {
    await act(async () => { view = render(<WopiApp loaded={loaded} />); });
    const input = view!.getByTestId("xlsx-input");
    input.addEventListener("compositionend", () => { deferredText = "変"; });
    await act(async () => {
      input.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      input.dispatchEvent(new Event("compositionend", { bubbles: true }));
    });
    expect(document.getText("body").toString()).toBe("");
    await act(async () => bridge.handleMessage({ origin: "https://host.example", data: {
      MessageId: "Action_Save", Values: { Notify: true, DontTerminateEdit },
    } }));
    expect(flushWorkbookInput).toHaveBeenCalledTimes(1);
    expect(captureSaveBarrier).not.toHaveBeenCalled();
    expect(saved).toEqual([]);
    expect(replies.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([]);
    await act(async () => { release(); await saving; });
    expect(saved).toEqual(["変"]);
    expect(replies.filter((message) => message.MessageId === "Action_Save_Resp")).toEqual([
      { MessageId: "Action_Save_Resp", SendTime: expect.any(Number), Values: { success: true } },
    ]);
  } finally {
    release();
    await saving?.catch(() => {});
    await act(async () => view?.unmount());
    connection.mockRestore();
    flushWorkbookInput.mockReset();
    flushWorkbookInput.mockImplementation(async () => "version");
    document.destroy();
  }
});

test.each(["xlsx", "pptx"] as const)("%s readiness waits for opening and hydration", async (format) => {
  let hydrate!: () => void;
  const pending = new Promise<void>((resolve) => { hydrate = resolve; });
  if (format === "xlsx") { deferWorkbookReady = true; hydrateWorkbook = () => pending; }
  else { deferPresentationReady = true; hydratePresentation = () => pending; }
  const { api } = await mount(format);
  let ready = false;
  const waiting = api.whenReady().then(() => { ready = true; });
  await act(async () => {});
  expect(ready).toBe(false);
  await act(async () => { if (format === "xlsx") readyWorkbook!(); else readyPresentation!(); });
  expect(ready).toBe(false);
  hydrate();
  await waiting;
  expect(ready).toBe(true);
});

test("xlsx state vector uses the asynchronous collaboration API", async () => {
  const { api } = await mount("xlsx", true);
  expect(await api.encodeStateVector()).toEqual(Uint8Array.of(8));
});

test("pptx state vector comes from the live replica", async () => {
  const { api } = await mount("pptx", true);
  expect(await api.encodeStateVector()).toEqual(Uint8Array.of(1));
});

test("docx ordered saves read the worker replica through its async API", async () => {
  const vector = Uint8Array.of(1, 42, 3);
  const readVector = mock(async () => vector);
  const syncVector = mock((): Uint8Array => {
    throw new Error("encodeStateVector is not available on a worker-hosted replica");
  });
  const replica: AsyncReplicaAdapter = {
    clientId: 42,
    encodeStateVector: syncVector,
    encodeStateAsUpdate() { throw new Error("worker-hosted replica"); },
    applyUpdate() {},
    onUpdate: () => () => {},
    async: {
      clientId: 42,
      encodeStateVector: readVector,
      encodeStateAsUpdate: async () => Uint8Array.of(0),
      applyUpdate: async () => {},
      onUpdate: () => () => {},
    },
  };
  docxReplica = replica;
  const { api } = await mount("docx", true);
  await api.flushPendingInput();
  expect(await api.encodeStateVector()).toEqual(vector);
  expect(readVector).toHaveBeenCalledTimes(1);
  expect(syncVector).not.toHaveBeenCalled();
});

test("docx ordered saves support a legacy synchronous replica", async () => {
  docxReplica = {
    clientId: 42,
    encodeStateVector: () => Uint8Array.of(1, 42, 3),
    encodeStateAsUpdate: () => Uint8Array.of(0),
    applyUpdate() {},
    onUpdate: () => () => {},
  };
  const { api } = await mount("docx", true);
  expect(await api.encodeStateVector()).toEqual(Uint8Array.of(1, 42, 3));
});

test.each(["", "&peerless=1", "&peerless=0"])("docx %s reports dirty state through document events", async (query) => {
  const previousUrl = location.href;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL(`http://localhost/?format=docx${query}`);
  try {
    const changed = mock(() => {});
    await mount("docx", false, changed);
    expect(docxProps!.legacySyncApi).toBe(query === "&peerless=0");
    expect(docxProps!.onChange).toBeUndefined();
    docxProps!.onDocumentChange?.();
    expect(changed).toHaveBeenCalledTimes(1);
  } finally {
    browser.happyDOM.setURL(previousUrl);
  }
});

test("peerless preview reports dirty state through worker document events", async () => {
  const previousUrl = location.href;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL("http://localhost/?format=docx&peerless=1");
  try {
    const changed = mock(() => {});
    await mount("docx", false, changed);
    expect(docxProps!.legacySyncApi).toBe(false);
    expect(docxProps!.onChange).toBeUndefined();
    docxProps!.onDocumentChange?.();
    expect(changed).toHaveBeenCalledTimes(1);
  } finally {
    browser.happyDOM.setURL(previousUrl);
  }
});

test.each(["", "&peerless=1", "&peerless=0"])("docx %s prepares bytes for its selected runtime", async (query) => {
  const previousUrl = location.href;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL(`http://localhost/?format=docx${query}`);
  parseDocx.mockClear();
  workerOpens.length = 0;
  workersDestroyed = 0;
  try {
    const bytes = Uint8Array.of(1, 2);
    expect(await prepareFile(bytes, "docx")).toEqual({ bytes, format: "docx" });
    expect(parseDocx).toHaveBeenCalledTimes(query === "&peerless=0" ? 1 : 0);
    expect(workerOpens).toEqual([]);
    expect(workersDestroyed).toBe(0);
  } finally {
    browser.happyDOM.setURL(previousUrl);
  }
});

test("recoverable DOCX errors preserve serialization and the mounted editor", async () => {
  const { api, errors, saved, view } = await mount();
  const editor = view.getByTestId("docx-mode");
  const stall = new Error("Full frame responses did not satisfy requested coverage and revision");
  act(() => docxProps.onError(stall, { recoverable: true }));
  save = async ({ onError, onSave }) => {
    onError(stall, { recoverable: true });
    onSave(Uint8Array.of(6).buffer);
    return Uint8Array.of(6).buffer;
  };
  expect(await api.serialize()).toEqual(Uint8Array.of(6));
  save = async ({ onError }) => {
    onError(stall, { recoverable: true });
    return null;
  };
  await expect(api.serialize()).rejects.toThrow("The editor is still opening the file. Try again in a moment.");
  expect(errors).toEqual([]);
  expect(saved).toEqual([]);
  expect(view.getByTestId("docx-mode")).toBe(editor);
});

test("recoverable DOCX errors are logged without reaching the host failure callback", async () => {
  const warn = spyOn(console, "warn").mockImplementation(() => {});
  try {
    const { errors, view } = await mount();
    const editor = view.getByTestId("docx-mode");
    const failure = new Error("Open request failed");
    act(() => docxProps.onError(failure, { recoverable: true }));
    expect(warn.mock.calls).toEqual([[failure]]);
    expect(errors).toEqual([]);
    expect(view.getByTestId("docx-mode")).toBe(editor);
  } finally {
    warn.mockRestore();
  }
});

test("a failed save rejects with the editor's error and a retry saves", async () => {
  const { api, errors, saved } = await mount();
  save = async ({ onError }) => {
    onError(new Error("The document changed while saving"));
    return null;
  };
  await expect(api.serialize()).rejects.toThrow(
    "The document changed while saving"
  );
  save = async ({ onSave }) => {
    onSave(Uint8Array.of(1).buffer);
    return Uint8Array.of(1).buffer;
  };
  expect(await api.serialize()).toEqual(Uint8Array.of(1));
  expect(errors).toEqual(["The document changed while saving"]);
  expect(saved).toEqual([]);
});

test("xlsx serialization awaits the save command and captures bytes without notifying a user save", async () => {
  const { api, saved } = await mount("xlsx");
  expect(workbookProps.legacySyncApi).toBeUndefined();
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  workbookSave = async (props) => {
    await gate;
    props.onSave?.(Uint8Array.of(3));
    return { ok: true, status: "executed" };
  };
  let settled = false;
  const pending = api.serialize().then((bytes) => { settled = true; return bytes; });
  await new Promise((resolve) => setTimeout(resolve, 20));
  expect(settled).toBe(false);
  release();
  expect(await pending).toEqual(Uint8Array.of(3));
  expect(saved).toEqual([]);
  workbookProps.onSave?.(Uint8Array.of(4));
  expect(saved).toEqual([Uint8Array.of(4)]);
});

test("xlsx serialization rejects a refused save command and allows a retry", async () => {
  const { api } = await mount("xlsx");
  workbookSave = async () => ({ ok: false, failure: { code: "input-failed", message: "Entry refused" } });
  await expect(api.serialize()).rejects.toThrow("Entry refused");
  workbookSave = async (props) => {
    props.onSave?.(Uint8Array.of(5));
    return { ok: true, status: "executed" };
  };
  expect(await api.serialize()).toEqual(Uint8Array.of(5));
});

test("xlsx preparation validates in a worker and disposes it without opening local wasm", async () => {
  const descriptors = ["Worker", "OffscreenCanvas"].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  try {
    for (const [key] of descriptors) Object.defineProperty(globalThis, key, { configurable: true, value: class {} });
    const bytes = Uint8Array.of(1);
    expect(await prepareFile(bytes, "xlsx")).toEqual({ bytes, format: "xlsx" });
    expect(openSession).toHaveBeenCalledWith(bytes);
    expect(disposeWorkbook).toHaveBeenCalledTimes(1);
    expect(initWorkbook).not.toHaveBeenCalled();
  } finally {
    for (const [key, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
});

test("overlapping serializations run in turn and keep their own failure", async () => {
  const { api } = await mount();
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const calls: string[] = [];
  save = async ({ onError }) => {
    calls.push("first");
    save = async () => {
      calls.push("second");
      return Uint8Array.of(2).buffer;
    };
    await gate;
    onError(new Error("first failed"));
    return null;
  };
  const first = api.serialize().catch((cause: Error) => cause.message);
  const second = api.serialize();
  await new Promise((resolve) => setTimeout(resolve, 20));
  expect(calls).toEqual(["first"]);
  release();
  expect(await first).toBe("first failed");
  expect(await second).toEqual(Uint8Array.of(2));
  expect(calls).toEqual(["first", "second"]);
});

test("presentation serialization uses the async default and preserves save failures", async () => {
  const { api, saved } = await mount("pptx");
  expect(presentationProps.legacySyncApi).toBeUndefined();
  savePresentation = async () => { throw new Error("Worker stopped"); };
  await expect(api.serialize()).rejects.toThrow("Worker stopped");
  savePresentation = async () => Uint8Array.of(4);
  expect(await api.serialize()).toEqual(Uint8Array.of(4));
  expect(saved).toEqual([]);
});

const asyncFormats = {
  xlsx: {
    disposed: "The workbook editor was disposed.",
    defer() { deferWorkbookReady = true; },
    ready() { readyWorkbook!(); },
    dispose() { disposeWorkbookApi!(); },
    props: () => workbookProps,
    saves(bytes: Uint8Array) {
      const saving = mock(async (props: XlsxAsyncEditorProps) => {
        props.onSave?.(bytes);
        return { ok: true, status: "executed" } as XlsxCommandResult;
      });
      workbookSave = saving;
      return saving;
    },
  },
  pptx: {
    disposed: "The presentation editor was disposed.",
    defer() { deferPresentationReady = true; },
    ready() { readyPresentation!(); },
    dispose() { disposePresentation!(); },
    props: () => presentationProps,
    saves(bytes: Uint8Array) {
      const saving = mock(async () => bytes);
      savePresentation = saving;
      return saving;
    },
  },
};

for (const [format, editor] of Object.entries(asyncFormats)) {
  test(`${format} serialization waits for the editor API before saving`, async () => {
    editor.defer();
    const saving = editor.saves(Uint8Array.of(5));
    const { api, saved } = await mount(format as "xlsx" | "pptx");
    const result = api.serialize().catch((cause: Error) => cause);
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    expect(saving).not.toHaveBeenCalled();
    await act(async () => { editor.ready(); });
    expect(await result).toEqual(Uint8Array.of(5));
    expect(saving).toHaveBeenCalledTimes(1);
    expect(saved).toEqual([]);
  });

  test(`${format} readiness timeout drains queued serializations`, async () => {
    editor.defer();
    const saving = editor.saves(Uint8Array.of(9));
    const { api } = await mount(format as "xlsx" | "pptx");
    const clock = readinessClock();
    const flush = async () => {
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 0));
        await new Promise((resolve) => setTimeout(resolve, 0));
      });
    };
    const opening = "The editor is still opening the file. Try again in a moment.";
    const results: (Uint8Array | string)[] = ["pending", "pending"];
    try {
      void api.serialize().then((bytes) => { results[0] = bytes; }, (cause: Error) => { results[0] = cause.message; });
      void api.serialize().then((bytes) => { results[1] = bytes; }, (cause: Error) => { results[1] = cause.message; });
      await flush();
      clock.advance(29_999);
      await flush();
      expect(results).toEqual(["pending", "pending"]);
      expect(saving).not.toHaveBeenCalled();
      clock.advance(1);
      await flush();
      expect(results).toEqual([opening, opening]);
      expect(clock.pending).toBe(0);
      expect(saving).not.toHaveBeenCalled();
      await act(async () => { editor.ready(); });
      expect(await api.serialize()).toEqual(Uint8Array.of(9));
    } finally {
      clock.restore();
    }
  });

  test(`unmount rejects waiting, queued, and subsequent ${format} saves`, async () => {
    editor.defer();
    const saving = editor.saves(Uint8Array.of(6));
    const { api, view } = await mount(format as "xlsx" | "pptx");
    const first = api.serialize().catch((cause: Error) => cause.message);
    const second = api.serialize().catch((cause: Error) => cause.message);
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    view.unmount();
    expect(await first).toBe(editor.disposed);
    expect(await second).toBe(editor.disposed);
    await expect(api.serialize()).rejects.toThrow(editor.disposed);
    expect(saving).not.toHaveBeenCalled();
  });

  test(`${format} opening errors reject waiting saves and readiness allows retry`, async () => {
    editor.defer();
    const saving = editor.saves(Uint8Array.of(7));
    const { api, errors } = await mount(format as "xlsx" | "pptx");
    const result = api.serialize().catch((cause: Error) => cause);
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    const failure = new Error("Worker failed to open");
    editor.props().onError?.(failure);
    expect(await result).toBe(failure);
    await expect(api.serialize()).rejects.toThrow("Worker failed to open");
    expect(saving).not.toHaveBeenCalled();
    await act(async () => { editor.ready(); });
    expect(await api.serialize()).toEqual(Uint8Array.of(7));
    expect(errors).toEqual(["Worker failed to open"]);
  });

  test(`${format} API disposal rejects saves until a replacement is ready`, async () => {
    const saving = editor.saves(Uint8Array.of(8));
    const { api } = await mount(format as "xlsx" | "pptx");
    editor.dispose();
    await expect(api.serialize()).rejects.toThrow(editor.disposed);
    expect(saving).not.toHaveBeenCalled();
    await act(async () => { editor.ready(); });
    expect(await api.serialize()).toEqual(Uint8Array.of(8));
  });
}

test.each(["xlsx", "pptx"] as const)("%s receives readOnly when there is no room", async (format) => {
  await mount(format, false, () => {}, true);
  expect(format === "xlsx" ? workbookProps.readOnly : presentationProps.readOnly).toBe(true);
});

test.each(["xlsx", "pptx"] as const)("%s receives readOnly in a collaboration room", async (format) => {
  await mount(format, true, () => {}, true);
  const props = format === "xlsx" ? workbookProps : presentationProps;
  expect(props.readOnly).toBe(true);
  expect(props.collaboration).toBeDefined();
});

test("pptx preparation loads fonts from a configured base url", async () => {
  const descriptors = ["Worker", "OffscreenCanvas"].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  loadFont.mockClear();
  try {
    for (const [key] of descriptors) Object.defineProperty(globalThis, key, { configurable: true, value: class {} });
    await prepareFile(Uint8Array.of(1), "pptx", { fontBaseUrl: "/fonts/" });
    expect(loadFont.mock.calls.length).toBeGreaterThan(0);
    for (const call of loadFont.mock.calls) expect(call[1]).toEqual({ baseUrl: "/fonts/" });
  } finally {
    for (const [key, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
});

test("collaboration transport is created from the relay origin prop", async () => {
  roomTransportCalls.length = 0;
  await mount("pptx", true);
  expect(roomTransportCalls).toEqual([["https://relay.example.test", "test"]]);
});

test("createTransport replaces relayOrigin and receives the room id", async () => {
  roomTransportCalls.length = 0;
  presentationOptions.length = 0;
  const createTransport = mock((_roomId: string) => ({
    connect() {},
    disconnect() {},
    send: () => true,
    onEvent: () => () => {},
  }));
  const initialUpdate = Uint8Array.of(9, 8);
  await act(async () => {
    render(
      <Editor
        file={{ bytes: new Uint8Array(), format: "pptx" }}
        name="deck.pptx"
        room={{ id: "wopi-room", name: "Alice A" }}
        createTransport={createTransport}
        initialUpdate={initialUpdate}
        collaborationClientId={4242}
        onSave={() => {}}
        onChange={() => {}}
        onError={() => {}}
        onReady={() => {}}
        onOpen={async () => {}}
        onStatus={() => {}}
      />
    );
  });
  expect(createTransport).toHaveBeenCalledTimes(1);
  expect(createTransport.mock.calls[0]).toEqual(["wopi-room"]);
  expect(roomTransportCalls).toEqual([]);
  const opening = presentationOptions.find((options) => options?.clientId === 4242);
  expect(opening?.initialUpdate).toEqual(initialUpdate);
});

test("presentation collaboration connects the async replica and presence provider", async () => {
  providerReplicas.length = 0;
  presentationOptions.length = 0;
  await mount("pptx", true);
  expect(providerReplicas).toEqual([presentationReplica]);
  expect(await presentationReplica.encodeStateVector()).toEqual(Uint8Array.of(1));
  expect(presentationProps.collaboration?.presence).toBeInstanceOf(Provider);
  const openingOptions = presentationOptions.filter((options) => options && !options.presence);
  expect(openingOptions.length).toBeGreaterThan(1);
  expect(new Set(openingOptions).size).toBe(1);
});

test.each([1, 4])("pptx preparation tolerates %i failed font faces", async (failed) => {
  const descriptors = ["Worker", "OffscreenCanvas"].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  const warning = spyOn(console, "warn").mockImplementation(() => {});
  let calls = 0;
  loadFont.mockImplementation(async () => {
    if (calls++ < failed) throw new Error("Font download failed");
    return Uint8Array.of(1);
  });
  try {
    for (const [key] of descriptors) Object.defineProperty(globalThis, key, { configurable: true, value: class {} });
    const bytes = Uint8Array.of(1);
    const prepared = await prepareFile(bytes, "pptx");
    expect(prepared.fonts).toHaveLength(4 - failed);
    if (failed === 4) expect(prepared.fonts).toEqual([]);
    else expect(prepared.fonts).toEqual([
      { family: "Arial", bold: false, italic: true, bytes: Uint8Array.of(1) },
      { family: "Arial", bold: true, italic: false, bytes: Uint8Array.of(1) },
      { family: "Arial", bold: true, italic: true, bytes: Uint8Array.of(1) },
    ]);
    expect(openPresentationSession).toHaveBeenLastCalledWith(bytes, { fonts: prepared.fonts });
    expect(warning).toHaveBeenCalledTimes(1);
    expect(warning.mock.calls[0]?.[0]).toContain(`${failed} font faces failed to load`);
  } finally {
    warning.mockRestore();
    loadFont.mockImplementation(async () => Uint8Array.of(1));
    for (const [key, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
});

test.each(["", "&peerless=0"])("docx %s waits for its first painted page", async (query) => {
  const browser = window as unknown as GlobalWindow;
  const previous = location.href;
  browser.happyDOM.setURL(`http://localhost/?format=docx${query}`);
  deferDocxPaint = true;
  try {
    const ready: Array<EditorApi | null> = [];
    await act(async () => {
      render(<Editor file={{ bytes: Uint8Array.of(1), format: "docx" }} name="file.docx" room={null}
        onSave={() => {}} onChange={() => {}} onError={() => {}} onReady={(api) => { ready.push(api); }}
        onOpen={async () => {}} onStatus={() => {}} />);
    });
    expect(ready).toEqual([]);
    await act(async () => { docxProps.onFirstPagePainted(); });
    expect(ready.filter(Boolean)).toHaveLength(1);
  } finally {
    browser.happyDOM.setURL(previous);
  }
});

test("local replica updates are reported independently of provider queues and idle document events", async () => {
  let update!: Parameters<CollaborationReplica["onUpdate"]>[0];
  let unsubscribed = false;
  docxReplica = {
    clientId: 42, encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0), applyUpdate() {},
    onUpdate(listener) { update = listener; return () => { unsubscribed = true; }; },
  };
  const local = mock(() => {});
  const h = await mount("docx", true, () => {}, false, local);
  update(Uint8Array.of(1), "remote");
  expect(local).not.toHaveBeenCalled();
  update(Uint8Array.of(2), "local");
  expect(local).toHaveBeenCalledTimes(1);
  h.view.unmount();
  expect(unsubscribed).toBe(true);
});

test("serializing a recovery copy bypasses the host save request", async () => {
  const requested = mock(async () => { throw new Error("session is rejoining"); });
  save = async (props) => {
    if (props.onSaveRequest && await props.onSaveRequest() !== true) return null;
    return Uint8Array.of(7).buffer;
  };
  const h = await mount("docx", false, () => {}, false, () => {}, requested);
  expect(await h.api.serialize()).toEqual(Uint8Array.of(7));
  expect(requested).not.toHaveBeenCalled();
});
