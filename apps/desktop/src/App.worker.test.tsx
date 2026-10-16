import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from "bun:test";
import { decodeDocxHostJson } from "@betteroffice/docx/yrs";
import * as measurement from "@betteroffice/docx/layout";
import { useEffect, useRef, type ComponentProps } from "react";
import type {
  DocxCommandFrame,
  ResidentCommandInitReply,
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from "@betteroffice/docx/yrs";

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
mock.module("@betteroffice/docx-react/styles.css", () => ({}));
const { act, cleanup, fireEvent, render, waitFor } = await import("@testing-library/react");
const editor = await import("./Editor");
const MountedEditor = editor.default;
const readied = mock(() => {});
const readiness = new Map<string, ReturnType<typeof deferred<void>>>();
let delayAcknowledgements = false;
let flushFailure: Error | null = null;
const delayedEdits = new Map<string, { pending: boolean }>();
function readyFor(name: string) {
  let ready = readiness.get(name);
  if (!ready) { ready = deferred<void>(); readiness.set(name, ready); }
  return ready;
}
function DelayedEditor(props: ComponentProps<typeof MountedEditor>) {
  const edits = useRef<{ ack: ReturnType<typeof deferred<void>>; changed: Promise<void> }[]>([]);
  const state = useRef({ pending: false });
  useEffect(() => {
    delayedEdits.set(props.name, state.current);
    const api = {
      serialize: async () => props.file.bytes,
      focus() {},
      print() {},
      async flushPendingInput() {
        if (flushFailure) throw flushFailure;
        for (const edit of edits.current.splice(0)) {
          edit.ack.resolve();
          await edit.changed;
        }
        state.current.pending = false;
      },
    };
    props.onReady(api);
    readied();
    readyFor(props.name).resolve();
    return () => props.onReady(null);
  }, []);
  return <textarea aria-label="Document input" onInput={() => {
    const ack = deferred<void>();
    state.current.pending = true;
    edits.current.push({ ack, changed: ack.promise.then(() => props.onChange()) });
  }} />;
}
mock.module("./Editor", () => ({ ...editor, default: (props: ComponentProps<typeof MountedEditor>) => delayAcknowledgements
  ? <DelayedEditor {...props} /> : (
  <MountedEditor {...props} onReady={(value) => {
    props.onReady(value);
    if (value) { readied(); readyFor(props.name).resolve(); }
  }} />
) }));
const { App } = await import("./App");
const workers: FakeWorker[] = [];
const opens: number[] = [];
const held = new Set<number>();
const restorers: (() => void)[] = [];
const hostJson = JSON.stringify({
  referencedFonts: [],
  envelope: {
    wireVersion: 1,
    document: { package: {
      document: { content: [], sections: [] }, theme: {}, numbering: {}, settings: {},
      fontTable: { fonts: [] }, relationshipEntries: [], mediaEntries: [], chartEntries: [],
    } },
    embeddedFontParts: [],
  },
});

function frame(revision: number, coverage: number): DocxCommandFrame {
  const u32 = (value: number) => {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setUint32(0, value, true);
    return [...bytes];
  };
  const number = (value: number) => {
    const bytes = new Uint8Array(8);
    new DataView(bytes.buffer).setFloat64(0, value, true);
    return [5, ...bytes];
  };
  const keys = ["pageIndex", "width", "height", "primitives"];
  const strings = [...u32(keys.length), ...keys.flatMap((key) => {
    const bytes = new TextEncoder().encode(key);
    return [...u32(bytes.length), ...bytes];
  })];
  const body = [number(0), number(200), number(300), [7, ...u32(0), ...u32(0)]]
    .flatMap((value, index) => [...u32(index), ...value]);
  const payload = [8, ...u32(body.length), ...u32(4), ...body];
  const dataOffset = Math.ceil((128 + strings.length) / 8) * 8;
  const bytes = new Uint8Array(dataOffset + payload.length);
  const view = new DataView(bytes.buffer);
  bytes.set([0x46, 0x44, 0x56, 0x31]);
  view.setUint16(4, 1, true);
  view.setUint16(6, 80, true);
  view.setUint32(8, bytes.length, true);
  view.setUint32(12, 1, true);
  for (const [offset, value] of [[16, 1], [24, 1], [32, revision]] as const)
    view.setBigUint64(offset, BigInt(value), true);
  for (const [offset, value] of [[48, 1], [52, 1], [56, 80], [60, 128],
    [64, strings.length], [68, dataOffset]] as const)
    view.setUint32(offset, value, true);
  bytes[80] = 1;
  view.setBigUint64(88, 1n, true);
  view.setBigUint64(96, BigInt(revision), true);
  view.setUint32(108, dataOffset, true);
  view.setUint32(112, dataOffset, true);
  view.setUint32(116, payload.length, true);
  bytes.set(strings, 128);
  bytes.set(payload, dataOffset);
  return { revision, coversThroughSeq: coverage, payload: {
    frame: bytes.buffer, caret: { frameEpoch: revision, caretRect: null }, layoutRevision: revision,
    stagedPresentation: { frameEpoch: revision, pages: [] },
  } };
}

class FakeWorker {
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;
  terminate = mock(() => {});
  initial?: ResidentCommandInitReply;
  source = 0;
  text = "";
  saves = 0;
  holdSave = false;
  pending: ResidentEngineWorkerResponse[] = [];
  constructor() { workers.push(this); }
  reply(response: ResidentEngineWorkerResponse) {
    queueMicrotask(() => {
      if (!this.terminate.mock.calls.length)
        this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
    });
  }
  release() { for (const response of this.pending.splice(0)) this.reply(response); }
  postMessage(message: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule) {
    if (!("id" in message)) return;
    let response: ResidentEngineWorkerResponse = { id: message.id, ok: true };
    if (message.type === "openSource") {
      this.source = message.source.bytes![0]!;
      opens.push(this.source);
      response = this.source === 9
        ? { id: message.id, ok: false, error: "Failed to parse DOCX: not a zip archive" }
        : { id: message.id, ok: true, hostJson,
            sourceDocument: decodeDocxHostJson(hostJson, message.source.bytes!).document,
            stateVector: new ArrayBuffer(0) };
      if (held.has(this.source)) { this.pending.push(response); return; }
    } else if (message.type === "fontRequirements") {
      response = { id: message.id, ok: true, requirementsJson: "[]" };
    } else if (message.type === "commandInit") {
      const { session, incarnation, epoch, readOnly } = message.request;
      this.initial = { session, incarnation, epoch, version: 0, throughSeq: 0,
        outcomes: [], selection: null,
        editState: { readOnly, toolbar: null, effectiveFormatting: null, storedMarks: null,
          canUndo: false, canRedo: false, hyperlink: null, tableContext: null },
        frame: frame(1, 0) };
      response = { id: message.id, ok: true, commandInitial: this.initial };
    } else if (message.type === "edit" && this.initial) {
      const command = message.command;
      if (command.kind === "insertText") this.text += (command.payload as { text: string }).text;
      const ack = { ...this.initial, version: this.initial.version + 1, throughSeq: command.seq,
        outcomes: [{ seq: command.seq, status: "applied" as const }],
        frame: frame(command.seq + 1, command.seq) };
      this.initial = ack;
      response = { id: message.id, ok: true, commandAck: ack };
    } else if (message.type === "requestFullFrame" && this.initial) {
      response = { id: message.id, ok: true, commandAck: this.initial };
    } else if (message.type === "save") {
      this.saves += 1;
      response = { id: message.id, ok: true, saved: Uint8Array.of(this.source).buffer,
        updates: [], version: String(this.initial?.version ?? 0) };
      if (this.holdSave) { this.pending.push(response); return; }
    } else if (message.type === "documentRead") {
      response = { id: message.id, ok: true, read: { version: "0", value:
        message.read.kind === "toolbarMetadata" ? { styles: [], fontFamilies: [] }
        : message.read.kind === "sectionProperties" ? { version: "0", sections: [], paragraphs: [] }
        : message.read.kind === "readParagraphs" ? { ok: true, view: "accepted", version: "0", paragraphs: [] }
        : [] } };
    }
    this.reply(response);
  }
}

beforeEach(() => {
  opens.length = 0;
  workers.length = 0;
  held.clear();
  readied.mockClear();
  readiness.clear();
  delayAcknowledgements = false;
  flushFailure = null;
  delayedEdits.clear();
  const textEngine = spyOn(measurement, "getRustTextEngine").mockResolvedValue({
    registerFont: () => 0, registerSubstituteFont: (base) => base, clearFonts() {},
  });
  restorers.push(() => textEngine.mockRestore());
  const previousWorker = globalThis.Worker;
  Object.defineProperty(globalThis, "Worker", { configurable: true, writable: true, value: FakeWorker });
  restorers.push(() => Object.defineProperty(globalThis, "Worker", {
    configurable: true, writable: true, value: previousWorker,
  }));
  for (const [target, key, value] of [
    [globalThis, "OffscreenCanvas", class OffscreenCanvas {}],
    [globalThis, "createImageBitmap", async () => ({ width: 200, height: 300, close() {} })],
    [HTMLCanvasElement.prototype, "transferControlToOffscreen", () => ({})],
  ] as const) {
    const previous = Object.getOwnPropertyDescriptor(target, key);
    Object.defineProperty(target, key, { configurable: true, value });
    restorers.push(() => {
      if (previous) Object.defineProperty(target, key, previous);
      else Reflect.deleteProperty(target, key);
    });
  }
  const browser = window as unknown as GlobalWindow;
  const previousUrl = location.href;
  browser.happyDOM.setURL("http://localhost/?format=docx");
  restorers.push(() => browser.happyDOM.setURL(previousUrl));
  const fetching = spyOn(globalThis, "fetch").mockImplementation(Object.assign(
    async (input: Parameters<typeof fetch>[0]) => new Response(Uint8Array.of(Number(String(input).split("/").at(-1)))),
    { preconnect: globalThis.fetch.preconnect }
  ));
  restorers.push(() => fetching.mockRestore());
  const confirming = spyOn(window, "confirm").mockReturnValue(true);
  restorers.push(() => confirming.mockRestore());
});
afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function pick(source: number, name = `file-${source}.docx`) {
  await waitFor(() => expect(window.desktop).toBeDefined());
  await act(async () => {
    await window.desktop!.open({ id: String(source), name, url: `http://localhost/${source}` });
  });
}
async function opened(view: ReturnType<typeof render>, name: string) {
  await readyFor(name).promise;
  await waitFor(() => expect(view.container.querySelector(".document-name")?.textContent).toBe(name));
  return await view.findByRole("textbox", { name: "Document input" }) as HTMLTextAreaElement;
}
function type(input: HTMLTextAreaElement, text: string) {
  act(() => { input.dispatchEvent(new InputEvent("beforeinput", {
    bubbles: true, cancelable: true, inputType: "insertText", data: text,
  })); });
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function nativeBridge() {
  const activation = deferred<void>();
  const confirmations: ReturnType<typeof deferred<string>>[] = [];
  const documents = new Set(["1", "2", "3"]);
  let active: string | null = null;
  let retained: string | null = null;
  let dirty = false;
  let configuration: ReturnType<typeof deferred<void>> | null = null;
  const configurationBlocked = mock(() => {});
  const messages = mock(async (message: { type: string; id?: string; retainId?: string; dirty?: boolean; name?: string; bytes?: string; saveAs?: boolean }) => {
    if (message.type === "configuration") {
      if (configuration) {
        configurationBlocked();
        await configuration.promise;
      }
      return { format: "docx", name: "BetterOffice Docs", recent: [] };
    }
    if (message.type === "activate") {
      if (!message.id || !documents.has(message.id)) throw new Error("File is no longer available.");
      active = message.id;
      retained = message.retainId ?? null;
      dirty = false;
      for (const id of documents)
        if (id !== message.id && id !== message.retainId) documents.delete(id);
      if (message.id === "2") await activation.promise;
    } else if (message.type === "state" && message.id === active) {
      dirty = message.dirty ?? true;
    } else if (message.type === "confirmClose") {
      const confirmation = deferred<string>();
      confirmations.push(confirmation);
      return await confirmation.promise;
    } else if (message.type === "save") {
      if (message.id !== active && message.id !== retained) throw new Error("The file is no longer open.");
      return { name: message.name };
    }
  });
  const previous = window.webkit;
  window.webkit = { messageHandlers: { desktop: { postMessage: messages } } };
  restorers.push(() => { window.webkit = previous; });
  return { activation, confirmations, documents, messages,
    configurationBlocked, holdConfiguration() { configuration = deferred<void>(); return configuration; },
    get active() { return active; }, get dirty() { return dirty; } };
}

test("one desktop docx file open parses exactly once through the real editor mount", async () => {
  const view = render(<App />);
  await pick(1);
  await opened(view, "file-1.docx");
  expect(workers.filter(worker => worker.initial)).toHaveLength(1);
  expect(opens).toEqual([1]);
});

test("a refused docx keeps the old editor mounted with its unsaved state", async () => {
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  const worker = workers.find(worker => worker.initial)!;
  type(input, "unsaved");
  await waitFor(() => expect(worker.text).toBe("unsaved"));
  held.add(9);
  await pick(9);
  await waitFor(() => expect(opens).toContain(9));
  expect(input.isConnected).toBe(true);
  expect(input.closest("[inert]")).toBeNull();
  type(input, " while opening");
  await waitFor(() => expect(worker.text).toBe("unsaved while opening"));
  act(() => workers.find(worker => worker.source === 9)!.release());
  expect((await view.findByRole("alert")).textContent).toContain("Failed to parse DOCX");
  expect(view.getByRole("textbox", { name: "Document input" })).toBe(input);
  expect(view.container.querySelector(".document-name")?.textContent).toBe("file-1.docx · Edited");
  expect(worker.terminate).not.toHaveBeenCalled();
  expect(workers.filter(worker => worker.source === 1 && worker.initial)).toHaveLength(1);
  await waitFor(() => expect(workers.find(worker => worker.source === 9)!.terminate).toHaveBeenCalledTimes(1));
});

test("a second pick unmounts the pending editor and promotes its own mounted instance", async () => {
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  held.add(2);
  await pick(2);
  await waitFor(() => expect(opens).toContain(2));
  const pending = workers.find(worker => worker.source === 2)!;
  expect(input.isConnected).toBe(true);
  await pick(3);
  const replacement = await opened(view, "file-3.docx");
  expect(replacement).not.toBe(input);
  expect(input.isConnected).toBe(false);
  await waitFor(() => expect(pending.terminate).toHaveBeenCalledTimes(1));
  act(() => pending.release());
  expect(view.getByRole("textbox", { name: "Document input" })).toBe(replacement);
  expect(opens).toEqual([1, 2, 3]);
});

test.each(["cancel", "discard"])("promotion asks again when the current document is edited while activation is pending: %s", async (choice) => {
  const native = nativeBridge();
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  const current = workers.find(worker => worker.source === 1)!;
  native.documents.add("2");
  await pick(2);
  await waitFor(() => expect(native.active).toBe("2"));
  type(input, "during activation");
  await waitFor(() => expect(view.container.querySelector(".document-name")?.textContent).toBe("file-1.docx · Edited"));
  act(() => native.activation.resolve());
  await waitFor(() => expect(native.confirmations).toHaveLength(1));
  act(() => native.confirmations[0]!.resolve(choice));
  if (choice === "cancel") {
    await waitFor(() => expect(workers.find(worker => worker.source === 2)!.terminate).toHaveBeenCalledTimes(1));
    expect(view.getByRole("textbox", { name: "Document input" })).toBe(input);
    expect(current.text).toBe("during activation");
    expect(current.terminate).not.toHaveBeenCalled();
    expect(view.container.querySelector(".document-name")?.textContent).toBe("file-1.docx · Edited");
    expect(native.active).toBe("1");
    expect(native.dirty).toBe(true);
    expect(view.queryByRole("alert")).toBeNull();
  } else {
    await opened(view, "file-2.docx");
    expect(input.isConnected).toBe(false);
    expect(native.active).toBe("2");
    await waitFor(() => expect(current.terminate).toHaveBeenCalledTimes(1));
  }
});

test("promotion confirms an acknowledgement-pending edit before unmounting the current document", async () => {
  delayAcknowledgements = true;
  const native = nativeBridge();
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  native.documents.add("2");
  await pick(2);
  await waitFor(() => expect(native.active).toBe("2"));
  fireEvent.input(input, { target: { value: "during activation" } });
  expect(delayedEdits.get("file-1.docx")?.pending).toBe(true);
  expect(view.container.querySelector(".document-name")?.textContent).toBe("file-1.docx");
  act(() => native.activation.resolve());
  await waitFor(() => expect(native.confirmations).toHaveLength(1));
  expect(delayedEdits.get("file-1.docx")?.pending).toBe(false);
  expect(input.isConnected).toBe(true);
  act(() => native.confirmations[0]!.resolve("cancel"));
  await waitFor(() => expect(native.active).toBe("1"));
  expect(view.getByRole("textbox", { name: "Document input" })).toBe(input);
  expect(view.container.querySelector(".document-name")?.textContent).toBe("file-1.docx · Edited");
  expect(native.dirty).toBe(true);
});

test("promotion flushes input admitted while discard consent is pending", async () => {
  delayAcknowledgements = true;
  const native = nativeBridge();
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  native.documents.add("2");
  await pick(2);
  await waitFor(() => expect(native.active).toBe("2"));
  fireEvent.input(input, { target: { value: "first" } });
  act(() => native.activation.resolve());
  await waitFor(() => expect(native.confirmations).toHaveLength(1));
  fireEvent.input(input, { target: { value: "first second" } });
  expect(delayedEdits.get("file-1.docx")?.pending).toBe(true);
  act(() => native.confirmations[0]!.resolve("discard"));
  await waitFor(() => expect(native.confirmations).toHaveLength(2));
  expect(delayedEdits.get("file-1.docx")?.pending).toBe(false);
  expect(input.isConnected).toBe(true);
  act(() => native.confirmations[1]!.resolve("cancel"));
  await waitFor(() => expect(native.active).toBe("1"));
  expect(view.getByRole("textbox", { name: "Document input" })).toBe(input);
  expect(native.dirty).toBe(true);
});

test("a rejected input flush reports the error and keeps the current document mounted", async () => {
  delayAcknowledgements = true;
  const native = nativeBridge();
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  native.documents.add("2");
  await pick(2);
  await waitFor(() => expect(native.active).toBe("2"));
  fireEvent.input(input, { target: { value: "unacknowledged" } });
  flushFailure = new Error("The document worker stopped.");
  act(() => native.activation.resolve());
  expect((await view.findByRole("alert")).textContent).toContain("The document worker stopped.");
  expect(view.getByRole("textbox", { name: "Document input" })).toBe(input);
  expect(view.container.querySelector(".document-name")?.textContent).toBe("file-1.docx");
  expect(input.value).toBe("unacknowledged");
  expect(delayedEdits.get("file-1.docx")?.pending).toBe(true);
  expect(native.active).toBe("1");
  expect(native.confirmations).toEqual([]);
});

test("a save started while a replacement stages writes the current document", async () => {
  const blobs: Blob[] = [];
  const urls = spyOn(URL, "createObjectURL").mockImplementation((blob) => {
    blobs.push(blob as Blob);
    return "blob:saved";
  });
  const downloads: { name: string; mounted: boolean }[] = [];
  const click = spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
    downloads.push({ name: this.download, mounted: input.isConnected });
  });
  restorers.push(() => urls.mockRestore(), () => click.mockRestore());
  const view = render(<App />);
  await pick(1, "old.docx");
  const input = await opened(view, "old.docx");
  const current = workers.find(worker => worker.source === 1)!;
  current.holdSave = true;
  held.add(2);
  await pick(2);
  await waitFor(() => expect(opens).toContain(2));
  const replacement = workers.find(worker => worker.source === 2)!;
  let releaseYield!: () => void;
  const timer = spyOn(globalThis, "setTimeout").mockImplementationOnce(Object.assign((callback: TimerHandler, delay?: number, ...args: unknown[]) => {
    expect(delay).toBe(0);
    releaseYield = () => { if (typeof callback === "function") callback(...args); };
    return 0 as unknown as ReturnType<typeof setTimeout>;
  }, { __promisify__: globalThis.setTimeout.__promisify__ }) as unknown as typeof setTimeout);
  fireEvent.click(view.getByRole("button", { name: /^Save$/ }));
  timer.mockRestore();
  expect(releaseYield).toBeDefined();
  act(() => replacement.release());
  await waitFor(() => expect(readied).toHaveBeenCalledTimes(2));
  act(() => releaseYield());
  await waitFor(() => expect(current.saves).toBe(1));
  expect(input.isConnected).toBe(true);
  expect(downloads).toEqual([]);
  expect(replacement.saves).toBe(0);
  act(() => current.release());
  await opened(view, "file-2.docx");
  expect(downloads).toEqual([{ name: "old.docx", mounted: true }]);
  expect(new Uint8Array(await blobs[0]!.arrayBuffer())).toEqual(Uint8Array.of(1));
  expect(current.saves).toBe(1);
  await waitFor(() => expect(current.terminate).toHaveBeenCalledTimes(1));
  expect(workers.filter(worker => worker.source === 2 && worker.initial)).toHaveLength(1);
});

test("promotion reconfirms edits committed while the discard dialog is pending", async () => {
  const native = nativeBridge();
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  native.documents.add("2");
  await pick(2);
  await waitFor(() => expect(native.active).toBe("2"));
  type(input, "first");
  await waitFor(() => expect(workers.find(worker => worker.source === 1)!.text).toBe("first"));
  act(() => native.activation.resolve());
  await waitFor(() => expect(native.confirmations).toHaveLength(1));
  type(input, " second");
  await waitFor(() => expect(workers.find(worker => worker.source === 1)!.text).toBe("first second"));
  act(() => native.confirmations[0]!.resolve("discard"));
  await waitFor(() => expect(native.confirmations).toHaveLength(2));
  act(() => native.confirmations[1]!.resolve("cancel"));
  await waitFor(() => expect(workers.find(worker => worker.source === 2)!.terminate).toHaveBeenCalledTimes(1));
  expect(view.getByRole("textbox", { name: "Document input" })).toBe(input);
  expect(native.active).toBe("1");
  expect(native.dirty).toBe(true);
});

test("promotion checks edits committed while native configuration refresh is pending", async () => {
  const native = nativeBridge();
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  native.documents.add("2");
  await pick(2);
  await waitFor(() => expect(native.active).toBe("2"));
  const configuration = native.holdConfiguration();
  act(() => native.activation.resolve());
  await waitFor(() => expect(native.configurationBlocked).toHaveBeenCalledTimes(1));
  type(input, "during refresh");
  await waitFor(() => expect(workers.find(worker => worker.source === 1)!.text).toBe("during refresh"));
  act(() => configuration.resolve());
  await waitFor(() => expect(native.confirmations).toHaveLength(1));
  act(() => native.confirmations[0]!.resolve("cancel"));
  await waitFor(() => expect(workers.find(worker => worker.source === 2)!.terminate).toHaveBeenCalledTimes(1));
  expect(view.getByRole("textbox", { name: "Document input" })).toBe(input);
  expect(native.active).toBe("1");
});

test("promotion waits for a native save and uses its saved checkpoint", async () => {
  const native = nativeBridge();
  const view = render(<App />);
  await pick(1);
  const input = await opened(view, "file-1.docx");
  const current = workers.find(worker => worker.source === 1)!;
  native.documents.add("2");
  await pick(2);
  await waitFor(() => expect(native.active).toBe("2"));
  type(input, "saved during activation");
  await waitFor(() => expect(current.text).toBe("saved during activation"));
  current.holdSave = true;
  fireEvent.click(view.getByRole("button", { name: /^Save$/ }));
  act(() => native.activation.resolve());
  await waitFor(() => expect(current.saves).toBe(1));
  expect(input.isConnected).toBe(true);
  expect(native.confirmations).toEqual([]);
  act(() => current.release());
  await opened(view, "file-2.docx");
  expect(native.messages.mock.calls.filter(([message]) => message.type === "save")).toEqual([
    [{ type: "save", id: "1", name: "file-1.docx", bytes: btoa(String.fromCharCode(1)), saveAs: false }],
  ]);
  expect(native.confirmations).toEqual([]);
  expect(view.queryByRole("alert")).toBeNull();
  expect(native.active).toBe("2");
  await waitFor(() => expect(current.terminate).toHaveBeenCalledTimes(1));
});
