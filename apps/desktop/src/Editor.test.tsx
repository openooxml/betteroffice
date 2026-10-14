import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, expect, mock, spyOn, test } from "bun:test";
import { forwardRef, useEffect, useImperativeHandle } from "react";
import type { XlsxAsyncEditorApi, XlsxAsyncEditorProps, XlsxCommandResult } from "@betteroffice/xlsx-react";
import type { PptxAsyncEditorApi, PptxAsyncEditorProps } from "@betteroffice/pptx-react";
import type { EditorApi } from "./Editor";

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

interface DocxProps {
  onSave(buffer: ArrayBuffer): void;
  onError(error: Error): void;
  mode?: "editing" | "viewing";
  readOnly?: boolean;
  experimentalPeerless?: boolean;
  onChange?: () => void;
  onDocumentChange?: () => void;
}
let docxProps: DocxProps;
let save: (props: DocxProps) => Promise<ArrayBuffer | null> = async () =>
  null;
let workbookProps: XlsxAsyncEditorProps;
let workbookSave: (props: XlsxAsyncEditorProps) => Promise<XlsxCommandResult>;
let deferWorkbookReady = false;
let readyWorkbook: (() => void) | undefined;
let disposeWorkbookApi: (() => void) | undefined;
mock.module("@betteroffice/xlsx-react", () => ({
  XlsxEditor: (props: XlsxAsyncEditorProps) => {
    workbookProps = props;
    useEffect(() => {
      let dispose: void | (() => void);
      readyWorkbook = () => {
        dispose = props.onReady?.({
          focus() {},
          commands: { execute: () => workbookSave(workbookProps) },
        } as unknown as XlsxAsyncEditorApi);
        disposeWorkbookApi = typeof dispose === "function" ? dispose : undefined;
      };
      if (!deferWorkbookReady) readyWorkbook();
      return () => { if (typeof dispose === "function") dispose(); };
    }, []);
    return null;
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
        dispose = props.onReady?.({ saveAsync: () => savePresentation(), focus() {} } as PptxAsyncEditorApi);
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
class Provider {
  constructor(replica: unknown) { providerReplicas.push(replica); }
  onStatus() { return () => {}; }
  onError() { return () => {}; }
  connect() {}
  destroy() {}
}
mock.module("@betteroffice/docx-react", () => ({
  DocxEditor: forwardRef<unknown, DocxProps>(function DocxEditor(props, ref) {
    docxProps = props;
    useImperativeHandle(ref, () => ({ save: () => save(props) }));
    return <span data-testid="docx-mode" data-mode={props.mode} data-readonly={String(props.readOnly ?? false)} />;
  }),
}));
for (const [specifier, exports] of Object.entries({
  "@betteroffice/docx-react/peerless-harness": { registerBrowserHarness: () => () => {} },
  "@betteroffice/docx-react/styles.css": {},
  "@betteroffice/docx/layout": { configureDefaultFonts: () => {} },
  "@betteroffice/docx/utils": { setGoogleFontsEnabled: () => {} },
  "@betteroffice/docx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/xlsx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/pptx/editor": { CollaborationProvider: Provider },
  "@betteroffice/fonts": {
    loadBundledFontBytes: () => {},
    resolveLastResortFace: () => {},
    resolveMetricCompatFace: () => {},
  },
}))
  mock.module(specifier, () => exports);

const { act, cleanup, fireEvent, render } = await import("@testing-library/react");
const { default: Editor, prepareFile } = await import("./Editor");

afterEach(() => {
  cleanup();
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

test("peerless harness mode controls commit the editor read-only prop", async () => {
  const scope = globalThis as typeof globalThis & { __DOCX_PEERLESS_HARNESS__?: boolean };
  const previousFlag = scope.__DOCX_PEERLESS_HARNESS__;
  const previousUrl = location.href;
  scope.__DOCX_PEERLESS_HARNESS__ = true;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL("http://localhost/?format=docx&peerless=1");
  try {
    const view = render(<Editor
      file={{ bytes: new Uint8Array(), format: "docx" }} name="input.docx" room={null}
      onSave={() => {}} onChange={() => {}} onError={() => {}} onReady={() => {}}
      onOpen={async () => {}} onStatus={() => {}}
    />);
    const mode = view.getByTestId("docx-mode");
    expect(mode.dataset.readonly).toBe("false");
    fireEvent.click(view.getByRole("button", { name: "Viewing" }));
    expect(mode.dataset.mode).toBe("viewing");
    expect(mode.dataset.readonly).toBe("true");
    fireEvent.click(view.getByRole("button", { name: "Editing" }));
    expect(mode.dataset.mode).toBe("editing");
    expect(mode.dataset.readonly).toBe("false");
    view.unmount();
  } finally {
    browser.happyDOM.setURL(previousUrl);
    if (previousFlag === undefined) delete scope.__DOCX_PEERLESS_HARNESS__;
    else scope.__DOCX_PEERLESS_HARNESS__ = previousFlag;
  }
});

async function mount(format: "docx" | "xlsx" | "pptx" = "docx", collaborative = false, onChange = () => {}) {
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
        onSave={(bytes) => saved.push(bytes)}
        onChange={onChange}
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

test("peerless preview reports dirty state through worker document events", async () => {
  const previousUrl = location.href;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL("http://localhost/?format=docx&peerless=1");
  try {
    const changed = mock(() => {});
    await mount("docx", false, changed);
    expect(docxProps!.experimentalPeerless).toBe(true);
    expect(docxProps!.onChange).toBeUndefined();
    docxProps!.onDocumentChange?.();
    expect(changed).toHaveBeenCalledTimes(1);
  } finally {
    browser.happyDOM.setURL(previousUrl);
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
