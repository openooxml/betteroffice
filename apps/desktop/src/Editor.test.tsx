import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, expect, mock, spyOn, test } from "bun:test";
import { forwardRef, useEffect, useImperativeHandle } from "react";
import type { XlsxAsyncEditorApi, XlsxAsyncEditorProps, XlsxCommandResult } from "@betteroffice/xlsx-react";
import type { PptxAsyncEditorApi, PptxAsyncEditorProps } from "@betteroffice/pptx-react";
import type { DocxErrorInfo } from "@betteroffice/docx-react";
import type { EditorApi } from "./Editor";

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

interface DocxProps {
  documentBuffer: Uint8Array;
  onFirstPagePainted(): void;
  onSave(buffer: ArrayBuffer): void;
  onError(error: Error, info?: DocxErrorInfo): void;
  mode?: "editing" | "viewing";
  readOnly?: boolean;
  legacySyncApi?: boolean;
  onChange?: () => void;
  onDocumentChange?: () => void;
}
let docxProps: DocxProps;
const parseDocx = mock(async () => ({}));
const flushDocxInput = mock(async () => {});
let save: (props: DocxProps) => Promise<ArrayBuffer | null> = async () =>
  null;
const workerOpens: Uint8Array[] = [];
let workerOpenFailure: Error | null = null;
let workersDestroyed = 0;
class ResidentEngineWorkerClient {
  async openSource({ bytes }: { bytes: Uint8Array }) {
    workerOpens.push(bytes);
    if (workerOpenFailure) throw workerOpenFailure;
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
    useImperativeHandle(ref, () => ({ save: () => save(props), flushPendingInput: flushDocxInput }));
    useEffect(() => {
      if (props.legacySyncApi) return;
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
const loadFont = mock(async (_face: unknown): Promise<Uint8Array> => Uint8Array.of(1));
const openPresentationSession = mock(async (_bytes: Uint8Array, _options: unknown) => ({ dispose: async () => {} }));
mock.module("@betteroffice/pptx/worker", () => ({ openPresentationSession }));

for (const [specifier, exports] of Object.entries({
  "@betteroffice/docx-react/peerless-harness": { registerBrowserHarness: () => () => {} },
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

const { act, cleanup, fireEvent, render, waitFor } = await import("@testing-library/react");
const { default: Editor } = await import("./Editor");
const { App } = await import("./App");

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
  browser.happyDOM.setURL("http://localhost/?format=docx");
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

test("a docx the worker cannot open is refused before it replaces the open file", async () => {
  const previousUrl = location.href;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL("http://localhost/?format=docx");
  parseDocx.mockClear();
  workerOpens.length = 0;
  workersDestroyed = 0;
  const fetching = mock(async () => new Response(Uint8Array.of(1, 2)));
  const originalFetch = globalThis.fetch;
  globalThis.fetch = Object.assign(fetching, { preconnect: originalFetch.preconnect });
  try {
    const view = render(<App />);
    await waitFor(() => expect(window.desktop).toBeDefined());
    act(() => { void window.desktop!.open({ id: "old", name: "old.docx", url: "http://localhost/old" }); });
    await waitFor(() => expect(view.container.querySelector(".document-name")?.textContent).toBe("old.docx"));
    const old = view.getByTestId("docx-mode");
    workerOpenFailure = new Error("Failed to parse DOCX: not a zip archive");
    act(() => { void window.desktop!.open({ id: "bad", name: "bad.docx", url: "http://localhost/bad" }); });
    expect((await view.findByRole("alert")).textContent).toContain("Failed to parse DOCX");
    expect(view.getByTestId("docx-mode")).toBe(old);
    expect(view.container.querySelector(".document-name")?.textContent).toBe("old.docx");
    expect(workerOpens).toHaveLength(2);
    await waitFor(() => expect(workersDestroyed).toBe(1));
    expect(parseDocx).not.toHaveBeenCalled();
  } finally {
    workerOpenFailure = null;
    globalThis.fetch = originalFetch;
    browser.happyDOM.setURL(previousUrl);
  }
});

