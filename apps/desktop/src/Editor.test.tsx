import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, expect, mock, test } from "bun:test";
import { forwardRef, useEffect, useImperativeHandle } from "react";
import type { DocxErrorInfo } from "@betteroffice/docx-react";

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
mock.module("@betteroffice/xlsx-react", () => ({
  XlsxEditor: () => null,
}));
mock.module("@betteroffice/pptx-react", () => ({
  PptxEditor: () => null,
}));
class Provider {
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
    loadBundledFontBytes: async () => Uint8Array.of(1),
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

