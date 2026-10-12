import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, expect, mock, test } from "bun:test";
import { forwardRef, useImperativeHandle } from "react";
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
class Provider {}
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
  "@betteroffice/xlsx-react": { XlsxEditor: () => null },
  "@betteroffice/pptx-react": { PptxEditor: () => null },
  "@betteroffice/docx/layout": { configureDefaultFonts: () => {} },
  "@betteroffice/docx/utils": { setGoogleFontsEnabled: () => {} },
  "@betteroffice/docx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/xlsx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/pptx": { CollaborationProvider: Provider },
  "@betteroffice/fonts": {
    loadBundledFontBytes: () => {},
    resolveLastResortFace: () => {},
    resolveMetricCompatFace: () => {},
  },
}))
  mock.module(specifier, () => exports);

const { act, cleanup, fireEvent, render } = await import("@testing-library/react");
const { default: Editor } = await import("./Editor");

afterEach(cleanup);
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

async function mount(onChange = () => {}) {
  const ready: { api?: EditorApi } = {};
  const errors: string[] = [];
  const saved: Uint8Array[] = [];
  await act(async () => {
    render(
      <Editor
        file={{ bytes: new Uint8Array(), format: "docx" }}
        name="saved.docx"
        room={null}
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
  return { api: ready.api!, errors, saved };
}

test("peerless preview reports dirty state through worker document events", async () => {
  const previousUrl = location.href;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL("http://localhost/?format=docx&peerless=1");
  try {
    const changed = mock(() => {});
    await mount(changed);
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
