import { useEffect, useState } from "react";
import SharedEditor, {
  legacyDocx,
  type EditorProps,
} from "../../../shared/editor-host/Editor";

export {
  prepareFile,
  type EditorApi,
  type PreparedFile,
  type Room,
} from "../../../shared/editor-host/Editor";

const RELAY_ORIGIN =
  "https://betteroffice-collaboration-relay.elia7.workers.dev";

export type DesktopEditorProps = Omit<EditorProps, "relayOrigin" | "mode" | "readOnly">;

export default function Editor(props: DesktopEditorProps) {
  const harness = (globalThis as typeof globalThis & { __DOCX_PEERLESS_HARNESS__?: boolean }).__DOCX_PEERLESS_HARNESS__
    && props.file.format === "docx" && !legacyDocx() && !props.room;
  const [harnessMode, setHarnessMode] = useState<"editing" | "viewing">("editing");
  useEffect(() => {
    if (!(globalThis as typeof globalThis & { __DOCX_PEERLESS_HARNESS__?: boolean }).__DOCX_PEERLESS_HARNESS__) return;
    if (!harness) return;
    let current = true;
    let unregister: (() => void) | undefined;
    void import("@betteroffice/docx-react/peerless-harness").then((module) => {
      if (current) unregister = module.registerBrowserHarness({ setMode: setHarnessMode });
    });
    return () => { current = false; unregister?.(); };
  }, [harness]);
  return (
    <>
      {harness && <div data-testid="peerless-harness-mode">
        <button type="button" aria-pressed={harnessMode === "viewing"} onClick={() => setHarnessMode("viewing")}>Viewing</button>
        <button type="button" aria-pressed={harnessMode === "editing"} onClick={() => setHarnessMode("editing")}>Editing</button>
      </div>}
      <SharedEditor
        {...props}
        relayOrigin={RELAY_ORIGIN}
        mode={harness ? harnessMode : undefined}
        readOnly={harness ? harnessMode === "viewing" : undefined}
      />
    </>
  );
}
