import { useEffect, useState } from "react";
import { locales as docxLocales } from "@betteroffice/docx-i18n";
import { locales as xlsxLocales } from "@betteroffice/xlsx-i18n";
import { locales as pptxLocales } from "@betteroffice/pptx-i18n";
import SharedEditor, {
  legacyDocx,
  type EditorMessages,
  type EditorProps,
  type EditorTranslations,
} from "../../../shared/editor-host/Editor";
import { locale, t } from "./i18n";

export {
  prepareFile,
  type EditorApi,
  type PreparedFile,
  type Room,
} from "../../../shared/editor-host/Editor";

const RELAY_ORIGIN =
  "https://betteroffice-collaboration-relay.elia7.workers.dev";

const TRANSLATIONS: EditorTranslations = {
  docx: docxLocales[locale],
  xlsx: xlsxLocales[locale],
  pptx: pptxLocales[locale],
};

const MESSAGES: EditorMessages = {
  connected: t("editor.connected"),
  connecting: t("titleBar.connecting"),
  disconnected: t("editor.disconnected"),
  documentDisposed: t("editor.documentDisposed"),
  workbookDisposed: t("editor.workbookDisposed"),
  presentationDisposed: t("editor.presentationDisposed"),
  stillOpening: t("editor.stillOpening"),
};

export type DesktopEditorProps = Omit<EditorProps, "relayOrigin" | "mode" | "readOnly" | "i18n" | "messages">;

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
        <button type="button" aria-pressed={harnessMode === "viewing"} onClick={() => setHarnessMode("viewing")}>{t("editor.viewing")}</button>
        <button type="button" aria-pressed={harnessMode === "editing"} onClick={() => setHarnessMode("editing")}>{t("editor.editing")}</button>
      </div>}
      <SharedEditor
        {...props}
        relayOrigin={RELAY_ORIGIN}
        i18n={TRANSLATIONS}
        messages={MESSAGES}
        mode={harness ? harnessMode : undefined}
        readOnly={harness ? harnessMode === "viewing" : undefined}
      />
    </>
  );
}
