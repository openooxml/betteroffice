import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { DocxEditor, type DocxEditorRef, type DocxErrorInfo, type TrackedChangesOptions } from "@betteroffice/docx-react";
import { XlsxEditor, type XlsxAsyncEditorApi } from "@betteroffice/xlsx-react";
import { PptxEditor, type PptxAsyncEditorApi } from "@betteroffice/pptx-react";
import { configureDefaultFonts } from "@betteroffice/docx/layout";
import { setGoogleFontsEnabled } from "@betteroffice/docx/utils";
import { CollaborationProvider as DocxProvider } from "@betteroffice/docx/collaboration";
import { CollaborationProvider as XlsxProvider } from "@betteroffice/xlsx/collaboration";
import { CollaborationProvider as PptxProvider, type AsyncCollaborationReplica } from "@betteroffice/pptx/collaboration";
import type { PptxFontFace } from "@betteroffice/pptx";
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
} from "@betteroffice/fonts";
import { createRoomTransport } from "../../demo/app/collab/createRoomTransport";
import type { CollaborationReplica } from "../../demo/app/collab/types";
import type { Format } from "./bridge";
import { savedBytes } from "./session";
import "@betteroffice/docx-react/styles.css";

const EDITOR_READY_TIMEOUT_MS = 30_000;
const WORKBOOK_DISPOSED = "The workbook editor was disposed.";
const PRESENTATION_DISPOSED = "The presentation editor was disposed.";

configureDefaultFonts({ load: () => import("@betteroffice/fonts") });
setGoogleFontsEnabled(false);

function legacyDocx() {
  if (typeof location === "undefined") return false;
  const query = new URLSearchParams(location.search);
  return query.get("peerless") === "0";
}

function trackedChanges(): TrackedChangesOptions | undefined {
  if (typeof location === "undefined") return undefined;
  const query = new URLSearchParams(location.search);
  const agents = new Set(query.get("agents")?.split(",").filter(Boolean));
  if (query.get("review") !== "dots" && !agents.size) return undefined;
  return {
    ...(query.get("review") === "dots" ? { display: "dots" as const } : {}),
    ...(agents.size
      ? { getAuthorAppearance: ({ name }: { name: string }) => (agents.has(name) ? { kind: "agent" as const } : undefined) }
      : {}),
  };
}

const TRACKED_CHANGES = trackedChanges();

function colorMode(): "light" | "dark" {
  if (typeof location === "undefined") return "light";
  return new URLSearchParams(location.search).get("theme") === "dark" ? "dark" : "light";
}

export interface PreparedFile {
  bytes: Uint8Array;
  format: Format;
  fonts?: PptxFontFace[];
}

export async function prepareFile(
  bytes: Uint8Array,
  format: Format
): Promise<PreparedFile> {
  if (format === "docx") {
    if (!legacyDocx()) {
      return { bytes, format };
    }
    const { parseDocx } = await import("@betteroffice/docx/docx");
    await parseDocx(bytes);
    return { bytes, format };
  }
  if (format === "xlsx") {
    if (typeof Worker === "function" && typeof OffscreenCanvas === "function") {
      const { openWorkbookSession } = await import("@betteroffice/xlsx");
      const session = await openWorkbookSession(bytes);
      await session.dispose();
      return { bytes, format };
    }
    const { initWasm, openWorkbook } = await import("@betteroffice/xlsx");
    await initWasm();
    openWorkbook(bytes).dispose();
    return { bytes, format };
  }
  const fonts = await Promise.all(
    [false, true].flatMap((bold) =>
      [false, true].map(async (italic) => {
        const face =
          resolveMetricCompatFace("Arial", bold, italic) ??
          resolveLastResortFace("Arial", bold, italic);
        return {
          family: "Arial",
          bold,
          italic,
          bytes: new Uint8Array(await loadBundledFontBytes(face)),
        };
      })
    )
  );
  if (typeof Worker === "function" && typeof OffscreenCanvas === "function") {
    const { openPresentationSession } = await import("@betteroffice/pptx/worker");
    const session = await openPresentationSession(bytes, { fonts });
    await session.dispose();
  } else {
    const { initWasm, openPresentation } = await import("@betteroffice/pptx");
    await initWasm();
    openPresentation(bytes, { fonts }).dispose();
  }
  return { bytes, format, fonts };
}

export interface EditorApi {
  flushPendingInput?(): Promise<void>;
  serialize(): Promise<Uint8Array>;
  focus(): void;
  print(): void;
}
export interface Room {
  id: string;
  name: string;
}

export default function Editor({
  file,
  name,
  room,
  onSave,
  onChange,
  onError,
  onReady,
  onOpen,
  onStatus,
}: {
  file: PreparedFile;
  name: string;
  room: Room | null;
  onSave(bytes: Uint8Array): void;
  onChange(): void;
  onError(error: Error): void;
  onReady(api: EditorApi | null): void;
  onOpen(file: File): Promise<void>;
  onStatus(status: string): void;
}) {
  const serialization = useRef<{ failure: Error | null; bytes?: Uint8Array } | null>(null);
  const serializations = useRef<Promise<unknown>>(Promise.resolve());
  const docx = useRef<DocxEditorRef>(null);
  const xlsx = useRef<XlsxAsyncEditorApi | null>(null);
  const pptx = useRef<PptxAsyncEditorApi | null>(null);
  const apiWaiter = useRef<{ resolve(): void; reject(error: Error): void } | null>(null);
  const apiFailure = useRef<Error | null>(null);
  const harness = (globalThis as typeof globalThis & { __DOCX_PEERLESS_HARNESS__?: boolean }).__DOCX_PEERLESS_HARNESS__
    && file.format === "docx" && !legacyDocx() && !room;
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
  const clientId = useRef(
    crypto.getRandomValues(new Uint32Array(1))[0] & 0x7fffffff || 1
  ).current;
  const [replica, setReplica] = useState<CollaborationReplica | null>(null);
  const [presentationReplica, setPresentationReplica] = useState<AsyncCollaborationReplica | null>(null);
  const [provider, setProvider] = useState<
    DocxProvider | XlsxProvider | PptxProvider | null
  >(null);
  const [painted, setPainted] = useState(false);
  const firstPainted = useCallback(() => {
    if (!legacyDocx()) setPainted(true);
  }, []);
  const presentationPresence = provider instanceof PptxProvider ? provider : undefined;
  const presentationCollaboration = useMemo(() => room ? {
    roomId: room.id,
    clientId,
    onReplica: setPresentationReplica,
    presence: presentationPresence,
  } : undefined, [room?.id, clientId, presentationPresence]);
  const callbacks = useRef({ onSave, onChange, onError, onStatus, onReady });
  callbacks.current = { onSave, onChange, onError, onStatus, onReady };

  useEffect(() => {
    if (!room || (file.format === "pptx" ? !presentationReplica : !replica)) return;
    const transport = createRoomTransport(
      "https://betteroffice-collaboration-relay.elia7.workers.dev",
      room.id
    );
    const options = { user: { name: room.name } };
    const next =
      file.format === "docx"
        ? new DocxProvider(replica!, transport, options)
        : file.format === "xlsx"
        ? new XlsxProvider(replica!, transport, options)
        : new PptxProvider(presentationReplica!, transport, options);
    setProvider(next);
    const unsubscribe = next.onStatus(({ status, synced }) =>
      callbacks.current.onStatus(synced ? "Connected" : status)
    );
    const unsubscribeError = next.onError((error) =>
      callbacks.current.onError(error)
    );
    void next.connect();
    return () => {
      unsubscribe();
      unsubscribeError();
      next.destroy();
      setProvider(null);
    };
  }, [file.format, replica, presentationReplica, room]);

  const changed = useCallback(() => callbacks.current.onChange(), []);
  const saveBytes = useCallback(
    (bytes: Uint8Array) => callbacks.current.onSave(bytes),
    []
  );
  const apiUnavailable = useCallback((failure: Error) => {
    apiFailure.current = failure;
    apiWaiter.current?.reject(failure);
    apiWaiter.current = null;
  }, []);
  const error = useCallback((error: Error) => {
    if (serialization.current) serialization.current.failure = error;
    if (file.format === "xlsx" ? !xlsx.current : file.format === "pptx" && !pptx.current)
      apiUnavailable(error);
    callbacks.current.onError(error);
  }, [file.format, apiUnavailable]);
  const docxError = useCallback((cause: Error, info?: DocxErrorInfo) => {
    if (info?.recoverable) console.warn(cause);
    else error(cause);
  }, [error]);
  const apiReady = useCallback(
    <Api,>(ref: { current: Api | null }, api: Api, disposed: string) => {
      ref.current = api;
      apiFailure.current = null;
      apiWaiter.current?.resolve();
      apiWaiter.current = null;
      return () => {
        if (ref.current !== api) return;
        ref.current = null;
        apiUnavailable(new Error(disposed));
      };
    },
    [apiUnavailable]
  );
  const workbookReady = useCallback(
    (api: XlsxAsyncEditorApi) => apiReady(xlsx, api, WORKBOOK_DISPOSED),
    [apiReady]
  );
  const presentationReady = useCallback(
    (api: PptxAsyncEditorApi) => apiReady(pptx, api, PRESENTATION_DISPOSED),
    [apiReady]
  );

  const serializeWorkbook = async (api: XlsxAsyncEditorApi) => {
    const result = await api.commands.execute("save", null);
    if (!result.ok) throw new Error(result.failure.message);
    return serialization.current?.bytes;
  };

  useEffect(() => {
    if (file.format === "docx" && !legacyDocx() && !painted) return;
    let active = true;
    const disposed = new Error(
      file.format === "xlsx" ? WORKBOOK_DISPOSED : PRESENTATION_DISPOSED
    );
    const whenReady = async <Api,>(ref: { current: Api | null }): Promise<Api> => {
      if (!active) throw disposed;
      let api = ref.current;
      if (!api) {
        if (apiFailure.current) throw apiFailure.current;
        let timeout: ReturnType<typeof setTimeout> | undefined;
        await new Promise<void>((resolve, reject) => {
          const waiter = { resolve, reject };
          apiWaiter.current = waiter;
          timeout = setTimeout(() => {
            if (apiWaiter.current !== waiter) return;
            apiUnavailable(new Error(
              "The editor is still opening the file. Try again in a moment."
            ));
          }, EDITOR_READY_TIMEOUT_MS);
        }).finally(() => clearTimeout(timeout));
        api = ref.current;
      }
      if (!active || !api || ref.current !== api) throw disposed;
      return api;
    };
    callbacks.current.onReady({
      ...(file.format === "docx" ? {
        async flushPendingInput() {
          if (!active || !docx.current) throw new Error("The document editor was disposed.");
          await docx.current.flushPendingInput();
        },
      } : {}),
      print() {
        if (file.format === "docx") docx.current?.print();
        else window.print();
      },
      focus() {
        if (file.format === "docx") docx.current?.focus();
        else if (file.format === "xlsx") xlsx.current?.focus();
        else pptx.current?.focus();
      },
      serialize() {
        const attempt = async () => {
          if (document.activeElement instanceof HTMLElement)
            document.activeElement.blur();
          await new Promise<void>((resolve) => setTimeout(resolve, 0));
          const current: { failure: Error | null; bytes?: Uint8Array } = { failure: null };
          serialization.current = current;
          try {
            const bytes =
              file.format === "docx"
                ? await docx.current?.save()
                : file.format === "xlsx"
                ? await serializeWorkbook(await whenReady(xlsx))
                : await (await whenReady(pptx)).saveAsync();
            return savedBytes(bytes, current.failure);
          } finally {
            serialization.current = null;
          }
        };
        const next = serializations.current.then(attempt, attempt);
        serializations.current = next.catch(() => undefined);
        return next;
      },
    });
    return () => {
      active = false;
      if (file.format !== "docx") apiUnavailable(disposed);
      callbacks.current.onReady(null);
    };
  }, [file.format, painted, apiUnavailable]);

  if (file.format === "docx")
    return (
      <>
      {harness && <div data-testid="peerless-harness-mode">
        <button type="button" aria-pressed={harnessMode === "viewing"} onClick={() => setHarnessMode("viewing")}>Viewing</button>
        <button type="button" aria-pressed={harnessMode === "editing"} onClick={() => setHarnessMode("editing")}>Editing</button>
      </div>}
      <DocxEditor
        ref={docx}
        {...(legacyDocx() ? { legacySyncApi: true as const } : { legacySyncApi: false as const })}
        trackedChanges={TRACKED_CHANGES}
        mode={harness ? harnessMode : undefined}
        readOnly={harness ? harnessMode === "viewing" : undefined}
        documentBuffer={file.bytes}
        downloadOnSave={false}
        onSave={(buffer) => {
          if (!serialization.current) saveBytes(new Uint8Array(buffer));
        }}
        documentName={name}
        onDocumentChange={changed}
        onFirstPagePainted={firstPainted}
        onError={docxError}
        onOpen={onOpen}
        showToolbar
        showRuler
        showZoomControl
        colorMode={colorMode()}
        onPrint={() => undefined}
        collaboration={
          room
            ? {
                clientId,
                user: { name: room.name },
                onReplica: setReplica,
                presence:
                  provider instanceof DocxProvider ? provider : undefined,
              }
            : undefined
        }
      />
      </>
    );
  if (file.format === "xlsx")
    return (
      <XlsxEditor
        file={file.bytes}
        fileName={name}
        onSave={(bytes) => {
          if (serialization.current) serialization.current.bytes = bytes;
          else saveBytes(bytes);
        }}
        onChange={changed}
        onError={error}
        onReady={workbookReady}
        collaboration={
          room
            ? {
                clientId,
                onReplica: setReplica,
                provider:
                  provider instanceof XlsxProvider ? provider : undefined,
              }
            : undefined
        }
      />
    );
  return (
    <PptxEditor
      file={file.bytes}
      fileName={name}
      fonts={file.fonts ?? []}
      onSave={saveBytes}
      onChange={changed}
      onReady={presentationReady}
      onError={error}
      collaboration={presentationCollaboration}
    />
  );
}
