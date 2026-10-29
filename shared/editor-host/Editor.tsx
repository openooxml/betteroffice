import { useCallback, useEffect, useMemo, useRef, useState, type ComponentProps } from "react";
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
  type FontDecoder,
} from "@betteroffice/fonts";
import { createRoomTransport } from "./collaboration/createRoomTransport";
import type { CollaborationReplica } from "./collaboration/types";
import { savedBytes } from "./session";
import "@betteroffice/docx-react/styles.css";

const EDITOR_READY_TIMEOUT_MS = 30_000;

export interface EditorMessages {
  connected: string;
  connecting: string;
  disconnected: string;
  documentDisposed: string;
  workbookDisposed: string;
  presentationDisposed: string;
  stillOpening: string;
}

export const DEFAULT_EDITOR_MESSAGES: EditorMessages = {
  connected: "Connected",
  connecting: "Connecting…",
  disconnected: "Disconnected",
  documentDisposed: "The document editor was disposed.",
  workbookDisposed: "The workbook editor was disposed.",
  presentationDisposed: "The presentation editor was disposed.",
  stillOpening: "The editor is still opening the file. Try again in a moment.",
};

export interface EditorTranslations {
  docx?: ComponentProps<typeof DocxEditor>["i18n"];
  xlsx?: ComponentProps<typeof XlsxEditor>["i18n"];
  pptx?: ComponentProps<typeof PptxEditor>["i18n"];
}

configureDefaultFonts({ load: () => import("@betteroffice/fonts") });
setGoogleFontsEnabled(false);

export function legacyDocx() {
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

export type Format = "docx" | "xlsx" | "pptx";

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
    await parseDocx(bytes, { preloadFonts: false });
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
  if (typeof location !== "undefined" && new URLSearchParams(location.search).get("testPptxEditorFonts") === "1") {
    return { bytes, format, fonts: [] };
  }
  const workerFonts = typeof Worker === "function" && typeof OffscreenCanvas === "function";
  const decoder = workerFonts
    ? import("@betteroffice/pptx/worker").then(module => module.preparePresentationFontDecoder(bytes))
    : import("@betteroffice/pptx").then(async module => {
      await module.initWasm();
      return { decoder: module.decodeFontFace,
        families: Promise.resolve(["Arial"]), release() {} };
    });
  void decoder.catch(() => {});
  const runtime: { decoder: FontDecoder } = {
    decoder: async (encoded, options) => (await decoder).decoder(encoded, options),
  };
  const load = (family: string) => [false, true].flatMap(bold => [false, true].map(async italic => {
    const face = resolveMetricCompatFace(family, bold, italic) ?? resolveLastResortFace(family, bold, italic);
    return { family, bold, italic, bytes: new Uint8Array(await loadBundledFontBytes(face, undefined, runtime)) };
  }));
  const primary = Promise.allSettled(load("Arial"));
  const families = await decoder.then(prepared => prepared.families).catch(() => []);
  const additional = [...new Set(families)].filter(family => family !== "Arial").flatMap(load);
  const results = [...await primary, ...await Promise.allSettled(additional)];
  const fonts = results.flatMap(result => result.status === "fulfilled" ? [result.value] : []);
  const failed = results.length - fonts.length;
  if (failed) console.warn(`${failed} font faces failed to load; text uses fallback fonts.`);
  try {
    if (workerFonts) {
      const { openPresentationSession } = await import("@betteroffice/pptx/worker");
      const session = await openPresentationSession(bytes, { fonts });
      await session.dispose();
    } else {
      const { initWasm, openPresentation } = await import("@betteroffice/pptx");
      await initWasm();
      openPresentation(bytes, { fonts }).dispose();
    }
  } catch (error) {
    void decoder.then(prepared => prepared.release(), () => {});
    throw error;
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

export interface EditorProps {
  file: PreparedFile;
  name: string;
  room: Room | null;
  relayOrigin: string;
  mode?: "editing" | "viewing";
  readOnly?: boolean;
  onSave(bytes: Uint8Array): void;
  onChange(): void;
  onError(error: Error): void;
  onReady(api: EditorApi | null): void;
  onOpen(file: File): Promise<void>;
  onStatus(status: string): void;
  i18n?: EditorTranslations;
  messages?: Partial<EditorMessages>;
}

export default function Editor({
  file,
  name,
  room,
  relayOrigin,
  mode,
  readOnly,
  onSave,
  onChange,
  onError,
  onReady,
  onOpen,
  onStatus,
  i18n,
  messages,
}: EditorProps) {
  const text = useRef(DEFAULT_EDITOR_MESSAGES);
  text.current = { ...DEFAULT_EDITOR_MESSAGES, ...messages };
  const serialization = useRef<{ failure: Error | null; bytes?: Uint8Array } | null>(null);
  const serializations = useRef<Promise<unknown>>(Promise.resolve());
  const docx = useRef<DocxEditorRef>(null);
  const xlsx = useRef<XlsxAsyncEditorApi | null>(null);
  const pptx = useRef<PptxAsyncEditorApi | null>(null);
  const apiWaiter = useRef<{ resolve(): void; reject(error: Error): void } | null>(null);
  const apiFailure = useRef<Error | null>(null);
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
    const transport = createRoomTransport(relayOrigin, room.id);
    const options = { user: { name: room.name } };
    const next =
      file.format === "docx"
        ? new DocxProvider(replica!, transport, options)
        : file.format === "xlsx"
        ? new XlsxProvider(replica!, transport, options)
        : new PptxProvider(presentationReplica!, transport, options);
    setProvider(next);
    const unsubscribe = next.onStatus(({ status, synced }) =>
      callbacks.current.onStatus(
        synced || status === "connected"
          ? text.current.connected
          : status === "connecting"
          ? text.current.connecting
          : text.current.disconnected
      )
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
  }, [file.format, replica, presentationReplica, room, relayOrigin]);

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
    (api: XlsxAsyncEditorApi) => apiReady(xlsx, api, text.current.workbookDisposed),
    [apiReady]
  );
  const presentationReady = useCallback(
    (api: PptxAsyncEditorApi) => apiReady(pptx, api, text.current.presentationDisposed),
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
      file.format === "xlsx" ? text.current.workbookDisposed : text.current.presentationDisposed
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
            apiUnavailable(new Error(text.current.stillOpening));
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
          if (!active || !docx.current) throw new Error(text.current.documentDisposed);
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
            return savedBytes(bytes, current.failure, text.current.stillOpening);
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
      <DocxEditor
        i18n={i18n?.docx}
        ref={docx}
        {...(legacyDocx() ? { legacySyncApi: true as const } : { legacySyncApi: false as const })}
        trackedChanges={TRACKED_CHANGES}
        mode={mode}
        readOnly={readOnly}
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
    );
  if (file.format === "xlsx")
    return (
      <XlsxEditor
        i18n={i18n?.xlsx}
        file={file.bytes}
        fileName={name}
        onSave={(bytes) => {
          if (serialization.current) serialization.current.bytes = bytes;
          else saveBytes(bytes);
        }}
        onChange={changed}
        onError={error}
        onReady={workbookReady}
        {...(readOnly ? { readOnly: true as const } : {})}
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
      i18n={i18n?.pptx}
      file={file.bytes}
      fileName={name}
      fonts={file.fonts ?? []}
      onSave={saveBytes}
      onChange={changed}
      onReady={presentationReady}
      onError={error}
      {...(readOnly ? { readOnly: true as const } : {})}
      collaboration={presentationCollaboration}
    />
  );
}
