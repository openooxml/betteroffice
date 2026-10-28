import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { DocxEditor, type DocxEditorRef, type DocxErrorInfo, type TrackedChangesOptions } from "@betteroffice/docx-react";
import { XlsxEditor, type XlsxAsyncEditorApi } from "@betteroffice/xlsx-react";
import { PptxEditor, type PptxAsyncEditorApi } from "@betteroffice/pptx-react";
import { configureDefaultFonts } from "@betteroffice/docx/layout";
import { setGoogleFontsEnabled } from "@betteroffice/docx/utils";
import { CollaborationProvider as DocxProvider, type AsyncReplicaAdapter } from "@betteroffice/docx/collaboration";
import { CollaborationProvider as XlsxProvider } from "@betteroffice/xlsx/collaboration";
import { CollaborationProvider as PptxProvider, type AsyncCollaborationReplica } from "@betteroffice/pptx/collaboration";
import type { PptxFontFace } from "@betteroffice/pptx";
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
} from "@betteroffice/fonts";
import { createRoomTransport } from "./collaboration/createRoomTransport";
import type { CollaborationReplica, CollaborationTransport } from "./collaboration/types";
import { savedBytes } from "./session";
import "@betteroffice/docx-react/styles.css";

const EDITOR_READY_TIMEOUT_MS = 30_000;
const WORKBOOK_DISPOSED = "The workbook editor was disposed.";
const PRESENTATION_DISPOSED = "The presentation editor was disposed.";

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
  format: Format,
  options?: { fontBaseUrl?: string }
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
  const results = await Promise.allSettled(
    [false, true].flatMap((bold) =>
      [false, true].map(async (italic) => {
        const face =
          resolveMetricCompatFace("Arial", bold, italic) ??
          resolveLastResortFace("Arial", bold, italic);
        return {
          family: "Arial",
          bold,
          italic,
          bytes: new Uint8Array(await loadBundledFontBytes(face, options?.fontBaseUrl ? { baseUrl: options.fontBaseUrl } : undefined)),
        };
      })
    )
  );
  const fonts = results.flatMap((result) => result.status === "fulfilled" ? [result.value] : []);
  const failed = results.length - fonts.length;
  if (failed) console.warn(`${failed} font faces failed to load; text uses fallback fonts.`);
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
  flushPendingInput(): Promise<void>;
  hasPendingUpdates(): boolean;
  whenReady(): Promise<void>;
  encodeStateVector(): Promise<Uint8Array>;
  serialize(): Promise<Uint8Array>;
  focus(): void;
  print(): void;
}
export interface Room {
  id: string;
  name: string;
}

/** One collaboration view, for presence reporting. */
export interface EditorView {
  clientId: number;
  name?: string;
}

export interface EditorProps {
  file: PreparedFile;
  name: string;
  room: Room | null;
  relayOrigin?: string;
  createTransport?: (roomId: string) => CollaborationTransport;
  initialUpdate?: Uint8Array;
  collaborationClientId?: number;
  mode?: "editing" | "viewing";
  readOnly?: boolean;
  onSave(bytes: Uint8Array): void;
  onSaveRequest?: () => boolean | void | Promise<boolean | void>;
  downloadOnSave?: boolean;
  onChange(): void;
  onError(error: Error): void;
  onReady(api: EditorApi | null): void;
  onOpen(file: File): Promise<void>;
  onStatus(status: string): void;
  onViewsChange?(views: EditorView[]): void;
}

function viewsFromPeers(peers: readonly unknown[]): EditorView[] {
  const views: EditorView[] = [];
  for (const peer of peers) {
    if (typeof peer !== "object" || peer === null) continue;
    const record = peer as { clientId?: unknown; state?: unknown; user?: unknown };
    const state = typeof record.state === "object" && record.state !== null
      ? (record.state as { clientId?: unknown; user?: unknown })
      : undefined;
    const clientId = typeof record.clientId === "number" ? record.clientId
      : state && typeof state.clientId === "number" ? state.clientId : undefined;
    if (clientId === undefined) continue;
    const user = (record.user ?? state?.user) as { name?: unknown } | undefined;
    views.push({ clientId, ...(typeof user?.name === "string" ? { name: user.name } : {}) });
  }
  return views;
}

export default function Editor({
  file,
  name,
  room,
  relayOrigin,
  createTransport,
  initialUpdate,
  collaborationClientId,
  mode,
  readOnly,
  onSave,
  onSaveRequest,
  downloadOnSave,
  onChange,
  onError,
  onReady,
  onOpen,
  onStatus,
  onViewsChange,
}: EditorProps) {
  const serialization = useRef<{ failure: Error | null; bytes?: Uint8Array } | null>(null);
  const serializations = useRef<Promise<unknown>>(Promise.resolve());
  const docx = useRef<DocxEditorRef>(null);
  const xlsx = useRef<XlsxAsyncEditorApi | null>(null);
  const pptx = useRef<PptxAsyncEditorApi | null>(null);
  const apiWaiter = useRef<{ resolve(): void; reject(error: Error): void } | null>(null);
  const apiFailure = useRef<Error | null>(null);
  const generatedClientId = useRef(
    crypto.getRandomValues(new Uint32Array(1))[0] & 0x7fffffff || 1
  ).current;
  const clientId = collaborationClientId ?? generatedClientId;
  const [replica, setReplica] = useState<CollaborationReplica | AsyncReplicaAdapter | null>(null);
  const [presentationReplica, setPresentationReplica] = useState<AsyncCollaborationReplica | null>(null);
  const [provider, setProvider] = useState<
    DocxProvider | XlsxProvider | PptxProvider | null
  >(null);
  const providerRef = useRef(provider);
  const replicas = useRef({ replica, presentationReplica });
  replicas.current = { replica, presentationReplica };
  const [painted, setPainted] = useState(false);
  const firstPainted = useCallback(() => setPainted(true), []);
  const presentationPresence = provider instanceof PptxProvider ? provider : undefined;
  const presentationCollaboration = useMemo(() => room ? {
    roomId: room.id,
    clientId,
    ...(initialUpdate ? { initialUpdate } : {}),
    onReplica: setPresentationReplica,
    presence: presentationPresence,
  } : undefined, [room?.id, clientId, initialUpdate, presentationPresence]);
  const callbacks = useRef({ onSave, onChange, onError, onStatus, onReady, onViewsChange });
  callbacks.current = { onSave, onChange, onError, onStatus, onReady, onViewsChange };

  useEffect(() => {
    if (!provider || !callbacks.current.onViewsChange) return;
    const report = (peers: readonly unknown[]) => callbacks.current.onViewsChange?.(viewsFromPeers(peers));
    const candidate = provider as unknown as {
      onPeers?: (listener: (peers: readonly unknown[]) => void) => () => void;
      onAwareness?: (listener: (peers: readonly unknown[]) => void) => () => void;
      onPresence?: (listener: (peers: readonly unknown[]) => void) => () => void;
    };
    const subscribe = candidate.onPeers ?? candidate.onAwareness ?? candidate.onPresence;
    if (!subscribe) return;
    return subscribe.call(provider, report);
  }, [provider]);

  useEffect(() => {
    if (!room || (file.format === "pptx" ? !presentationReplica : !replica)) return;
    const transport = createTransport ? createTransport(room.id) : createRoomTransport(relayOrigin ?? "", room.id);
    const options = { user: { name: room.name } };
    const next =
      file.format === "docx"
        ? new DocxProvider(replica!, transport, options)
        : file.format === "xlsx"
        ? new XlsxProvider(replica!, transport, options)
        : new PptxProvider(presentationReplica!, transport, options);
    providerRef.current = next;
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
      if (providerRef.current === next) providerRef.current = null;
      setProvider(null);
    };
  }, [file.format, replica, presentationReplica, room, relayOrigin, createTransport]);

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
    if (file.format === "docx" && !painted) return;
    let active = true;
    const disposed = new Error(
      file.format === "xlsx" ? WORKBOOK_DISPOSED : PRESENTATION_DISPOSED
    );
    let waiting: Promise<void> | null = null;
    const whenReady = async <Api,>(ref: { current: Api | null }): Promise<Api> => {
      if (!active) throw disposed;
      let api = ref.current;
      if (!api) {
        if (apiFailure.current) throw apiFailure.current;
        let timeout: ReturnType<typeof setTimeout> | undefined;
        if (!waiting) waiting = new Promise<void>((resolve, reject) => {
          const waiter = { resolve, reject };
          apiWaiter.current = waiter;
          timeout = setTimeout(() => {
            if (apiWaiter.current !== waiter) return;
            apiUnavailable(new Error(
              "The editor is still opening the file. Try again in a moment."
            ));
          }, EDITOR_READY_TIMEOUT_MS);
        }).finally(() => {
          clearTimeout(timeout);
          waiting = null;
        });
        await waiting;
        api = ref.current;
      }
      if (!active || !api || ref.current !== api) throw disposed;
      return api;
    };
    callbacks.current.onReady({
      hasPendingUpdates() {
        if (!active) throw disposed;
        const current = providerRef.current;
        return !current || current.pendingBytes > 0;
      },
      async flushPendingInput() {
        if (file.format === "docx") {
          if (!active || !docx.current) throw new Error("The document editor was disposed.");
          await docx.current.flushPendingInput();
        } else if (file.format === "xlsx") {
          await (await whenReady(xlsx)).version();
        } else {
          await (await whenReady(pptx)).flushPendingInput();
        }
      },
      async whenReady() {
        if (file.format === "xlsx") await (await whenReady(xlsx)).whenHydrated();
        else if (file.format === "pptx") await (await whenReady(pptx)).whenHydrated();
        if (!active) throw disposed;
      },
      async encodeStateVector() {
        if (!active) throw disposed;
        if (file.format === "docx") {
          const current = replicas.current.replica;
          if (!current) throw new Error("The collaboration replica is not ready.");
          return await ("async" in current ? current.async : current).encodeStateVector();
        }
        const current = file.format === "xlsx"
          ? (await whenReady(xlsx)).collaboration
          : replicas.current.presentationReplica;
        if (!current) throw new Error("The collaboration replica is not ready.");
        return await current.encodeStateVector();
      },
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
      <DocxEditor
        ref={docx}
        {...(legacyDocx() ? { legacySyncApi: true as const } : { legacySyncApi: false as const })}
        trackedChanges={TRACKED_CHANGES}
        mode={mode}
        readOnly={readOnly}
        documentBuffer={file.bytes}
        downloadOnSave={downloadOnSave ?? false}
        {...(onSaveRequest ? { onSaveRequest } : {})}
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
                ...(initialUpdate ? { initialUpdate } : {}),
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
        file={file.bytes}
        fileName={name}
        {...(onSaveRequest ? { onSaveRequest } : {})}
        {...(downloadOnSave !== undefined ? { downloadOnSave } : {})}
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
                ...(initialUpdate ? { initialUpdate } : {}),
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
      {...(onSaveRequest ? { onSaveRequest } : {})}
      {...(downloadOnSave !== undefined ? { downloadOnSave } : {})}
      onSave={saveBytes}
      onChange={changed}
      onReady={presentationReady}
      onError={error}
      {...(readOnly ? { readOnly: true as const } : {})}
      collaboration={presentationCollaboration}
    />
  );
}
