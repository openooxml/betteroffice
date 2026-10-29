import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import Editor, { prepareFile, type EditorApi, type EditorView, type PreparedFile } from "../../../../shared/editor-host/Editor";
import type { SessionApi, JoinResponse } from "./api";
import { displayUserName, randomClientId, showCloseButton, type WopiPageConfig } from "./config";
import type { PostMessageBridge, SaveOptions, WopiView } from "./postMessage";
import { closeDocument, finishClose, saveDocument } from "./documentActions";
import { createWopiTransport, type WopiTransport } from "./transport";

type SaveState = "idle" | "saving" | "saved" | "failed";

/** Records page progress for diagnostics. */
export function wopiDebug(patch: Record<string, unknown>): void {
  const holder = window as unknown as { __wopiDebug?: Record<string, unknown> };
  holder.__wopiDebug = { ...(holder.__wopiDebug ?? {}), ...patch };
}

export interface Loaded {
  config: WopiPageConfig;
  api: SessionApi;
  bridge: PostMessageBridge;
  file: PreparedFile;
  join: JoinResponse;
  clientId: number;
  actions: {
    save?: (options?: SaveOptions) => Promise<void>;
    close?: () => Promise<void>;
    grabFocus?: () => void;
    views: WopiView[];
  };
}

export async function loadBaseline(api: SessionApi, signal?: AbortSignal): Promise<Pick<Loaded, "file" | "join" | "clientId">> {
  for (;;) {
    signal?.throwIfAborted();
    const source = await api.source(signal);
    signal?.throwIfAborted();
    wopiDebug({ step: "source", sourceBytes: source.bytes.byteLength });
    const join = await api.join(signal);
    signal?.throwIfAborted();
    if (source.baselineGeneration !== join.baselineGeneration) continue;
    wopiDebug({ step: "join", joinFormat: join.format, initialUpdateBytes: join.initialUpdate.byteLength });
    const file = await prepareFile(source.bytes, join.format, { fontBaseUrl: "/fonts/" });
    signal?.throwIfAborted();
    wopiDebug({ step: "prepared", preparedFormat: file.format });
    return { file, join, clientId: randomClientId() };
  }
}

export default function WopiApp({ loaded }: { loaded: Loaded }) {
  const [current, setCurrent] = useState<Loaded | null>(loaded);
  const [error, setError] = useState<string | null>(null);
  const [recoveryName, setRecoveryName] = useState<string | null>(null);
  const rejoin = useCallback((copyName?: string) => { setRecoveryName(copyName ?? null); setCurrent(null); }, []);

  useEffect(() => {
    if (current) return;
    const controller = new AbortController();
    loaded.bridge.frameReady();
    void loadBaseline(loaded.api, controller.signal).then((baseline) => {
      if (controller.signal.aborted) return;
      loaded.actions.views = [{ ViewId: baseline.clientId, UserName: displayUserName(loaded.config.userFriendlyName) }];
      setCurrent({ ...loaded, ...baseline });
    }).catch((failure: unknown) => {
      if (controller.signal.aborted) return;
      const message = failure instanceof Error ? failure.message : "Could not reopen the document.";
      wopiDebug({ step: "rejoinError", rejoinError: message });
      loaded.bridge.bootFailed();
      setError(message);
    });
    return () => controller.abort();
  }, [current, loaded]);

  if (current) return <WopiEditor loaded={current} onRejoin={rejoin} recoveryName={recoveryName} />;
  if (error) return <div className="wopi-error" role="alert">{error}</div>;
  return <div className="wopi-shell" role="status">Reopening document...</div>;
}

function toWopiViews(views: EditorView[], self: EditorView, friendlyName: string): WopiView[] {
  const all = new Map<number, WopiView>();
  all.set(self.clientId, { ViewId: self.clientId, UserName: friendlyName });
  for (const view of views) {
    if (!all.has(view.clientId)) all.set(view.clientId, { ViewId: view.clientId, ...(view.name ? { UserName: view.name } : {}) });
  }
  return [...all.values()];
}

function WopiEditor({ loaded, onRejoin, recoveryName }: { loaded: Loaded; onRejoin(copyName?: string): void; recoveryName: string | null }) {
  const { config, api, bridge } = loaded;
  const [saveState, setSaveState] = useState<SaveState>("idle");
  const [conflictName, setConflictName] = useState<string | null>(null);
  const [banner, setBanner] = useState<string | null>(
    recoveryName ? `Your edits were saved as ${recoveryName}.` : config.viewOnlyReason === "lockedElsewhere" ? "This file is being edited elsewhere. You are viewing a live read-only copy." : null,
  );
  const [closed, setClosed] = useState(false);
  const [readOnly, setReadOnly] = useState(!config.canWrite);
  const editorApi = useRef<EditorApi | null>(null);
  const transportRef = useRef<WopiTransport | null>(null);
  const editorContainer = useRef<HTMLElement | null>(null);
  const loadedSent = useRef(false);
  const modified = useRef(false);
  const changeVersion = useRef(0);
  const closing = useRef<Promise<void> | null>(null);
  const finalSaved = useRef(false);
  const closeCompleted = useRef(false);
  const sessionEnded = useRef(false);
  const rejoining = useRef(false);
  const recoveryController = useRef<AbortController | null>(null);
  const handleClosed = useMemo(() => {
    let resolve!: () => void;
    const promise = new Promise<void>((done) => { resolve = done; });
    return { promise, resolve };
  }, []);
  const mounted = useRef(true);
  const lifetime = useRef(new AbortController()).current;
  const userName = displayUserName(config.userFriendlyName);
  const room = useMemo(() => ({ id: "wopi", name: userName }), [userName]);
  const views = useRef<WopiView[]>([{ ViewId: loaded.clientId, UserName: userName }]);

  const completeClose = useCallback(() => {
    if (closeCompleted.current) return;
    closeCompleted.current = true;
    transportRef.current?.disconnect();
    setClosed(true);
    finishClose(config, bridge, window, () => setClosed(true));
  }, [bridge, config]);

  const endSession = useCallback(() => {
    if (lifetime.signal.aborted) return;
    sessionEnded.current = true;
    handleClosed.resolve();
    transportRef.current?.disconnect();
    if (editorContainer.current) editorContainer.current.inert = true;
    setReadOnly(true);
    if (finalSaved.current) completeClose();
    else setBanner("This session has ended. Reload the page to continue editing.");
  }, [completeClose, handleClosed, lifetime]);

  const rejoin = useCallback(() => {
    if (lifetime.signal.aborted || rejoining.current) return;
    rejoining.current = true;
    if (editorContainer.current) editorContainer.current.inert = true;
    setReadOnly(true);
    lifetime.abort(new Error("rejoinRequired"));
    const oldTransport = transportRef.current;
    oldTransport?.disconnect();
    const editor = editorApi.current;
    const controller = new AbortController();
    recoveryController.current = controller;
    void (async () => {
      await editor?.flushPendingInput();
      controller.signal.throwIfAborted();
      let copyName: string | undefined;
      if (editor?.hasPendingUpdates() || oldTransport?.hasUnacknowledgedUpdates()) {
        if (!editor) throw new Error("The editor is not ready to preserve local work. Reload the page to retry.");
        const bytes = await editor.serialize();
        controller.signal.throwIfAborted();
        copyName = (await api.recover(bytes, controller.signal)).copyName;
      }
      controller.signal.throwIfAborted();
      onRejoin(copyName);
    })().catch((failure: unknown) => {
      if (!mounted.current || controller.signal.aborted) return;
      setBanner(failure instanceof Error ? failure.message : "Could not save a recovery copy. Reload the page to retry.");
    });
  }, [api, lifetime, onRejoin]);

  const requestSave = useCallback(async (options: SaveOptions = {}): Promise<void> => {
    lifetime.signal.throwIfAborted();
    if (sessionEnded.current || finalSaved.current) throw new Error("sessionEnded");
    setSaveState("saving");
    try {
      let version = changeVersion.current;
      await saveDocument(editorApi.current, api, {
        ...options,
        signal: lifetime.signal,
        hasPendingResync: () => config.canWrite && (transportRef.current?.hasPendingResync() ?? true),
        captureBarrier: () => transportRef.current?.captureSaveBarrier() ?? null,
        onFlushed: () => { version = changeVersion.current; },
        onProgress: wopiDebug,
      });
      lifetime.signal.throwIfAborted();
      setSaveState("saved");
      if (modified.current && changeVersion.current === version) {
        modified.current = false;
        bridge.modified(false);
      }
    } catch (error) {
      if (!lifetime.signal.aborted) setSaveState("failed");
      throw error;
    }
  }, [api, bridge, config.canWrite, lifetime]);

  const onSaveRequest = useCallback(async (): Promise<boolean> => {
    await requestSave();
    return false;
  }, [requestSave]);

  const close = useCallback((): Promise<void> => {
    if (closeCompleted.current) return Promise.resolve();
    if (lifetime.signal.aborted) return Promise.reject(lifetime.signal.reason);
    if (closing.current) return closing.current;
    setSaveState("saving");
    const pending = closeDocument(editorApi.current, api, {
      signal: lifetime.signal,
      whenClosed: handleClosed.promise,
      finalSaved: finalSaved.current,
      onFinalSaved: () => { finalSaved.current = true; },
      hasPendingResync: () => config.canWrite && (transportRef.current?.hasPendingResync() ?? true),
      captureBarrier: () => transportRef.current?.captureSaveBarrier() ?? null,
      onProgress: wopiDebug,
      setInputBlocked: (blocked) => {
        if (lifetime.signal.aborted) return;
        if (editorContainer.current) editorContainer.current.inert = blocked || sessionEnded.current;
      },
      onFailure: (error) => {
        if (lifetime.signal.aborted) return;
        setSaveState("failed");
        setBanner(error instanceof Error ? error.message : "Could not close the document.");
      },
    }, completeClose).catch(() => {}).finally(() => { closing.current = null; });
    closing.current = pending;
    return pending;
  }, [api, completeClose, config.canWrite, handleClosed, lifetime]);

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      recoveryController.current?.abort();
      lifetime.abort(new Error("editorDisposed"));
    };
  }, [lifetime]);

  useEffect(() => {
    loaded.actions.save = requestSave;
    loaded.actions.close = close;
    loaded.actions.grabFocus = () => editorApi.current?.focus();
    loaded.actions.views = views.current;
  }, [loaded, requestSave, close]);

  useEffect(() => {
    if (closed) return;
    wopiDebug({ editorMounted: true, bridgeStarted: true, bridgeOrigin: config.postMessageOrigin, bridgeEnabled: bridge.enabled });
    bridge.start();
    return () => bridge.dispose();
  }, [bridge, config.postMessageOrigin, closed]);

  useEffect(() => {
    if (closed) return;
    const controller = new AbortController();
    const abort = () => controller.abort();
    lifetime.signal.addEventListener("abort", abort, { once: true });
    if (lifetime.signal.aborted) controller.abort();
    void api.events(controller.signal, (event) => {
      if (lifetime.signal.aborted) return;
      if (event.type === "closed") {
        endSession();
      } else if (event.type === "rejoinRequired") {
        rejoin();
      } else if (event.type === "saved") {
        setSaveState("saved");
      } else if (event.type === "saveFailed") {
        setSaveState("failed");
      } else if (event.type === "conflict") {
        setConflictName(event.copyName || "a copy");
        setBanner(`The file changed on the host. Your edits were saved as ${event.copyName || "a copy"}.`);
      } else if (event.type === "tokenExpiring") {
        setBanner("Your access is about to expire. Save your work and reopen the document from the host.");
      } else if (event.type === "viewOnly") {
        setReadOnly(true);
        if (event.reason === "lockedElsewhere") {
          setBanner("This file is being edited elsewhere. You are viewing a live read-only copy.");
        }
      }
    }, { baselineGeneration: loaded.join.baselineGeneration }).catch(() => undefined);
    return () => {
      lifetime.signal.removeEventListener("abort", abort);
      controller.abort();
    };
  }, [api, closed, lifetime, loaded.join.baselineGeneration, rejoin, endSession]);

  useEffect(() => {
    if (config.tokenExpiresAt === null) return;
    const delay = config.tokenExpiresAt - Date.now();
    if (delay <= 0) {
      setReadOnly(true);
      return;
    }
    const timer = setTimeout(() => setReadOnly(true), delay);
    return () => clearTimeout(timer);
  }, [config.tokenExpiresAt]);

  const createTransport = useCallback(
    (_roomId: string) => {
      wopiDebug({ transportCreated: true });
      const transport = createWopiTransport({ handle: config.sessionHandle, baselineGeneration: loaded.join.baselineGeneration, origin: window.location.origin, checkSession: () => api.isActive(lifetime.signal) });
      transportRef.current = transport;
      return {
        ...transport,
        connect: () => { if (!lifetime.signal.aborted) transport.connect(); },
        send: (data: Uint8Array) => !lifetime.signal.aborted && transport.send(data),
        onEvent: (listener: Parameters<typeof transport.onEvent>[0]) =>
          transport.onEvent((event) => {
            if (lifetime.signal.aborted) return;
            if (event.type === "close" && event.code === 4401) {
              endSession();
            }
            if (event.type === "close" && event.code === 4409) {
              rejoin();
              return;
            }
            if (event.type === "open") wopiDebug({ socketOpen: true });
            listener(event);
          }),
      };
    },
    [api, config.sessionHandle, loaded.join.baselineGeneration, lifetime, rejoin, endSession],
  );

  const onViewsChange = useCallback(
    (next: EditorView[]) => {
      if (lifetime.signal.aborted) return;
      const previous = new Map(views.current.map((view) => [view.ViewId, view]));
      const merged = toWopiViews(next, { clientId: loaded.clientId }, userName);
      const current = new Map(merged.map((view) => [view.ViewId, view]));
      for (const view of merged) if (!previous.has(view.ViewId)) bridge.viewAdded(view);
      for (const view of views.current) if (!current.has(view.ViewId)) bridge.viewRemoved(view);
      views.current = merged;
      loaded.actions.views = merged;
    },
    [bridge, userName, loaded, lifetime],
  );

  if (closed) return <div className="wopi-shell" role="status">Document closed.</div>;

  return (
    <div className="wopi-shell">
      <header className="wopi-bar">
        <span className="wopi-name">{config.fileName}</span>
        <span className="wopi-status" role="status">
          {saveState === "saving" ? "Saving..." : saveState === "saved" ? "Saved" : saveState === "failed" ? "Save failed" : ""}
        </span>
        {showCloseButton(config) ? (
          <button type="button" className="wopi-close" onClick={() => void close()}>
            Close
          </button>
        ) : null}
      </header>
      {banner ? <div className="wopi-banner" role="alert">{banner}</div> : null}
      {conflictName ? <div className="wopi-banner" role="alert">Saved as a conflict copy: {conflictName}</div> : null}
      <main className="wopi-editor" ref={editorContainer}>
        <Editor
          file={loaded.file}
          name={config.fileName}
          room={room}
          createTransport={createTransport}
          initialUpdate={loaded.join.initialUpdate}
          collaborationClientId={loaded.clientId}
          readOnly={readOnly}
          downloadOnSave={false}
          onSaveRequest={onSaveRequest}
          onSave={() => undefined}
          onLocalUpdate={() => { if (mounted.current) transportRef.current?.noteLocalUpdate(); }}
          onChange={() => {
            if (lifetime.signal.aborted) return;
            changeVersion.current += 1;
            if (!modified.current) {
              modified.current = true;
              bridge.modified(true);
            }
          }}
          onError={(error) => {
            if (lifetime.signal.aborted) return;
            if (!loadedSent.current) bridge.bootFailed();
            setBanner(error.message);
          }}
          onReady={(ready) => {
            if (lifetime.signal.aborted) return;
            wopiDebug({ onReadyCalled: true, onReadyHasApi: ready !== null });
            editorApi.current = ready;
            if (ready && !loadedSent.current) {
              void ready.whenReady().then(() => {
                if (!mounted.current || editorApi.current !== ready || loadedSent.current) return;
                loadedSent.current = true;
                bridge.documentLoaded();
              }).catch((error: unknown) => {
                if (!mounted.current || editorApi.current !== ready) return;
                bridge.bootFailed();
                setBanner(error instanceof Error ? error.message : "Could not open the document.");
              });
            }
          }}
          onOpen={async () => undefined}
          onStatus={() => undefined}
          onViewsChange={onViewsChange}
        />
      </main>
    </div>
  );
}
