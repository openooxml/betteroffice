import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import Editor, { prepareFile, type EditorApi, type EditorView, type PreparedFile } from "../../../../shared/editor-host/Editor";
import { SessionEndedError, type SessionApi, type JoinResponse } from "./api";
import { displayUserName, randomClientId, showCloseButton, type WopiPageConfig } from "./config";
import type { PostMessageBridge, SaveOptions, WopiView } from "./postMessage";
import { closeDocument, finishClose, flushDocumentInput, recoverDocument, saveDocument } from "./documentActions";
import { createWopiTransport, type WopiTransport } from "./transport";

type SaveState = "idle" | "saving" | "saved" | "failed";
const COMPOSITION_ACTIVE_MESSAGE = "Finish typing before saving or closing the document.";
const LOCKED_ELSEWHERE_MESSAGE = "This file is being edited elsewhere. You are viewing a live read-only copy.";
const MAX_TIMER_MS = 2_147_483_647;

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
    loaded.bridge.start();
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
    recoveryName ? `Your edits were saved as ${recoveryName}.` : config.viewOnlyReason === "lockedElsewhere" ? LOCKED_ELSEWHERE_MESSAGE : null,
  );
  const [closed, setClosed] = useState(false);
  const [readOnly, setReadOnly] = useState(!loaded.join.canWrite);
  const writable = useRef(loaded.join.canWrite);
  const [localRecovery, setLocalRecovery] = useState<{ bytes: Uint8Array; name: string } | null>(null);
  const closeRecovery = useRef<"pending" | "downloaded" | null>(null);
  const recoveryVersion = useRef(0);
  const downloadedVersion = useRef(0);
  const [recoveryFailed, setRecoveryFailed] = useState(false);
  const editorApi = useRef<EditorApi | null>(null);
  const transportRef = useRef<WopiTransport | null>(null);
  const editorContainer = useRef<HTMLElement | null>(null);
  const compositionActive = useRef(false);
  const readOnlyPending = useRef(false);
  const permissionController = useRef<AbortController | null>(null);
  const permissionFlush = useRef<Promise<void> | null>(null);
  const loadedSent = useRef(false);
  const modified = useRef(false);
  const changeVersion = useRef(0);
  const localVersion = useRef(0);
  const savedLocalVersion = useRef(0);
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
  const applyReadOnly = useCallback((): Promise<void> => {
    if (permissionFlush.current) return permissionFlush.current;
    if (!readOnlyPending.current) {
      setReadOnly(true);
      return Promise.resolve();
    }
    const controller = new AbortController();
    permissionController.current = controller;
    const pending = flushDocumentInput(editorApi.current, {
      signal: controller.signal,
      hasActiveComposition: () => compositionActive.current,
      blockInput: () => { if (editorContainer.current) editorContainer.current.inert = true; },
    }).then(() => {
      if (controller.signal.aborted || !mounted.current) return;
      readOnlyPending.current = false;
      setReadOnly(true);
      if (editorContainer.current && !closing.current && !recoveryController.current && !sessionEnded.current) editorContainer.current.inert = false;
      setBanner((message) => message === COMPOSITION_ACTIVE_MESSAGE ? null : message);
    }).catch((error: unknown) => {
      if (controller.signal.aborted || !mounted.current) return;
      if (editorContainer.current && compositionActive.current) editorContainer.current.inert = false;
      setBanner(error instanceof Error && error.message === "compositionActive"
        ? COMPOSITION_ACTIVE_MESSAGE
        : "Could not finish pending input. Your document is still open. Retry Close to preserve your edits.");
      throw error;
    }).finally(() => {
      if (permissionController.current === controller) {
        permissionController.current = null;
        permissionFlush.current = null;
      }
    });
    permissionFlush.current = pending;
    return pending;
  }, []);

  const setWritable = useCallback((canWrite: boolean) => {
    const next = canWrite && (config.tokenExpiresAt === null || config.tokenExpiresAt > Date.now());
    writable.current = next;
    if (next) {
      permissionController.current?.abort();
      permissionController.current = null;
      permissionFlush.current = null;
      readOnlyPending.current = false;
      if (editorContainer.current && !closing.current && !recoveryController.current && !closeRecovery.current) editorContainer.current.inert = false;
      setReadOnly(false);
    } else {
      if (compositionActive.current) {
        readOnlyPending.current = true;
        if (editorContainer.current) editorContainer.current.inert = false;
      }
      void applyReadOnly().catch(() => {});
    }
  }, [applyReadOnly, config.tokenExpiresAt]);

  useEffect(() => {
    const container = editorContainer.current;
    compositionActive.current = false;
    if (!container) return;
    const start = () => { compositionActive.current = true; };
    const end = () => {
      compositionActive.current = false;
      queueMicrotask(() => {
        if (mounted.current && readOnlyPending.current && !writable.current && !sessionEnded.current && !rejoining.current) void applyReadOnly().catch(() => {});
      });
    };
    container.addEventListener("compositionstart", start, true);
    container.addEventListener("compositionend", end, true);
    return () => {
      container.removeEventListener("compositionstart", start, true);
      container.removeEventListener("compositionend", end, true);
      compositionActive.current = false;
    };
  }, [applyReadOnly, closed]);

  const completeClose = useCallback(() => {
    if (closeCompleted.current) return;
    if (localVersion.current > Math.max(savedLocalVersion.current, downloadedVersion.current) || compositionActive.current || readOnlyPending.current) {
      setBanner("Your document is still open. Retry Close to preserve your local edits.");
      throw new Error("localRecoveryRequired");
    }
    closeCompleted.current = true;
    transportRef.current?.disconnect();
    setClosed(true);
    finishClose(config, bridge, window, () => setClosed(true));
  }, [bridge, config]);

  const offerLocalRecovery = useCallback((bytes: Uint8Array, beforeClose = false) => {
    const suffix = ` (recovered ${new Date().toISOString().replace(/[:.]/g, "-")})`;
    const dot = config.fileName.lastIndexOf(".");
    const name = dot > 0 ? `${config.fileName.slice(0, dot)}${suffix}${config.fileName.slice(dot)}` : `${config.fileName}${suffix}.${config.format}`;
    setLocalRecovery({ bytes, name });
    setBanner(beforeClose
      ? "Download your recovered document to preserve your unsaved edits and finish closing."
      : "This session has ended. Download your recovered document before reloading the page to continue editing.");
  }, [config.fileName, config.format]);

  const recover = useCallback((kind: "baseline" | "local") => {
    if (recoveryController.current) return;
    const controller = new AbortController();
    recoveryController.current = controller;
    if (compositionActive.current && editorContainer.current) editorContainer.current.inert = false;
    let serialized: Uint8Array | null = null;
    setRecoveryFailed(false);
    setBanner(kind === "baseline" ? "Preserving local edits before reopening the document..." : "This session has ended. Checking for local edits...");
    void recoverDocument(editorApi.current, {
      signal: controller.signal,
      hasActiveComposition: () => compositionActive.current,
      blockInput: () => { if (editorContainer.current) editorContainer.current.inert = true; },
      onFlushed: () => {
        readOnlyPending.current = false;
        setReadOnly(true);
      },
      hasUnsavedLocalChanges: () => localVersion.current > savedLocalVersion.current,
      hasUnacknowledgedUpdates: () => transportRef.current?.hasUnacknowledgedUpdates() ?? false,
      ...(kind === "baseline" ? { upload: (bytes: Uint8Array, signal: AbortSignal) => {
        serialized = bytes;
        return api.recover(bytes, signal);
      } } : {}),
    }).then((result) => {
      if (!mounted.current || controller.signal.aborted) return;
      if (kind === "baseline") {
        onRejoin(result?.copyName);
      } else if (result) {
        offerLocalRecovery(result.bytes);
      } else {
        setBanner("This session has ended. Reload the page to continue editing.");
      }
    }).catch((failure: unknown) => {
      if (!mounted.current || controller.signal.aborted) return;
      if (failure instanceof SessionEndedError && serialized) {
        sessionEnded.current = true;
        handleClosed.resolve();
        offerLocalRecovery(serialized);
        return;
      }
      setRecoveryFailed(true);
      if (!compositionActive.current) setReadOnly(true);
      const message = failure instanceof Error ? failure.message : "Could not preserve local edits.";
      setBanner(`${message} Your document is still open. Retry recovery to preserve your edits.`);
    }).finally(() => {
      if (recoveryController.current === controller) recoveryController.current = null;
    });
  }, [api, handleClosed, offerLocalRecovery, onRejoin]);

  const endSession = useCallback(() => {
    if (lifetime.signal.aborted || sessionEnded.current) return;
    sessionEnded.current = true;
    permissionController.current?.abort();
    handleClosed.resolve();
    transportRef.current?.disconnect();
    if (finalSaved.current && localVersion.current <= Math.max(savedLocalVersion.current, downloadedVersion.current) && !compositionActive.current && !readOnlyPending.current) completeClose();
    else recover("local");
  }, [completeClose, handleClosed, lifetime, recover]);

  const rejoin = useCallback(() => {
    if (lifetime.signal.aborted || rejoining.current || sessionEnded.current) return;
    rejoining.current = true;
    permissionController.current?.abort();
    lifetime.abort(new Error("rejoinRequired"));
    transportRef.current?.disconnect();
    recover("baseline");
  }, [lifetime, recover]);

  const requestSave = useCallback(async (options: SaveOptions = {}): Promise<void> => {
    if (sessionEnded.current || finalSaved.current) throw new Error("sessionEnded");
    if (rejoining.current) throw new Error("rejoinRequired: Recovery must finish before saving or closing.");
    lifetime.signal.throwIfAborted();
    setSaveState("saving");
    try {
      let version = changeVersion.current;
      let local = localVersion.current;
      await saveDocument(editorApi.current, api, {
        ...options,
        signal: lifetime.signal,
        canWrite: () => writable.current,
        hasActiveComposition: () => compositionActive.current,
        hasPendingResync: () => writable.current && (transportRef.current?.hasPendingResync() ?? true),
        captureBarrier: () => {
          const barrier = transportRef.current?.captureSaveBarrier() ?? null;
          version = changeVersion.current;
          local = localVersion.current;
          return barrier;
        },
        onProgress: wopiDebug,
      });
      lifetime.signal.throwIfAborted();
      savedLocalVersion.current = Math.max(savedLocalVersion.current, local);
      setSaveState(localVersion.current <= savedLocalVersion.current ? "saved" : "idle");
      setBanner((message) => message === COMPOSITION_ACTIVE_MESSAGE ? null : message);
      if (modified.current && changeVersion.current === version && localVersion.current <= savedLocalVersion.current) {
        modified.current = false;
        bridge.modified(false);
      }
    } catch (error) {
      if (error instanceof SessionEndedError) endSession();
      if (!lifetime.signal.aborted) {
        setSaveState("failed");
        if (error instanceof Error && error.message === "compositionActive") setBanner(COMPOSITION_ACTIVE_MESSAGE);
      }
      throw error;
    }
  }, [api, bridge, lifetime, endSession]);

  const onSaveRequest = useCallback(async (): Promise<boolean> => {
    await requestSave();
    return false;
  }, [requestSave]);

  const close = useCallback((): Promise<void> => {
    if (closeCompleted.current) return Promise.resolve();
    if (sessionEnded.current && !finalSaved.current) return Promise.reject(new Error("sessionEnded"));
    if (rejoining.current) return Promise.reject(new Error("rejoinRequired: Recovery must finish before saving or closing."));
    if (lifetime.signal.aborted) return Promise.reject(lifetime.signal.reason);
    if (closing.current) return closing.current;
    setSaveState("saving");
    let local = localVersion.current;
    let saveCaptured = false;
    const pending = closeDocument(editorApi.current, api, {
      signal: lifetime.signal,
      whenClosed: handleClosed.promise,
      finalSaved: finalSaved.current && localVersion.current <= Math.max(savedLocalVersion.current, downloadedVersion.current) && !compositionActive.current && !readOnlyPending.current,
      onFinalSaved: () => {
        if (saveCaptured) savedLocalVersion.current = Math.max(savedLocalVersion.current, local);
        finalSaved.current = true;
      },
      canWrite: () => writable.current,
      preserveLocalWork: async () => {
        if (permissionFlush.current) await permissionFlush.current;
        if (closeRecovery.current && recoveryVersion.current === localVersion.current && !compositionActive.current && !readOnlyPending.current) return closeRecovery.current === "downloaded";
        if (localVersion.current <= savedLocalVersion.current && !compositionActive.current && !readOnlyPending.current && !editorApi.current?.hasPendingUpdates() && !transportRef.current?.hasUnacknowledgedUpdates()) return true;
        const result = await recoverDocument(editorApi.current, {
          signal: lifetime.signal,
          hasActiveComposition: () => compositionActive.current,
          blockInput: () => { if (editorContainer.current) editorContainer.current.inert = true; },
          onFlushed: () => {
            readOnlyPending.current = false;
            if (!writable.current) setReadOnly(true);
          },
          hasUnsavedLocalChanges: () => localVersion.current > savedLocalVersion.current,
          hasUnacknowledgedUpdates: () => transportRef.current?.hasUnacknowledgedUpdates() ?? false,
        });
        lifetime.signal.throwIfAborted();
        if (!result) return true;
        recoveryVersion.current = localVersion.current;
        closeRecovery.current = "pending";
        offerLocalRecovery(result.bytes, true);
        setSaveState("idle");
        return false;
      },
      hasActiveComposition: () => compositionActive.current,
      hasPendingResync: () => writable.current && (transportRef.current?.hasPendingResync() ?? true),
      captureBarrier: () => {
        const barrier = transportRef.current?.captureSaveBarrier() ?? null;
        local = localVersion.current;
        saveCaptured = barrier !== null;
        return barrier;
      },
      onProgress: wopiDebug,
      setInputBlocked: (blocked) => {
        if (lifetime.signal.aborted) return;
        if (editorContainer.current) editorContainer.current.inert = (blocked && (writable.current || !compositionActive.current)) || sessionEnded.current || (readOnlyPending.current && !compositionActive.current);
      },
      onFailure: (error) => {
        if (lifetime.signal.aborted) return;
        if (error instanceof SessionEndedError) {
          endSession();
          return;
        }
        setSaveState("failed");
        const message = error instanceof Error ? error.message : "Could not close the document.";
        setBanner(message === "compositionActive" ? COMPOSITION_ACTIVE_MESSAGE : message);
      },
    }, completeClose).catch(() => {}).finally(() => { closing.current = null; });
    closing.current = pending;
    return pending;
  }, [api, completeClose, handleClosed, lifetime, endSession, offerLocalRecovery]);

  const downloadRecovery = useCallback(() => {
    if (!localRecovery) return;
    const url = URL.createObjectURL(new Blob([localRecovery.bytes.slice().buffer], { type: "application/octet-stream" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = localRecovery.name;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 0);
    if (closeRecovery.current === "pending") {
      closeRecovery.current = "downloaded";
      downloadedVersion.current = recoveryVersion.current;
      void close().catch(() => {});
    }
  }, [close, localRecovery]);

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      permissionController.current?.abort();
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
        if (localVersion.current <= savedLocalVersion.current) setSaveState("saved");
      } else if (event.type === "saveFailed") {
        setSaveState("failed");
      } else if (event.type === "conflict") {
        setConflictName(event.copyName || "a copy");
        setBanner(`The file changed on the host. Your edits were saved as ${event.copyName || "a copy"}.`);
      } else if (event.type === "tokenExpiring") {
        setBanner("Your access is about to expire. Save your work and reopen the document from the host.");
      } else if (event.type === "viewOnly") {
        setWritable(false);
        if (event.reason === "lockedElsewhere") {
          setBanner(LOCKED_ELSEWHERE_MESSAGE);
        }
      } else if (event.type === "snapshot") {
        setWritable(event.canWrite);
        setBanner((message) => event.viewOnlyReason === "lockedElsewhere"
          ? LOCKED_ELSEWHERE_MESSAGE
          : message === LOCKED_ELSEWHERE_MESSAGE ? null : message);
      } else if (event.type === "writable") {
        setWritable(event.canWrite);
        setBanner((message) => message === LOCKED_ELSEWHERE_MESSAGE ? null : message);
      }
    }, { baselineGeneration: loaded.join.baselineGeneration }).catch(() => undefined);
    return () => {
      lifetime.signal.removeEventListener("abort", abort);
      controller.abort();
    };
  }, [api, closed, lifetime, loaded.join.baselineGeneration, rejoin, endSession, setWritable]);

  useEffect(() => {
    const target = config.tokenExpiresAt;
    if (target === null) return;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const schedule = () => {
      const delay = target - Date.now();
      if (delay <= 0) {
        setWritable(false);
        return;
      }
      timer = setTimeout(schedule, Math.min(delay, MAX_TIMER_MS));
    };
    schedule();
    return () => { if (timer !== null) clearTimeout(timer); };
  }, [config.tokenExpiresAt, setWritable]);

  const createTransport = useCallback(
    (_roomId: string) => {
      wopiDebug({ transportCreated: true });
      const transport = createWopiTransport({ handle: config.sessionHandle, baselineGeneration: loaded.join.baselineGeneration, origin: window.location.origin, checkSession: (signal) => api.isActive(signal) });
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
          <button type="button" className="wopi-close" onClick={() => void close().catch(() => {})}>
            Close
          </button>
        ) : null}
      </header>
      {banner ? <div className="wopi-banner" role="alert">
        {banner}
        {localRecovery ? <button type="button" onClick={downloadRecovery}>Download recovered document</button> : null}
        {recoveryFailed ? <button type="button" onClick={() => recover(rejoining.current ? "baseline" : "local")}>Retry</button> : null}
      </div> : null}
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
          onLocalUpdate={() => {
            if (!mounted.current) return;
            localVersion.current += 1;
            setSaveState((state) => state === "saved" ? "idle" : state);
            transportRef.current?.noteLocalUpdate();
          }}
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
            if (!ready) compositionActive.current = false;
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
