import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";

import { configureDefaultFonts } from "@betteroffice/docx/layout";
import Editor, { prepareFile, type EditorApi, type EditorView, type PreparedFile } from "../../../../shared/editor-host/Editor";
import { SessionApi, type JoinResponse } from "./api";
import { displayUserName, randomClientId, readWopiConfig, showCloseButton, type WopiPageConfig } from "./config";
import { PostMessageBridge, type SaveOptions, type WopiView } from "./postMessage";
import { closeDocument, finishClose, saveDocument } from "./documentActions";
import { createWopiTransport } from "./transport";
import "./styles.css";

type SaveState = "idle" | "saving" | "saved" | "failed";

/** Records page progress for diagnostics. */
export function wopiDebug(patch: Record<string, unknown>): void {
  const holder = window as unknown as { __wopiDebug?: Record<string, unknown> };
  holder.__wopiDebug = { ...(holder.__wopiDebug ?? {}), ...patch };
}

interface Loaded {
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

function toWopiViews(views: EditorView[], self: EditorView, friendlyName: string): WopiView[] {
  const all = new Map<number, WopiView>();
  all.set(self.clientId, { ViewId: self.clientId, UserName: friendlyName });
  for (const view of views) {
    if (!all.has(view.clientId)) all.set(view.clientId, { ViewId: view.clientId, ...(view.name ? { UserName: view.name } : {}) });
  }
  return [...all.values()];
}

function WopiApp({ loaded }: { loaded: Loaded }) {
  const { config, api, bridge } = loaded;
  const [saveState, setSaveState] = useState<SaveState>("idle");
  const [conflictName, setConflictName] = useState<string | null>(null);
  const [banner, setBanner] = useState<string | null>(
    config.viewOnlyReason === "lockedElsewhere" ? "This file is being edited elsewhere. You are viewing a live read-only copy." : null,
  );
  const [closed, setClosed] = useState(false);
  const [readOnly, setReadOnly] = useState(!config.canWrite);
  const editorApi = useRef<EditorApi | null>(null);
  const editorContainer = useRef<HTMLElement | null>(null);
  const loadedSent = useRef(false);
  const modified = useRef(false);
  const changeVersion = useRef(0);
  const closing = useRef<Promise<void> | null>(null);
  const mounted = useRef(true);
  const userName = displayUserName(config.userFriendlyName);
  const room = useMemo(() => ({ id: "wopi", name: userName }), [userName]);
  const views = useRef<WopiView[]>([{ ViewId: loaded.clientId, UserName: userName }]);

  const requestSave = useCallback(async (options: SaveOptions = {}): Promise<void> => {
    setSaveState("saving");
    try {
      let version = changeVersion.current;
      await saveDocument(editorApi.current, api, {
        ...options,
        isModified: () => modified.current,
        onFlushed: () => { version = changeVersion.current; },
        onProgress: wopiDebug,
      });
      setSaveState("saved");
      if (modified.current && changeVersion.current === version) {
        modified.current = false;
        bridge.modified(false);
      }
    } catch (error) {
      setSaveState("failed");
      throw error;
    }
  }, [api, bridge]);

  const onSaveRequest = useCallback(async (): Promise<boolean> => {
    await requestSave();
    return false;
  }, [requestSave]);

  const close = useCallback((): Promise<void> => {
    if (closing.current) return closing.current;
    const attempt = async () => {
      const editor = editorApi.current;
      if (!editor) throw new Error("editorNotReady");
      if (editorContainer.current) editorContainer.current.inert = true;
      await closeDocument(editor, {
        save: async () => {
          do { await requestSave({ DontTerminateEdit: true }); } while (modified.current);
        },
        close: () => api.close(),
      }, () => modified.current, () => {
        setClosed(true);
        finishClose(config, bridge, window, () => setClosed(true));
      });
    };
    const pending = attempt().catch((error: unknown) => {
      if (editorContainer.current) editorContainer.current.inert = false;
      setSaveState("failed");
      setBanner(error instanceof Error ? error.message : "Could not close the document.");
    }).finally(() => { closing.current = null; });
    closing.current = pending;
    return pending;
  }, [api, bridge, config, requestSave]);

  useEffect(() => {
    mounted.current = true;
    return () => { mounted.current = false; };
  }, []);

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
    void api.events(controller.signal, (event) => {
      if (event.type === "saved") {
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
    }).catch(() => undefined);
    return () => controller.abort();
  }, [api, bridge, closed]);

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
      const transport = createWopiTransport({ handle: config.sessionHandle, origin: window.location.origin });
      return {
        ...transport,
        onEvent: (listener: Parameters<typeof transport.onEvent>[0]) =>
          transport.onEvent((event) => {
            if (event.type === "open") wopiDebug({ socketOpen: true });
            listener(event);
          }),
      };
    },
    [config.sessionHandle],
  );

  const onViewsChange = useCallback(
    (next: EditorView[]) => {
      const previous = new Map(views.current.map((view) => [view.ViewId, view]));
      const merged = toWopiViews(next, { clientId: loaded.clientId }, userName);
      const current = new Map(merged.map((view) => [view.ViewId, view]));
      for (const view of merged) if (!previous.has(view.ViewId)) bridge.viewAdded(view);
      for (const view of views.current) if (!current.has(view.ViewId)) bridge.viewRemoved(view);
      views.current = merged;
      loaded.actions.views = merged;
    },
    [bridge, userName, loaded],
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
          onChange={() => {
            changeVersion.current += 1;
            if (!modified.current) {
              modified.current = true;
              bridge.modified(true);
            }
          }}
          onError={(error) => {
            if (!loadedSent.current) bridge.bootFailed();
            setBanner(error.message);
          }}
          onReady={(ready) => {
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

async function boot(): Promise<void> {
  wopiDebug({ step: "boot" });
  const config = readWopiConfig(document);
  wopiDebug({ step: "config", format: config.format, postMessageOrigin: config.postMessageOrigin, canWrite: config.canWrite });
  const api = new SessionApi(config.sessionHandle);
  const actions: Loaded["actions"] = { views: [{ ViewId: 0, UserName: displayUserName(config.userFriendlyName) }] };
  const bridge = new PostMessageBridge(window, config.postMessageOrigin, {
    onSave: (options) => {
      if (!actions.save) return Promise.reject(new Error("editorNotReady"));
      return actions.save(options);
    },
    onClose: () => {
      void actions.close?.();
    },
    onBlurFocus: () => {
      if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    },
    onGrabFocus: () => {
      window.focus();
      actions.grabFocus?.();
    },
    onGetViews: () => actions.views,
  });
  wopiDebug({ bridgeEnabled: bridge.enabled, bridgeOrigin: config.postMessageOrigin });
  bridge.frameReady();
  wopiDebug({ frameReadySent: true });

  let root: ReturnType<typeof createRoot> | null = null;
  try {
    configureDefaultFonts({ load: () => import("@betteroffice/fonts"), baseUrl: "/fonts/" });
    let rootElement = document.getElementById("root");
    if (!rootElement) {
      rootElement = document.createElement("div");
      rootElement.id = "root";
      document.body.appendChild(rootElement);
    }
    root = createRoot(rootElement);
    const source = await api.source();
    wopiDebug({ step: "source", sourceBytes: source.byteLength });
    const join = await api.join();
    wopiDebug({ step: "join", joinFormat: join.format, initialUpdateBytes: join.initialUpdate.byteLength });
    const file = await prepareFile(source, join.format, { fontBaseUrl: "/fonts/" });
    wopiDebug({ step: "prepared", preparedFormat: file.format });
    const clientId = randomClientId();
    actions.views = [{ ViewId: clientId, UserName: displayUserName(config.userFriendlyName) }];
    const loaded: Loaded = { config, api, bridge, file, join, clientId, actions };
    root.render(<WopiApp loaded={loaded} />);
    wopiDebug({ step: "rendered" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Could not open the document.";
    wopiDebug({ step: "bootError", bootError: message });
    bridge.bootFailed();
    if (root) root.render(<div className="wopi-error" role="alert">{message}</div>);
    else document.body.textContent = message;
  }
}

void boot();
