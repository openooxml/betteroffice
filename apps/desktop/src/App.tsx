import {
  lazy,
  Suspense,
  useCallback,
  useEffect,
  useRef,
  useState,
} from "react";
import { StartScreen } from "./StartScreen";
import {
  dispatchEditorShortcut,
  encodeBytes,
  formatOf,
  invoke,
  isNative,
  type Configuration,
  type OpenFile,
  type RecentFile,
} from "./bridge";
import { SessionRevision } from "./session";
import type { EditorApi, PreparedFile, Room } from "./Editor";

const Editor = lazy(() => import("./Editor"));
interface Session {
  file: OpenFile;
  prepared: PreparedFile;
  revision: SessionRevision;
}
const names = { docx: "Docs", xlsx: "Sheets", pptx: "Slides" };

export function App() {
  const [config, setConfig] = useState<Configuration>(() => {
    const format = formatOf(
      `file.${new URLSearchParams(location.search).get("format") ?? "docx"}`
    );
    return { format, name: `BetterOffice ${names[format]}`, recent: [] };
  });
  const [configured, setConfigured] = useState(!isNative());
  const [session, setSession] = useState<Session | null>(null);
  const sessionRef = useRef(session);
  sessionRef.current = session;
  const [dirty, setDirty] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const [room, setRoom] = useState<Room | null>(null);
  const [roomPanel, setRoomPanel] = useState(false);
  const [status, setStatus] = useState("");
  const api = useRef<EditorApi | null>(null);
  const sequence = useRef(0);
  const fileInput = useRef<HTMLInputElement>(null);
  const saving = useRef<Promise<boolean> | null>(null);
  const busyRef = useRef(false);

  const report = useCallback(
    (cause: unknown) =>
      setError(cause instanceof Error ? cause.message : String(cause)),
    []
  );
  const refreshConfig = useCallback(async () => {
    if (isNative())
      setConfig(await invoke<Configuration>({ type: "configuration" }));
    setConfigured(true);
  }, []);
  useEffect(() => {
    void refreshConfig().catch(report);
  }, [refreshConfig, report]);
  const synchronizeDirty = useCallback(
    (current: Session) => {
      if (sessionRef.current !== current) return;
      setDirty(current.revision.dirty);
      if (isNative())
        void invoke({
          type: "state",
          id: current.file.id,
          dirty: current.revision.dirty,
        }).catch(report);
    },
    [report]
  );
  const changed = useCallback(() => {
    const current = sessionRef.current;
    if (!current) return;
    current.revision.change();
    synchronizeDirty(current);
  }, [synchronizeDirty]);
  const ready = useCallback((value: EditorApi | null) => {
    api.current = value;
  }, []);

  const save = useCallback(
    (saveAs = false, supplied?: Uint8Array): Promise<boolean> => {
      if (saving.current) return saving.current;
      const current = sessionRef.current;
      if (!current || !api.current) return Promise.resolve(false);
      const suppliedCheckpoint = current.revision.checkpoint();
      const task = (async () => {
        try {
          if (document.activeElement instanceof HTMLElement)
            document.activeElement.blur();
          await new Promise<void>((resolve) => setTimeout(resolve, 0));
          const checkpoint = supplied
            ? suppliedCheckpoint
            : current.revision.checkpoint();
          const bytes = supplied ?? (await api.current!.serialize());
          if (isNative()) {
            const result = await invoke<{ name: string } | null>({
              type: "save",
              id: current.file.id,
              name: current.file.name,
              bytes: encodeBytes(bytes),
              saveAs,
            });
            if (!result) return false;
            current.file.name = result.name;

            await refreshConfig();
          } else {
            const url = URL.createObjectURL(new Blob([bytes.slice().buffer]));
            const link = document.createElement("a");
            link.href = url;
            link.download = current.file.name;
            link.click();
            setTimeout(() => URL.revokeObjectURL(url), 1000);
          }
          current.revision.didSave(checkpoint);
          synchronizeDirty(current);
          setError(null);
          return true;
        } catch (cause) {
          report(cause);
          return false;
        }
      })();
      saving.current = task;
      void task.finally(() => {
        saving.current = null;
      });
      return task;
    },
    [refreshConfig, report, synchronizeDirty]
  );

  const mayLeave = useCallback(async (): Promise<boolean> => {
    if (document.activeElement instanceof HTMLElement)
      document.activeElement.blur();
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    const current = sessionRef.current;
    if (!current || !current.revision.dirty) return true;
    const choice = isNative()
      ? await invoke<string>({ type: "confirmClose", name: current.file.name })
      : window.confirm("Discard unsaved changes?")
      ? "discard"
      : "cancel";
    return (
      choice === "discard" ||
      (choice === "save" && (await save()) && !current.revision.dirty)
    );
  }, [save]);

  const open = useCallback(
    async (file?: OpenFile) => {
      const request = ++sequence.current;
      try {
        if (!(await mayLeave()) || request !== sequence.current) return;
        if (!file) {
          if (!isNative()) {
            fileInput.current?.click();
            return;
          }
          file = (await invoke<OpenFile | null>({ type: "open" })) ?? undefined;
          if (!file || request !== sequence.current) return;
        }
        const format = formatOf(file.name);
        if (format !== config.format)
          throw new Error(
            `Open .${format} files in BetterOffice ${names[format]}.`
          );
        setBusy(true);
        busyRef.current = true;
        setError(null);
        const response = await fetch(file.url);
        if (!response.ok) throw new Error(`Could not read ${file.name}.`);
        const { prepareFile } = await import("./Editor");
        const prepared = await prepareFile(
          new Uint8Array(await response.arrayBuffer()),
          format
        );
        if (request !== sequence.current) return;
        const next = { file, prepared, revision: new SessionRevision() };
        if (isNative()) await invoke({ type: "activate", id: file.id });
        setRoom(null);
        setStatus("");
        setSession(next);
        sessionRef.current = next;
        setDirty(false);
        await refreshConfig();
      } catch (cause) {
        if (request === sequence.current) report(cause);
      } finally {
        if (request === sequence.current) {
          setBusy(false);
          busyRef.current = false;
        }
      }
    },
    [config.format, mayLeave, refreshConfig, report]
  );

  const openBrowserFile = useCallback(
    async (file: File) => {
      if (formatOf(file.name) !== config.format)
        throw new Error(`Choose a .${config.format} file.`);
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (isNative()) {
        await open(
          await invoke<OpenFile>({
            type: "import",
            name: file.name,
            bytes: encodeBytes(bytes),
          })
        );
      } else {
        const url = URL.createObjectURL(file);
        try {
          await open({ id: crypto.randomUUID(), name: file.name, url });
        } finally {
          URL.revokeObjectURL(url);
        }
      }
    },
    [config.format, open]
  );

  const close = useCallback(async () => {
    if (busyRef.current || !(await mayLeave())) return false;
    sequence.current += 1;
    setSession(null);
    sessionRef.current = null;
    setDirty(false);
    setRoom(null);
    setStatus("");
    if (isNative()) await invoke({ type: "deactivate" });
    return true;
  }, [mayLeave]);

  useEffect(() => {
    if (!configured) return;
    window.desktop = {
      open,
      close,
      command(command) {
        if (command === "open") void open();
        else if (command === "close") void close().catch(report);
        else if (command === "save" || command === "saveAs")
          void save(command === "saveAs");
        else if (command === "undo" || command === "redo") {
          api.current?.focus();
          dispatchEditorShortcut("z", command === "redo");
        } else if (command === "find") dispatchEditorShortcut("f");
        else if (command === "print") api.current?.print();
      },
    };
    if (isNative()) void invoke({ type: "ready" }).catch(report);
    const shortcut = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.repeat) return;
      const key = event.key.toLowerCase();
      if (!["o", "s", "w", "p"].includes(key)) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      window.desktop?.command(
        key === "o"
          ? "open"
          : key === "w"
          ? "close"
          : key === "p"
          ? "print"
          : event.shiftKey
          ? "saveAs"
          : "save"
      );
    };
    const beforeUnload = (event: BeforeUnloadEvent) => {
      if (sessionRef.current?.revision.dirty) event.preventDefault();
    };
    window.addEventListener("keydown", shortcut, true);
    window.addEventListener("beforeunload", beforeUnload);
    return () => {
      delete window.desktop;
      window.removeEventListener("keydown", shortcut, true);
      window.removeEventListener("beforeunload", beforeUnload);
    };
  }, [configured, open, close, save, report]);

  const switchRoom = async (nextRoom: Room | null) => {
    const current = sessionRef.current;
    if (!current || !api.current || busyRef.current) return;
    try {
      setBusy(true);
      busyRef.current = true;
      const bytes = await api.current.serialize();
      const { prepareFile } = await import("./Editor");
      const prepared = await prepareFile(bytes, current.prepared.format);
      if (sessionRef.current !== current) return;
      const next = { ...current, prepared };
      sessionRef.current = next;
      setSession(next);
      setRoom(nextRoom);
      setStatus("");
      setRoomPanel(false);
    } catch (cause) {
      report(cause);
    } finally {
      setBusy(false);
      busyRef.current = false;
    }
  };

  const openRecent = async (file: RecentFile) => {
    try {
      await open(await invoke<OpenFile>({ type: "openRecent", id: file.id }));
    } catch (cause) {
      report(cause);
    }
  };

  return (
    <div
      className="desktop"
      onDragOver={(event) => {
        event.preventDefault();
        setDragging(true);
      }}
      onDragLeave={(event) => {
        if (
          !(event.relatedTarget instanceof Node) ||
          !event.currentTarget.contains(event.relatedTarget)
        )
          setDragging(false);
      }}
      onDrop={(event) => {
        event.preventDefault();
        setDragging(false);
        const file = event.dataTransfer.files[0];
        if (file) void openBrowserFile(file).catch(report);
      }}
    >
      <header className="titlebar">
        <div className="brand">
          BetterOffice <span>/ {names[config.format]}</span>
        </div>
        {session && (
          <span className="document-name">
            {session.file.name}
            {dirty ? " · Edited" : ""}
          </span>
        )}
        {session && (
          <div
            className="title-actions"
            onMouseDown={(event) => event.preventDefault()}
          >
            <button
              className="quiet-button"
              onClick={() => setRoomPanel(!roomPanel)}
            >
              {room ? status || "Connecting…" : "Collaborate"}
            </button>
            <button
              className="quiet-button"
              onClick={() => void close().catch(report)}
              disabled={busy}
            >
              Home
            </button>
            <button
              className="quiet-button"
              onClick={() => void open()}
              disabled={busy}
            >
              Open…
            </button>
            <button
              className="quiet-button"
              onClick={() => void save()}
              disabled={busy}
            >
              Save
            </button>
          </div>
        )}
      </header>
      {error && (
        <div className="notice" role="alert">
          {error}
          <button aria-label="Dismiss error" onClick={() => setError(null)}>
            ×
          </button>
        </div>
      )}
      {roomPanel && session && (
        <form
          className="room-panel"
          onSubmit={(event) => {
            event.preventDefault();
            const form = new FormData(event.currentTarget);
            if (session.revision.dirty) {
              report(
                "Save your changes before joining a room. Everyone should open the same file."
              );
              return;
            }
            void switchRoom({
              id: String(form.get("room")).trim(),
              name: String(form.get("name")).trim(),
            });
          }}
        >
          <input
            name="room"
            aria-label="Room ID"
            placeholder="Room ID"
            defaultValue={room?.id}
            required
            maxLength={128}
          />
          <input
            name="name"
            aria-label="Your name"
            placeholder="Your name"
            defaultValue={room?.name}
            required
            maxLength={80}
          />
          <button className="quiet-button" type="submit">
            Join room
          </button>
          {room && (
            <button
              className="quiet-button"
              type="button"
              onClick={() => void switchRoom(null)}
            >
              Leave room
            </button>
          )}
          <p>Open the same file and enter the same room ID to edit together.</p>
        </form>
      )}
      {session ? (
        <main className="editor-stage" inert={busy}>
          <Suspense fallback={<p className="loading">Opening editor…</p>}>
            <Editor
              key={`${session.file.id}:${room?.id ?? ""}`}
              file={session.prepared}
              name={session.file.name}
              room={room}
              onSave={(bytes) => void save(false, bytes)}
              onChange={changed}
              onError={report}
              onReady={ready}
              onOpen={openBrowserFile}
              onStatus={setStatus}
            />
          </Suspense>
        </main>
      ) : (
        <StartScreen
          config={config}
          open={() => void open()}
          openRecent={(file) => void openRecent(file)}
          dragging={dragging}
          busy={busy}
        />
      )}
      <input
        ref={fileInput}
        type="file"
        accept={`.${config.format}`}
        hidden
        onChange={(event) => {
          const file = event.currentTarget.files?.[0];
          event.currentTarget.value = "";
          if (file) void openBrowserFile(file).catch(report);
        }}
      />
    </div>
  );
}
