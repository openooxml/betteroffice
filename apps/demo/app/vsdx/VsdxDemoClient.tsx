"use client";

import dynamic from "next/dynamic";
import Link from "next/link";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { CollaborationProvider, initWasm, openDiagram, type CollaborationUser, type VsdxFontFace } from "@betteroffice/vsdx";
import type { VsdxEditorApi } from "@betteroffice/vsdx-react";
import { loadBundledFontBytes, resolveLastResortFace, resolveMetricCompatFace } from "@betteroffice/fonts";
import { Logo } from "../components/Logo";
import { CollaborationControls, COLLAB_RELAY_ORIGIN, useCollabRoom, useDemoRoom, useLeaveRoom, type CollaborationReplica, type CollaborationTransport } from "../collab";
import { planDemoSession } from "../../lib/demoSession";
import { readLocalDiagram } from "../../lib/localDiagram";
import { clearVsdxApi, exposeVsdxApi, shouldExposeVsdxApi } from "./vsdxApiExposure";

const VsdxEditor = dynamic(
  () => import("@betteroffice/vsdx-react").then((module) => module.VsdxEditor),
  { ssr: false },
);

const SHOWCASE = {
  url: "/betteroffice-demo.vsdx",
  name: "betteroffice-demo.vsdx",
};

/** Loaded bytes plus the room seed; a local file carries no seed and stays private. */
interface DemoSource {
  id: number;
  file: Uint8Array;
  name: string;
  seed: Uint8Array | null;
}

export function VsdxDemoClient() {
  const [source, setSource] = useState<DemoSource | null>(null);
  const [fonts, setFonts] = useState<VsdxFontFace[] | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [openError, setOpenError] = useState<string | null>(null);
  const [opening, setOpening] = useState(false);
  const [dragging, setDragging] = useState(false);
  const openSequence = useRef(0);
  const dragDepth = useRef(0);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const room = useDemoRoom(source ? source.seed !== null : true);
  const leaveRoom = useLeaveRoom();
  const createProvider = useCallback((replica: CollaborationReplica, transport: CollaborationTransport) => new CollaborationProvider(replica, transport, { user: { name: presenceName() } }), []);
  const collab = useCollabRoom(COLLAB_RELAY_ORIGIN, room, createProvider);
  const readyApiRef = useRef<VsdxEditorApi | null>(null);
  const handleReady = useCallback((api: VsdxEditorApi) => {
    readyApiRef.current = api;
    if (shouldExposeVsdxApi(window.location.search)) exposeVsdxApi(window, api);
  }, []);
  useEffect(() => () => {
    const api = readyApiRef.current;
    readyApiRef.current = null;
    if (api) clearVsdxApi(window, api);
  }, [room, source?.id]);

  useEffect(() => {
    let cancelled = false;
    void Promise.all([loadDiagram(), loadCollaborationSeed(), loadDiagramFonts()]).then(
      ([file, seed, loadedFonts]) => {
        if (cancelled) return;
        setFonts(loadedFonts);
        if (openSequence.current !== 0) return;
        setSource({ id: 0, file, name: SHOWCASE.name, seed });
      },
      (value: unknown) => { if (!cancelled) setLoadError(value instanceof Error ? value.message : String(value)); },
    );
    return () => {
      cancelled = true;
    };
  }, []);

  // Ordered by selection, not by completion: a slower earlier pick must not land on a later one.
  const openChosenFile = useCallback(async (picked: File) => {
    const sequence = (openSequence.current += 1);
    setOpening(true);
    const result = await readLocalDiagram(picked, ensureVisioOpenable);
    if (openSequence.current !== sequence) return;
    setOpening(false);
    if ("refused" in result) {
      setOpenError(result.refused);
      return;
    }
    leaveRoom();
    setOpenError(null);
    setSource({ id: sequence, file: result.opened.bytes, name: result.opened.name, seed: null });
  }, [leaveRoom]);

  const onDrop = useCallback((event: React.DragEvent) => {
    event.preventDefault();
    dragDepth.current = 0;
    setDragging(false);
    const picked = event.dataTransfer.files[0];
    if (picked) void openChosenFile(picked);
  }, [openChosenFile]);

  const session = useMemo(
    () => planDemoSession({ document: source, room, clientId: collab.clientId, identified: true }),
    [collab.clientId, room, source],
  );

  const collaboration = useMemo(
    () => session.status === "shared" && source?.seed
      ? { clientId: session.clientId, initialUpdate: source.seed, onReplica: collab.onReplica, presence: collab.provider ?? undefined }
      : undefined,
    [collab.onReplica, collab.provider, session, source],
  );

  return (
    <div className="fixed inset-0 z-20 flex flex-col bg-surface text-fg">
      <header className="z-2 flex items-center gap-3.5 border-b border-hairline bg-white/92 px-4 py-[11px] backdrop-blur-lg">
        <div className="flex min-w-0 items-baseline gap-2.5">
          <Link href="/" className="inline-flex items-baseline gap-2 text-[14px] font-[650] tracking-[-0.01em] whitespace-nowrap text-fg no-underline">
            <Logo height={18} className="self-center" />
            BetterOffice <span className="font-normal text-faint">/ vsdx</span>
          </Link>
          <span className="overflow-hidden text-[12.5px] text-ellipsis whitespace-nowrap text-mute">In-browser Visio diagram editor</span>
        </div>
        <div className="flex-1" />
        {source && <span className="max-w-[180px] overflow-hidden text-[12.5px] text-ellipsis whitespace-nowrap text-mute">{source.name}</span>}
        <div className="flex flex-none items-center gap-2">
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={opening || !fonts}
            aria-label="Open a Visio file from your computer"
            className="inline-flex h-8 cursor-pointer items-center rounded-[5px] border border-hairline-strong bg-white px-[11px] text-[12.5px] text-fg transition-colors duration-[140ms] ease-[ease] hover:bg-surface disabled:cursor-default disabled:opacity-50"
          >
            {opening ? "Opening…" : "Open file"}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".vsdx,.vstx"
            aria-label="Choose a Visio file"
            data-testid="vsdx-file-input"
            style={{ display: "none" }}
            onChange={(event) => {
              const picked = event.target.files?.[0];
              event.target.value = "";
              if (picked) void openChosenFile(picked);
            }}
          />
          <CollaborationControls status={collab.status} synced={collab.synced} peerCount={collab.peerCount} error={collab.error} shared={session.status === "shared"} />
          <a className="inline-flex size-8 items-center justify-center rounded-[5px] text-mute transition-colors duration-[140ms] ease-[ease] hover:bg-surface hover:text-fg" href="https://github.com/openooxml/betteroffice" target="_blank" rel="noreferrer" aria-label="View on GitHub" title="View on GitHub">
          <svg width="18" height="18" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z" /></svg>
          </a>
        </div>
      </header>
      {openError && (
        <div className="z-2 flex items-center gap-3 border-b border-[#f3c7cf] bg-[#fdecef] px-4 py-2 text-[13px] text-danger" role="alert">
          <span>{openError}</span>
          <span className="flex-1" />
          <button type="button" onClick={() => setOpenError(null)} aria-label="Dismiss error" className="cursor-pointer rounded bg-transparent px-1.5 py-0.5 text-[16px] leading-none text-danger hover:bg-danger/10">
            ×
          </button>
        </div>
      )}
      <main
        className="relative flex min-h-0 flex-1 flex-col *:min-h-0 *:flex-1"
        data-testid="vsdx-demo-stage"
        onDragEnter={(event) => {
          if (!dragsFiles(event)) return;
          event.preventDefault();
          dragDepth.current += 1;
          setDragging(true);
        }}
        onDragOver={(event) => {
          if (dragsFiles(event)) event.preventDefault();
        }}
        onDragLeave={(event) => {
          if (!dragsFiles(event)) return;
          dragDepth.current = Math.max(0, dragDepth.current - 1);
          if (dragDepth.current === 0) setDragging(false);
        }}
        onDrop={onDrop}
      >
        {loadError ? <p className="m-auto text-mute" role="alert">Failed to load the demo diagram: {loadError}</p> : source && fonts && session.status !== "loading" ? <VsdxEditor key={`${room ?? "private"}:${source.id}`} file={source.file} fonts={fonts} collaboration={collaboration} onReady={handleReady} /> : <p className="m-auto text-mute">Loading diagram…</p>}
        {dragging && (
          <div className="pointer-events-none absolute inset-0 z-10 grid place-items-center bg-white/70 p-8" role="status">
            <div className="grid w-[min(440px,100%)] place-items-center rounded-md border-2 border-dashed border-acc bg-white px-8 py-10 text-center">
              <p className="mb-1 text-[16px] font-[650]">Drop to open the diagram</p>
              <p className="text-[13px] text-mute">Only .vsdx and .vstx files, opened locally in your browser.</p>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

/** Opens the bytes once with the engine, so a file it refuses never replaces the loaded diagram. */
async function ensureVisioOpenable(bytes: Uint8Array): Promise<void> {
  await initWasm();
  openDiagram(bytes).dispose();
}

function dragsFiles(event: React.DragEvent): boolean {
  return event.dataTransfer.types.includes("Files");
}

function presenceName(): CollaborationUser["name"] {
  return `VSDX ${crypto.getRandomValues(new Uint32Array(1))[0].toString(36)}`;
}

async function loadCollaborationSeed(): Promise<Uint8Array> {
  const response = await fetch("/seeds/vsdx.bin");
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
  return new Uint8Array(await response.arrayBuffer());
}

async function loadDiagram(): Promise<Uint8Array> {
  const response = await fetch(SHOWCASE.url);
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
  return new Uint8Array(await response.arrayBuffer());
}

async function loadDiagramFonts(): Promise<VsdxFontFace[]> {
  const styles = [
    { bold: false, italic: false },
    { bold: true, italic: false },
    { bold: false, italic: true },
    { bold: true, italic: true },
  ];
  return Promise.all(
    styles.map(async ({ bold, italic }) => {
      const face =
        resolveMetricCompatFace("Arial", bold, italic) ??
        resolveLastResortFace("Arial", bold, italic);
      const bytes = await loadBundledFontBytes(face);
      return {
        family: "Arial",
        bold,
        italic,
        bytes: new Uint8Array(bytes.slice(0)),
      };
    }),
  );
}
