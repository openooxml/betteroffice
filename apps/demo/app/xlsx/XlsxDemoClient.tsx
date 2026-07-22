"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import { useSearchParams } from "next/navigation";
import { XlsxEditor } from "@betteroffice/xlsx-react";
import type { XlsxEditorApi } from "@betteroffice/xlsx-react";
import { isProposalsAvailable } from "@betteroffice/xlsx";
import { CollaborationProvider } from "@betteroffice/xlsx/collaboration";
import { Logo } from "../components/Logo";
import {
  CollaborationControls,
  COLLAB_RELAY_ORIGIN,
  useCollabRoom,
  useDemoRoom,
  type CollaborationReplica,
  type CollaborationTransport,
} from "../collab";
import { cn } from "../../lib/cn";
import { buildTotalsEdits } from "./demoAgent";

const SHOWCASE = { url: "/showcase.xlsx", name: "showcase.xlsx" };
const SAMPLE = { url: "/sample.xlsx", name: "sample.xlsx" };

const btn =
  "inline-flex cursor-pointer items-center gap-1.5 rounded-[5px] border border-hairline-strong bg-white px-[11px] py-[7px] font-mono text-[11px] leading-none font-normal whitespace-nowrap text-fg transition-colors duration-[140ms] ease-[ease] hover:bg-surface focus-visible:outline-1 focus-visible:outline-offset-2 focus-visible:outline-fg disabled:cursor-default disabled:opacity-50 disabled:hover:bg-white";
const btnPrimary = cn(
  btn,
  "border-fg bg-fg text-white hover:border-[#333] hover:bg-[#333]",
);
const btnGhost = cn(
  btn,
  "border-transparent bg-transparent text-mute hover:bg-surface hover:text-fg",
);

// a styled label wrapping a hidden file input, so "Open file" reads as a button.
function OpenFileLabel({
  className,
  testId,
  onPick,
}: {
  className: string;
  testId?: string;
  onPick: (files: FileList | null) => void;
}) {
  return (
    <label className={className}>
      <input
        data-testid={testId}
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        onChange={(e) => onPick(e.target.files)}
        className="hidden"
      />
      Open file
    </label>
  );
}

export function XlsxDemoClient() {
  const bootEmpty = useSearchParams().get("empty") === "1";
  const room = useDemoRoom();
  const createProvider = useCallback(
    (replica: CollaborationReplica, transport: CollaborationTransport) =>
      new CollaborationProvider(replica, transport),
    [],
  );
  const collab = useCollabRoom(COLLAB_RELAY_ORIGIN, room, createProvider);
  const [file, setFile] = useState<Uint8Array | undefined>();
  const [seed, setSeed] = useState<Uint8Array | null>(null);
  const [fileName, setFileName] = useState(SHOWCASE.name);
  const [collaborativeFile, setCollaborativeFile] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [ready, setReady] = useState(false);
  const [loading, setLoading] = useState(!bootEmpty);
  const [error, setError] = useState<string | null>(null);
  const apiRef = useRef<XlsxEditorApi | null>(null);
  // set once the user opens their own document, so the async showcase auto-load
  // can't clobber that choice if it resolves afterwards.
  const userActedRef = useRef(false);
  const autoloadedRef = useRef(false);

  const proposalsAvailable = isProposalsAvailable();

  const onReady = useCallback((api: XlsxEditorApi) => {
    apiRef.current = api;
    setReady(true);
  }, []);

  // the demo agent: stage =SUM totals for the numeric columns as a proposal.
  const proposeTotals = useCallback(() => {
    const api = apiRef.current;
    if (!api) return;
    const edits = buildTotalsEdits(api.handle);
    if (edits.length === 0) return;
    try {
      api.handle.propose("demo-agent", "column totals", edits);
      api.refreshProposals();
    } catch {
      // proposals not in this wasm build — nothing staged.
    }
  }, []);

  // read a blob's bytes into the editor under the given display name.
  const openBlob = useCallback(
    async (blob: Blob, name: string, collaborative = false) => {
      setLoading(true);
      try {
        const bytes = new Uint8Array(await blob.arrayBuffer());
        setError(null);
        setCollaborativeFile(collaborative);
        setFile(bytes);
        setFileName(name);
      } catch (e) {
        setError(
          `Could not read ${name}: ${e instanceof Error ? e.message : String(e)}`,
        );
      } finally {
        setLoading(false);
      }
    },
    [],
  );

  // fetch a bundled workbook by url; surface fetch failures as a page banner.
  const openUrl = useCallback(
    async (url: string, name: string, collaborative = false) => {
      setLoading(true);
      try {
        const res = await fetch(url);
        if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
        await openBlob(await res.blob(), name, collaborative);
      } catch (e) {
        setError(
          `Could not load ${name}: ${e instanceof Error ? e.message : String(e)}`,
        );
        setLoading(false);
      }
    },
    [openBlob],
  );

  const onPick = useCallback(
    (list: FileList | null) => {
      const picked = list?.[0];
      if (!picked) return;
      userActedRef.current = true;
      void openBlob(picked, picked.name, false);
    },
    [openBlob],
  );

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const picked = e.dataTransfer.files?.[0];
      if (!picked) return;
      userActedRef.current = true;
      void openBlob(picked, picked.name, false);
    },
    [openBlob],
  );

  const loadSample = useCallback(() => {
    userActedRef.current = true;
    void openUrl(SAMPLE.url, SAMPLE.name, false);
  }, [openUrl]);

  const loadShowcase = useCallback(() => {
    userActedRef.current = true;
    void openUrl(SHOWCASE.url, SHOWCASE.name, true);
  }, [openUrl]);

  const closeFile = useCallback(() => {
    userActedRef.current = true;
    setFile(undefined);
    setFileName("");
    setCollaborativeFile(false);
    setError(null);
  }, []);

  // auto-load the showcase on first visit, unless the user already opened
  // something. runs once (the ref guards react's strict-mode double effect).
  useEffect(() => {
    if (autoloadedRef.current) return;
    autoloadedRef.current = true;
    if (bootEmpty || userActedRef.current) {
      setLoading(false);
      return;
    }
    void openUrl(SHOWCASE.url, SHOWCASE.name, true);
  }, [bootEmpty, openUrl]);

  useEffect(() => {
    let cancelled = false;
    fetch("/seeds/xlsx.bin")
      .then((response) => {
        if (!response.ok) {
          throw new Error(`${response.status} ${response.statusText}`);
        }
        return response.arrayBuffer();
      })
      .then((bytes) => {
        if (!cancelled) setSeed(new Uint8Array(bytes));
      })
      .catch((nextError: unknown) => {
        if (cancelled) return;
        setError(
          `Could not load collaboration seed: ${
            nextError instanceof Error ? nextError.message : String(nextError)
          }`,
        );
      });
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <div className="fixed inset-0 z-20 flex flex-col bg-surface text-fg">
      <header className="z-2 flex items-center gap-3.5 border-b border-hairline bg-white/92 px-4 py-[11px] backdrop-blur-lg max-[720px]:flex-wrap max-[720px]:gap-2 max-[720px]:px-2.5 max-[720px]:py-2">
        <div className="flex min-w-0 items-baseline gap-2.5 max-[720px]:flex-1">
          <Link
            href="/"
            className="inline-flex items-baseline gap-2 text-[14px] font-[650] tracking-[-0.01em] whitespace-nowrap text-fg no-underline"
          >
            <Logo height={18} className="self-center" />
            BetterOffice <span className="font-normal text-faint">/ xlsx</span>
          </Link>
          <span className="overflow-hidden text-[12.5px] text-ellipsis whitespace-nowrap text-mute max-[720px]:hidden">
            In-browser spreadsheet editor
          </span>
        </div>

        <div className="flex-1 max-[720px]:hidden" />

        {file && fileName && (
          <span className="max-w-[180px] overflow-hidden text-[12.5px] text-ellipsis whitespace-nowrap text-mute max-[720px]:hidden">
            {fileName}
          </span>
        )}

        <div className="flex flex-none items-center gap-2 max-[720px]:order-3 max-[720px]:w-full max-[720px]:overflow-x-auto max-[720px]:pb-0.5">
          {collaborativeFile && (
            <CollaborationControls
              status={collab.status}
              synced={collab.synced}
              peerCount={collab.peerCount}
              error={collab.error}
            />
          )}
          <div className="flex items-center gap-2 max-[720px]:flex-none">
            <OpenFileLabel className={btn} testId="file-input" onPick={onPick} />
            <button
              className={btn}
              data-testid="load-sample"
              onClick={loadSample}
            >
              Load sample
            </button>
            <button className={btn} onClick={loadShowcase}>
              Load showcase
            </button>
            {file && (
              <button className={btnGhost} onClick={closeFile} title="Close file">
                Close
              </button>
            )}
            <a
              className="inline-flex size-8 items-center justify-center rounded-[5px] text-mute transition-colors duration-[140ms] ease-[ease] hover:bg-surface hover:text-fg max-[720px]:hidden"
              href="https://github.com/openooxml/betteroffice"
              target="_blank"
              rel="noreferrer"
              aria-label="View on GitHub"
              title="View on GitHub"
            >
              <svg
                width="18"
                height="18"
                viewBox="0 0 16 16"
                fill="currentColor"
                aria-hidden="true"
              >
                <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z" />
              </svg>
            </a>
          </div>

          {proposalsAvailable && (
            <div className="ml-0.5 flex items-center gap-2 border-l border-hairline pl-2.5 max-[720px]:flex-none">
              <button
                className={btnPrimary}
                data-testid="propose-totals"
                onClick={proposeTotals}
                disabled={!ready}
                title="Stage =SUM totals for numeric columns as an agent proposal"
              >
                Agent: propose totals
              </button>
            </div>
          )}
        </div>
      </header>

      <div
        className="relative flex min-h-0 flex-1 flex-col"
        data-testid="editor-drop"
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
      >
        {error && (
          <div
            className="flex items-center gap-3 border-b border-[#f3c7cf] bg-[#fdecef] px-4 py-2 text-[13px] text-danger"
            role="alert"
          >
            <span>{error}</span>
            <span className="flex-1" />
            <button
              className="cursor-pointer rounded bg-transparent px-1.5 py-0.5 text-[16px] leading-none text-danger hover:bg-danger/10"
              onClick={() => setError(null)}
              aria-label="Dismiss error"
            >
              ×
            </button>
          </div>
        )}

        {/* the editor stays mounted even with no file (it paints a demo frame),
            so the empty state is an overlay on top rather than an unmount. */}
        <div
          className={cn(
            "relative flex min-h-0 min-w-0 flex-1",
            dragging &&
              file &&
              "outline-2 outline-dashed -outline-offset-3 outline-acc",
          )}
        >
          <XlsxEditor
            file={file}
            fileName={fileName}
            collaboration={
              collaborativeFile && room && seed && collab.clientId
                ? {
                    clientId: collab.clientId,
                    initialUpdate: seed,
                    onReplica: collab.onReplica,
                  }
                : undefined
            }
            onReady={onReady}
          />

          {!file &&
            (loading ? (
              <div className="absolute inset-0 z-1 grid place-items-center bg-surface text-[13.5px] text-mute">
                Opening workbook…
              </div>
            ) : (
              <div className="absolute inset-0 z-1 grid place-items-center bg-surface p-8 max-[720px]:p-4">
                <div
                  className={cn(
                    "w-[min(460px,100%)] rounded-md border border-hairline bg-white px-9 py-[42px] text-center transition-colors duration-[140ms] ease-[ease] max-[720px]:px-5 max-[720px]:py-8",
                    dragging && "border-acc bg-[#f4fbf8]",
                  )}
                >
                  <p className="mb-2 text-[20px] font-[650] tracking-[-0.02em]">
                    Drop an .xlsx here
                  </p>
                  <p className="mb-[22px] text-[13.5px] leading-normal text-mute">
                    Open a file from your computer — nothing is uploaded,
                    everything runs locally in your browser.
                  </p>
                  <div className="flex flex-wrap justify-center gap-2.5">
                    <OpenFileLabel className={btnPrimary} onPick={onPick} />
                    <button className={btn} onClick={loadSample}>
                      Load sample
                    </button>
                    <button className={btn} onClick={loadShowcase}>
                      Load showcase
                    </button>
                  </div>
                </div>
              </div>
            ))}
        </div>
      </div>
    </div>
  );
}
