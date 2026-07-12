"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import { useSearchParams } from "next/navigation";
import { XlsxEditor } from "@betteroffice/xlsx-react";
import type { XlsxEditorApi } from "@betteroffice/xlsx-react";
import { isProposalsAvailable } from "@betteroffice/xlsx";
import { Logo } from "../components/Logo";
import { buildTotalsEdits } from "./demoAgent";

const SHOWCASE = { url: "/showcase.xlsx", name: "showcase.xlsx" };
const SAMPLE = { url: "/sample.xlsx", name: "sample.xlsx" };

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
        style={{ display: "none" }}
      />
      Open file
    </label>
  );
}

export function XlsxDemoClient() {
  const bootEmpty = useSearchParams().get("empty") === "1";
  const [file, setFile] = useState<Uint8Array | undefined>();
  const [fileName, setFileName] = useState(SHOWCASE.name);
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
  const openBlob = useCallback(async (blob: Blob, name: string) => {
    setLoading(true);
    try {
      const bytes = new Uint8Array(await blob.arrayBuffer());
      setError(null);
      setFile(bytes);
      setFileName(name);
    } catch (e) {
      setError(
        `Could not read ${name}: ${e instanceof Error ? e.message : String(e)}`,
      );
    } finally {
      setLoading(false);
    }
  }, []);

  // fetch a bundled workbook by url; surface fetch failures as a page banner.
  const openUrl = useCallback(
    async (url: string, name: string) => {
      setLoading(true);
      try {
        const res = await fetch(url);
        if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
        await openBlob(await res.blob(), name);
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
      void openBlob(picked, picked.name);
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
      void openBlob(picked, picked.name);
    },
    [openBlob],
  );

  const loadSample = useCallback(() => {
    userActedRef.current = true;
    void openUrl(SAMPLE.url, SAMPLE.name);
  }, [openUrl]);

  const loadShowcase = useCallback(() => {
    userActedRef.current = true;
    void openUrl(SHOWCASE.url, SHOWCASE.name);
  }, [openUrl]);

  const closeFile = useCallback(() => {
    userActedRef.current = true;
    setFile(undefined);
    setFileName("");
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
    void openUrl(SHOWCASE.url, SHOWCASE.name);
  }, [bootEmpty, openUrl]);

  return (
    <div className="app">
      <header className="app-header">
        <div className="brand">
          <Link href="/" className="brand-mark">
            <Logo height={18} />
            BetterOffice <span className="brand-context">/ xlsx</span>
          </Link>
          <span className="brand-tagline">In-browser spreadsheet editor</span>
        </div>

        <div className="spacer" />

        {file && fileName && <span className="filename">{fileName}</span>}

        <div className="actions">
          <div className="action-group">
            <OpenFileLabel
              className="btn"
              testId="file-input"
              onPick={onPick}
            />
            <button
              className="btn"
              data-testid="load-sample"
              onClick={loadSample}
            >
              Load sample
            </button>
            <button className="btn" onClick={loadShowcase}>
              Load showcase
            </button>
            {file && (
              <button
                className="btn btn-ghost"
                onClick={closeFile}
                title="Close file"
              >
                Close
              </button>
            )}
            <a
              className="github-link"
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
            <div className="action-group">
              <button
                className="btn btn-primary"
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
        className={`app-body${dragging && file ? " drag" : ""}`}
        data-testid="editor-drop"
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
      >
        {error && (
          <div className="error-banner" role="alert">
            <span>{error}</span>
            <span className="spacer" />
            <button
              className="error-dismiss"
              onClick={() => setError(null)}
              aria-label="Dismiss error"
            >
              ×
            </button>
          </div>
        )}

        {/* the editor stays mounted even with no file (it paints a demo frame),
            so the empty state is an overlay on top rather than an unmount. */}
        <div className="editor-host">
          <XlsxEditor file={file} fileName={fileName} onReady={onReady} />

          {!file &&
            (loading ? (
              <div className="overlay busy">Opening workbook…</div>
            ) : (
              <div className="overlay empty">
                <div className={`dropzone${dragging ? " drag" : ""}`}>
                  <p className="dropzone-title">Drop an .xlsx here</p>
                  <p className="dropzone-sub">
                    Open a file from your computer — nothing is uploaded,
                    everything runs locally in your browser.
                  </p>
                  <div className="dropzone-actions">
                    <OpenFileLabel
                      className="btn btn-primary"
                      onPick={onPick}
                    />
                    <button className="btn" onClick={loadSample}>
                      Load sample
                    </button>
                    <button className="btn" onClick={loadShowcase}>
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
