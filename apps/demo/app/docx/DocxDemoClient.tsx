"use client";

import { useEffect, useMemo, useState } from "react";
import dynamic from "next/dynamic";
import Link from "next/link";
import type { BundledFontProvider } from "@betteroffice/docx-react";
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
  resolveScriptFallbackFace,
} from "@betteroffice/docx-fonts";
import { Logo } from "../components/Logo";

// The editor is browser-only (canvas + wasm + worker); keep it out of SSR.
const DocxEditor = dynamic(
  () => import("@betteroffice/docx-react").then((m) => m.DocxEditor),
  { ssr: false }
);

const SHOWCASE = { url: "/betteroffice-demo.docx", name: "betteroffice-demo.docx" };

export function DocxDemoClient() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  // Bundled metric-compatible faces (Carlito↔Calibri, Liberation↔Arial, …) so
  // the Rust measurement engine gets real bytes for documents that embed none.
  const measurementFontProvider = useMemo<BundledFontProvider>(
    () => ({
      resolve(family, bold, italic) {
        const face = resolveMetricCompatFace(family, bold, italic);
        return face ? () => loadBundledFontBytes(face) : undefined;
      },
      resolveScriptFallback(script, bold, italic) {
        const face = resolveScriptFallbackFace(script, bold, italic);
        return face ? () => loadBundledFontBytes(face) : undefined;
      },
      resolveLastResort(family, bold, italic) {
        return () => loadBundledFontBytes(resolveLastResortFace(family, bold, italic));
      },
    }),
    []
  );
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    fetch(SHOWCASE.url)
      .then((res) => {
        if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
        return res.arrayBuffer();
      })
      .then((bytes) => {
        if (!cancelled) setBuffer(bytes);
      })
      .catch((e: unknown) => {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <div className="docx-app">
      <header className="app-header">
        <div className="brand">
          <Link href="/" className="brand-mark">
            <Logo height={18} />
            BetterOffice <span className="brand-context">/ docx</span>
          </Link>
          <span className="brand-tagline">In-browser .docx editor</span>
        </div>

        <div className="spacer" />

        {buffer && <span className="filename">{SHOWCASE.name}</span>}

        <div className="actions">
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
      </header>
      <main className="stage" data-testid="docx-demo-stage">
        {error ? (
          <p className="stage-message" role="alert">
            Failed to load the demo document: {error}
          </p>
        ) : buffer ? (
          <DocxEditor
            documentBuffer={buffer}
            measurementFontProvider={measurementFontProvider}
            showToolbar
            showRuler
            showZoomControl
          />
        ) : (
          <p className="stage-message">Loading document…</p>
        )}
      </main>
    </div>
  );
}
