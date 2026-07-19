"use client";

import dynamic from "next/dynamic";
import Link from "next/link";
import { useEffect, useState } from "react";
import type { PptxFontFace } from "@betteroffice/pptx";
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
} from "@betteroffice/docx-fonts";
import { Logo } from "../components/Logo";

const PptxEditor = dynamic(
  () => import("@betteroffice/pptx-react").then((module) => module.PptxEditor),
  { ssr: false },
);

const SHOWCASE = {
  url: "/betteroffice-demo.pptx",
  name: "betteroffice-demo.pptx",
};

type DemoAssets = {
  file: Uint8Array;
  fonts: PptxFontFace[];
};

export function PptxDemoClient() {
  const [assets, setAssets] = useState<DemoAssets | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    Promise.all([loadPresentation(), loadPresentationFonts()]).then(
      ([file, fonts]) => {
        if (!cancelled) setAssets({ file, fonts });
      },
      (value: unknown) => {
        if (!cancelled) {
          setError(value instanceof Error ? value.message : String(value));
        }
      },
    );
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <div className="demo-shell pptx-app">
      <header className="app-header">
        <div className="brand">
          <Link href="/" className="brand-mark">
            <Logo height={18} />
            BetterOffice <span className="brand-context">/ pptx</span>
          </Link>
          <span className="brand-tagline">In-browser presentation editor</span>
        </div>

        <div className="spacer" />

        {assets && <span className="filename">{SHOWCASE.name}</span>}

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
      <main className="stage" data-testid="pptx-demo-stage">
        {error ? (
          <p className="stage-message" role="alert">
            Failed to load the demo presentation: {error}
          </p>
        ) : assets ? (
          <PptxEditor file={assets.file} fonts={assets.fonts} />
        ) : (
          <p className="stage-message">Loading presentation…</p>
        )}
      </main>
    </div>
  );
}

async function loadPresentation(): Promise<Uint8Array> {
  const response = await fetch(SHOWCASE.url);
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
  return new Uint8Array(await response.arrayBuffer());
}

async function loadPresentationFonts(): Promise<PptxFontFace[]> {
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
