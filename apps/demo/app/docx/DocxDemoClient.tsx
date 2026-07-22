"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import dynamic from "next/dynamic";
import Link from "next/link";
import type { BundledFontProvider } from "@betteroffice/docx-react";
import { CollaborationProvider } from "@betteroffice/docx/collaboration";
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
  resolveScriptFallbackFace,
} from "@betteroffice/docx-fonts";
import { Logo } from "../components/Logo";
import {
  CollaborationControls,
  COLLAB_RELAY_ORIGIN,
  useCollabRoom,
  useDemoRoom,
  type CollaborationReplica,
  type CollaborationTransport,
} from "../collab";

// The editor is browser-only (canvas + wasm + worker); keep it out of SSR.
const DocxEditor = dynamic(
  () => import("@betteroffice/docx-react").then((m) => m.DocxEditor),
  { ssr: false }
);

const SHOWCASE = { url: "/betteroffice-demo.docx", name: "betteroffice-demo.docx" };

export function DocxDemoClient() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const [seed, setSeed] = useState<Uint8Array | null>(null);
  const room = useDemoRoom();
  const createProvider = useCallback(
    (replica: CollaborationReplica, transport: CollaborationTransport) =>
      new CollaborationProvider(replica, transport),
    [],
  );
  const collab = useCollabRoom(
    COLLAB_RELAY_ORIGIN,
    room,
    createProvider,
  );
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
    Promise.all([fetch(SHOWCASE.url), fetch("/seeds/docx.bin")])
      .then(async ([documentResponse, seedResponse]) => {
        if (!documentResponse.ok) {
          throw new Error(
            `${documentResponse.status} ${documentResponse.statusText}`,
          );
        }
        if (!seedResponse.ok) {
          throw new Error(`${seedResponse.status} ${seedResponse.statusText}`);
        }
        return Promise.all([
          documentResponse.arrayBuffer(),
          seedResponse.arrayBuffer(),
        ]);
      })
      .then(([documentBytes, seedBytes]) => {
        if (cancelled) return;
        setBuffer(documentBytes);
        setSeed(new Uint8Array(seedBytes));
      })
      .catch((e: unknown) => {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const collaboration = useMemo(
    () =>
      room && seed && collab.clientId
        ? {
            clientId: collab.clientId,
            initialUpdate: seed,
            onReplica: collab.onReplica,
          }
        : undefined,
    [collab.clientId, collab.onReplica, room, seed],
  );

  return (
    <div className="fixed inset-0 z-20 flex flex-col bg-surface text-fg">
      <header className="z-2 flex items-center gap-3.5 border-b border-hairline bg-white/92 px-4 py-[11px] backdrop-blur-lg">
        <div className="flex min-w-0 items-baseline gap-2.5">
          <Link
            href="/"
            className="inline-flex items-baseline gap-2 text-[14px] font-[650] tracking-[-0.01em] whitespace-nowrap text-fg no-underline"
          >
            <Logo height={18} className="self-center" />
            BetterOffice <span className="font-normal text-faint">/ docx</span>
          </Link>
          <span className="overflow-hidden text-[12.5px] text-ellipsis whitespace-nowrap text-mute">
            In-browser .docx editor
          </span>
        </div>

        <div className="flex-1" />

        {buffer && (
          <span className="max-w-[180px] overflow-hidden text-[12.5px] text-ellipsis whitespace-nowrap text-mute">
            {SHOWCASE.name}
          </span>
        )}

        <div className="flex flex-none items-center gap-2">
          <CollaborationControls
            status={collab.status}
            synced={collab.synced}
            peerCount={collab.peerCount}
            error={collab.error}
          />
          <a
            className="inline-flex size-8 items-center justify-center rounded-[5px] text-mute transition-colors duration-[140ms] ease-[ease] hover:bg-surface hover:text-fg"
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
      <main
        className="flex min-h-0 flex-1 flex-col *:min-h-0 *:flex-1"
        data-testid="docx-demo-stage"
      >
        {error ? (
          <p className="m-auto text-mute" role="alert">
            Failed to load the demo document: {error}
          </p>
        ) : buffer ? (
          <DocxEditor
            documentBuffer={buffer}
            collaboration={collaboration}
            measurementFontProvider={measurementFontProvider}
            showToolbar
            showRuler
            showZoomControl
          />
        ) : (
          <p className="m-auto text-mute">Loading document…</p>
        )}
      </main>
    </div>
  );
}
