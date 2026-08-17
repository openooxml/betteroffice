"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import dynamic from "next/dynamic";
import Link from "next/link";
import { CollaborationProvider } from "@betteroffice/docx/collaboration";
import { configureDefaultFonts } from "@betteroffice/docx/layout";
import * as fonts from "@betteroffice/fonts";
import { Logo } from "../components/Logo";
import {
  CollaborationControls,
  COLLAB_RELAY_ORIGIN,
  useCollabRoom,
  useDemoIdentity,
  useDemoRoom,
  useLeaveRoom,
  type CollaborationReplica,
  type CollaborationTransport,
} from "../collab";
import { planDemoSession } from "../../lib/demoSession";

// The engine loads no fonts on its own; hand it the bundled set before any editor mounts.
configureDefaultFonts({ fonts });

// The editor is browser-only (canvas + wasm + worker); keep it out of SSR.
const DocxEditor = dynamic(
  () => import("@betteroffice/docx-react").then((m) => m.DocxEditor),
  { ssr: false }
);

const SHOWCASE = { url: "/betteroffice-demo.docx", name: "betteroffice-demo.docx" };

/** `id` keys the editor, so each loaded document gets its own session. */
interface DemoSource {
  id: number;
  buffer: ArrayBuffer;
  name: string;
  seed: Uint8Array | null;
}

export function DocxDemoClient() {
  const [source, setSource] = useState<DemoSource | null>(null);
  const openSequence = useRef(0);
  const room = useDemoRoom(source ? source.seed !== null : true);
  const leaveRoom = useLeaveRoom();
  const user = useDemoIdentity();
  const createProvider = useCallback(
    (replica: CollaborationReplica, transport: CollaborationTransport) =>
      new CollaborationProvider(replica, transport, {
        user: user ?? undefined,
      }),
    [user],
  );
  const collab = useCollabRoom(
    COLLAB_RELAY_ORIGIN,
    room,
    createProvider,
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
        if (cancelled || openSequence.current !== 0) return;
        setSource({
          id: 0,
          buffer: documentBytes,
          name: SHOWCASE.name,
          seed: new Uint8Array(seedBytes),
        });
      })
      .catch((e: unknown) => {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const session = useMemo(
    () =>
      planDemoSession({
        document: source,
        room,
        clientId: collab.clientId,
        identified: user !== null,
      }),
    [collab.clientId, room, source, user],
  );

  const collaboration = useMemo(
    () =>
      session.status === "shared" && source?.seed && user
        ? {
            clientId: session.clientId,
            initialUpdate: source.seed,
            user,
            onReplica: collab.onReplica,
            presence: collab.provider ?? undefined,
          }
        : undefined,
    [collab.onReplica, collab.provider, session, source, user],
  );

  // Ordered by selection, not by completion: a slower earlier pick must not
  // land on top of a later one.
  const handleOpen = useCallback(
    async (file: File) => {
      const sequence = (openSequence.current += 1);
      try {
        const buffer = await file.arrayBuffer();
        if (sequence !== openSequence.current) return;
        leaveRoom();
        setError(null);
        setSource({ id: sequence, buffer, name: file.name, seed: null });
      } catch (cause) {
        if (sequence !== openSequence.current) return;
        setError(cause instanceof Error ? cause.message : String(cause));
      }
    },
    [leaveRoom],
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

        {source && (
          <span className="max-w-[180px] overflow-hidden text-[12.5px] text-ellipsis whitespace-nowrap text-mute">
            {source.name}
          </span>
        )}

        <div className="flex flex-none items-center gap-2">
          <CollaborationControls
            status={collab.status}
            synced={collab.synced}
            peerCount={collab.peerCount}
            error={collab.error}
            shared={session.status === "shared"}
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
        ) : source && session.status !== "loading" ? (
          <DocxEditor
            key={source.id}
            documentBuffer={source.buffer}
            collaboration={collaboration}
            documentName={source.name}
            onOpen={handleOpen}
            onError={(cause) => setError(cause.message)}
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
