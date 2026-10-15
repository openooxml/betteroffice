"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import dynamic from "next/dynamic";
import { CollaborationProvider } from "@betteroffice/docx/collaboration";
import { configureDefaultFonts } from "@betteroffice/docx/layout";
import { setGoogleFontsEnabled } from "@betteroffice/docx/utils";
import type { DocxPlugin, DocxPluginGrant } from "@betteroffice/docx-react";
import { DemoTopBar, barText } from "../components/DemoTopBar";
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
import { isFatalDocxError } from "./docxError";
import { planDemoSession } from "../../lib/demoSession";

configureDefaultFonts({ load: () => import("@betteroffice/fonts/cdn") });
setGoogleFontsEnabled(false);

// The editor is browser-only (canvas + wasm + worker); keep it out of SSR.
const DocxEditor = dynamic(
  () => import("@betteroffice/docx-react").then((m) => m.DocxEditor),
  { ssr: false, loading: () => <DocumentLoading /> }
);
const CompactToolbar = dynamic(
  () => import("./CompactToolbar").then((m) => m.CompactToolbar),
  { ssr: false }
);

const SHOWCASE = { url: "/betteroffice-demo.docx", name: "betteroffice-demo.docx" };

function DocumentLoading() {
  return (
    <div
      className="oox-root flex flex-col items-center justify-center gap-5 bg-[var(--doc-bg)] text-[var(--doc-text-muted)]"
      role="status"
    >
      <div className="size-9 animate-[spin_0.8s_linear_infinite] rounded-full border-3 border-[var(--doc-border)] border-t-[var(--doc-primary)]" />
      <div className="text-[14px]">Loading document...</div>
    </div>
  );
}

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
  const [compact, setCompact] = useState(false);
  const [reviewPlugin, setReviewPlugin] = useState<DocxPlugin | null>(null);
  const [reviewEnabled, setReviewEnabled] = useState(false);
  const [reviewWrites, setReviewWrites] = useState(false);
  const plugins = useMemo(
    () => (reviewEnabled && reviewPlugin ? [reviewPlugin] : []),
    [reviewEnabled, reviewPlugin],
  );
  const pluginGrants = useMemo<Record<string, DocxPluginGrant> | undefined>(
    () =>
      reviewWrites && reviewPlugin
        ? { [reviewPlugin.id]: { document: "write", editBatches: true } }
        : undefined,
    [reviewPlugin, reviewWrites],
  );

  useEffect(() => {
    let cancelled = false;
    void import("./ReviewPlugin").then((module) => {
      if (!cancelled) setReviewPlugin(module.reviewPlugin);
    });
    return () => {
      cancelled = true;
    };
  }, []);

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
      <DemoTopBar
        format="docx"
        tagline="In-browser .docx editor"
        fileName={source?.name}
      >
        <button
          type="button"
          className={barText}
          aria-pressed={compact}
          onClick={() => setCompact((value) => !value)}
        >
          Compact toolbar
        </button>
        <button
          type="button"
          className={barText}
          aria-pressed={reviewEnabled}
          disabled={!reviewPlugin}
          onClick={() => setReviewEnabled((value) => !value)}
        >
          Review plugin
        </button>
        <button
          type="button"
          className={barText}
          aria-pressed={reviewWrites}
          disabled={!reviewEnabled}
          onClick={() => setReviewWrites((value) => !value)}
        >
          Plugin write access
        </button>
        <CollaborationControls
          status={collab.status}
          synced={collab.synced}
          peerCount={collab.peerCount}
          error={collab.error}
          shared={session.status === "shared"}
        />
      </DemoTopBar>
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
            onError={(cause, info) => {
              if (isFatalDocxError(info)) setError(cause.message);
              else console.warn(cause);
            }}
            plugins={plugins}
            pluginGrants={pluginGrants}
            onPluginError={(failure) =>
              console.error(
                `Plugin ${failure.pluginId} failed (${failure.phase})`,
                failure.error,
              )
            }
            toolbar={
              compact ? (
                <CompactToolbar
                  onShare={() =>
                    void navigator.clipboard?.writeText(window.location.href)
                  }
                />
              ) : undefined
            }
            showToolbar
            showRuler
            showZoomControl
          />
        ) : (
          <DocumentLoading />
        )}
      </main>
    </div>
  );
}
