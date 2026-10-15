"use client";

import dynamic from "next/dynamic";
import { useCallback, useEffect, useMemo, useState } from "react";
import {
  CollaborationProvider,
  type CollaborationUser,
  type PptxFontFace,
} from "@betteroffice/pptx";
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
} from "@betteroffice/fonts";
import type { PptxLegacyEditorProps, PptxPlugin, PptxPluginGrant } from "@betteroffice/pptx-react";
import { DemoTopBar, barText } from "../components/DemoTopBar";
import {
  CollaborationControls,
  COLLAB_RELAY_ORIGIN,
  useCollabRoom,
  useDemoRoom,
  type CollaborationReplica,
  type CollaborationTransport,
} from "../collab";

const PptxEditor = dynamic<PptxLegacyEditorProps & { experimentalWorkerOpen: false }>(
  () => import("@betteroffice/pptx-react").then((module) => module.PptxEditor),
  { ssr: false },
);
const CompactToolbar = dynamic(
  () => import("./CompactToolbar").then((module) => module.CompactToolbar),
  { ssr: false },
);

const SHOWCASE = {
  url: "/betteroffice-demo.pptx",
  name: "betteroffice-demo.pptx",
};

const PRESENCE_USER_KEY = "betteroffice:pptx:presence-user";
const PRESENCE_ADJECTIVES = [
  "Brisk",
  "Calm",
  "Clever",
  "Kind",
  "Lively",
  "Merry",
  "Sunny",
  "Swift",
] as const;
const PRESENCE_ANIMALS = [
  "Badger",
  "Dolphin",
  "Falcon",
  "Fox",
  "Otter",
  "Panda",
  "Robin",
  "Turtle",
] as const;

type DemoAssets = {
  file: Uint8Array;
  fonts: PptxFontFace[];
  seed: Uint8Array;
};

export function PptxDemoClient() {
  const [assets, setAssets] = useState<DemoAssets | null>(null);
  const [user, setUser] = useState<CollaborationUser | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [compact, setCompact] = useState(false);
  const [reviewPlugin, setReviewPlugin] = useState<PptxPlugin | null>(null);
  const [reviewEnabled, setReviewEnabled] = useState(false);
  const [reviewWrites, setReviewWrites] = useState(false);
  const [readOnly, setReadOnly] = useState(false);
  const plugins = useMemo(
    () => (reviewEnabled && reviewPlugin ? [reviewPlugin] : []),
    [reviewEnabled, reviewPlugin],
  );
  const pluginGrants = useMemo<Record<string, PptxPluginGrant> | undefined>(
    () =>
      reviewWrites && reviewPlugin
        ? { [reviewPlugin.id]: { document: "write", editBatches: true } }
        : undefined,
    [reviewPlugin, reviewWrites],
  );
  const room = useDemoRoom();
  const createProvider = useCallback(
    (replica: CollaborationReplica, transport: CollaborationTransport) =>
      new CollaborationProvider(replica, transport, {
        user: user ?? undefined,
      }),
    [user],
  );
  const collab = useCollabRoom(COLLAB_RELAY_ORIGIN, room, createProvider);

  useEffect(() => setUser(loadPresenceUser()), []);

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
    Promise.all([
      loadPresentation(),
      loadPresentationFonts(),
      loadCollaborationSeed(),
    ]).then(
      ([file, fonts, seed]) => {
        if (!cancelled) setAssets({ file, fonts, seed });
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

  const collaboration = useMemo(
    () =>
      room && assets && collab.clientId && user
        ? {
            clientId: collab.clientId,
            initialUpdate: assets.seed,
            onReplica: collab.onReplica,
            presence: collab.provider ?? undefined,
          }
        : undefined,
    [assets, collab.clientId, collab.onReplica, collab.provider, room, user],
  );

  return (
    <div className="fixed inset-0 z-20 flex flex-col bg-surface text-fg">
      <DemoTopBar
        format="pptx"
        tagline="In-browser presentation editor"
        fileName={assets ? SHOWCASE.name : null}
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
        <button
          type="button"
          className={barText}
          aria-pressed={readOnly}
          onClick={() => setReadOnly((value) => !value)}
        >
          Read-only
        </button>
        <CollaborationControls
          status={collab.status}
          synced={collab.synced}
          peerCount={collab.peerCount}
          error={collab.error}
        />
      </DemoTopBar>
      <main
        className="flex min-h-0 flex-1 flex-col *:min-h-0 *:flex-1"
        data-testid="pptx-demo-stage"
      >
        {error ? (
          <p className="m-auto text-mute" role="alert">
            Failed to load the demo presentation: {error}
          </p>
        ) : assets ? (
          <PptxEditor legacySyncApi experimentalWorkerOpen={false}
            file={assets.file}
            fonts={assets.fonts}
            collaboration={collaboration}
            readOnly={readOnly}
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
          />
        ) : (
          <p className="m-auto text-mute">Loading presentation…</p>
        )}
      </main>
    </div>
  );
}

function loadPresenceUser(): CollaborationUser {
  try {
    const stored = sessionStorage.getItem(PRESENCE_USER_KEY)?.trim();
    if (stored) return { name: stored };
    const name = generatePresenceName();
    sessionStorage.setItem(PRESENCE_USER_KEY, name);
    return { name };
  } catch {
    return { name: generatePresenceName() };
  }
}

function generatePresenceName(): string {
  const values = crypto.getRandomValues(new Uint32Array(2));
  const adjective = PRESENCE_ADJECTIVES[values[0] % PRESENCE_ADJECTIVES.length];
  const animal = PRESENCE_ANIMALS[values[1] % PRESENCE_ANIMALS.length];
  return `${adjective} ${animal}`;
}

async function loadPresentation(): Promise<Uint8Array> {
  const response = await fetch(SHOWCASE.url);
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
  return new Uint8Array(await response.arrayBuffer());
}

async function loadCollaborationSeed(): Promise<Uint8Array> {
  const response = await fetch("/seeds/pptx.bin");
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
