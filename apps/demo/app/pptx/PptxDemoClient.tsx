"use client";

import dynamic from "next/dynamic";
import Link from "next/link";
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

const PptxEditor = dynamic(
  () => import("@betteroffice/pptx-react").then((module) => module.PptxEditor),
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
      <header className="z-2 flex items-center gap-3.5 border-b border-hairline bg-white/92 px-4 py-[11px] backdrop-blur-lg">
        <div className="flex min-w-0 items-baseline gap-2.5">
          <Link
            href="/"
            className="inline-flex items-baseline gap-2 text-[14px] font-[650] tracking-[-0.01em] whitespace-nowrap text-fg no-underline"
          >
            <Logo height={18} className="self-center" />
            BetterOffice <span className="font-normal text-faint">/ pptx</span>
          </Link>
          <span className="overflow-hidden text-[12.5px] text-ellipsis whitespace-nowrap text-mute">
            In-browser presentation editor
          </span>
        </div>

        <div className="flex-1" />

        {assets && (
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
        data-testid="pptx-demo-stage"
      >
        {error ? (
          <p className="m-auto text-mute" role="alert">
            Failed to load the demo presentation: {error}
          </p>
        ) : assets ? (
          <PptxEditor
            file={assets.file}
            fonts={assets.fonts}
            collaboration={collaboration}
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
