import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  resolvePresenceColor,
  type CollaborationPeer,
  type CollaborationPresence,
} from '@betteroffice/docx/collaboration';
import type { DisplayList, DisplayListQueries } from '@betteroffice/docx/layout/render';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { YrsLoc, YrsSelection, YrsSession } from '@betteroffice/docx/yrs';

import { projectPageLocalRect } from '../internals/canvasProjection';
import {
  buildRemotePresencePageMetrics,
  clampRemoteSelectionRange,
  pageInRemotePresenceWindow,
  remotePresencePagePositionRange,
  remotePresencePageWindow,
  simplifyRemoteSelectionRects,
} from './remotePresenceGeometry';

const FLAG_VISIBLE_MS = 3_000;

interface ProjectedRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

interface PeerGeometry {
  peer: CollaborationPeer;
  caret: ProjectedRect | null;
  selection: ProjectedRect[];
}

export interface RemotePresenceOverlayProps {
  presence: CollaborationPresence;
  session: YrsSession;
  locToDisplayPosition(loc: YrsLoc): number | null;
  overlayTarget: HTMLElement;
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  displayListQueries: DisplayListQueries;
  displayListIdentity: DisplayList | null;
  displayListFrameEpoch: number | null;
  sidebarOpen: boolean;
  zoom: number;
}

function monotonicNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function inferredSelection(peer: CollaborationPeer): YrsSelection | null {
  const inferred = peer.inferredCursor;
  if (!inferred) return null;
  const loc = {
    story: inferred.story,
    paraId: inferred.paraId,
    offset: inferred.endOffset,
  };
  return { anchor: loc, head: loc };
}

function positionFor(
  selection: YrsSelection,
  locToDisplayPosition: (loc: YrsLoc) => number | null
): { anchor: number; head: number } | null {
  const anchor = locToDisplayPosition(selection.anchor);
  const head = locToDisplayPosition(selection.head);
  return anchor == null || head == null ? null : { anchor, head };
}

export function RemotePresenceOverlay({
  presence,
  session,
  locToDisplayPosition,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  displayListIdentity,
  displayListFrameEpoch,
  sidebarOpen,
  zoom,
}: RemotePresenceOverlayProps) {
  const [peers, setPeers] = useState<readonly CollaborationPeer[]>(presence.peers);
  const [geometry, setGeometry] = useState<PeerGeometry[]>([]);
  const [flagNow, setFlagNow] = useState(monotonicNow);
  const invalidStickyKeysRef = useRef(new Set<string>());

  useEffect(() => presence.onPeers(setPeers), [presence]);

  useEffect(() => {
    const liveKeys = new Set(peers.map((peer) => `${peer.clientId}:${peer.clock}`));
    for (const key of invalidStickyKeysRef.current) {
      if (!liveKeys.has(key)) invalidStickyKeysRef.current.delete(key);
    }
  }, [peers]);

  useEffect(() => {
    const deadlines = peers
      .filter((peer) => peer.cursorMovedAt > 0)
      .map((peer) => peer.cursorMovedAt + FLAG_VISIBLE_MS)
      .filter((deadline) => deadline > flagNow);
    if (deadlines.length === 0) return;
    const timer = setTimeout(
      () => setFlagNow(monotonicNow()),
      Math.max(0, Math.min(...deadlines) - monotonicNow())
    );
    (timer as ReturnType<typeof setTimeout> & { unref?: () => void }).unref?.();
    return () => clearTimeout(timer);
  }, [flagNow, peers]);

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!host) {
      setGeometry([]);
      return;
    }
    const pageMetrics = buildRemotePresencePageMetrics(displayListQueries.displayList, zoom);
    const scrollParent = findVerticalScrollParentOrRoot(host);
    const usesWindow =
      scrollParent === document.scrollingElement || scrollParent === document.documentElement;
    const scrollTarget: EventTarget = usesWindow ? window : scrollParent;

    const recompute = () => {
      if (!displayListQueries.isReady()) {
        setGeometry([]);
        return;
      }
      const column = host.firstElementChild as HTMLElement | null;
      if (!column) {
        setGeometry([]);
        return;
      }
      const viewportTop = usesWindow ? 0 : scrollParent.getBoundingClientRect().top;
      const viewportBottom =
        viewportTop + (usesWindow ? window.innerHeight : scrollParent.clientHeight);
      const pageWindow = remotePresencePageWindow(
        pageMetrics,
        column.getBoundingClientRect().top,
        viewportTop,
        viewportBottom
      );
      if (!pageWindow) {
        setGeometry([]);
        return;
      }
      const pagePositionRange = remotePresencePagePositionRange(
        displayListQueries.displayList,
        pageWindow
      );
      if (!pagePositionRange) {
        setGeometry([]);
        return;
      }
      const next: PeerGeometry[] = [];
      for (const peer of peers) {
        let selection = inferredSelection(peer);
        if (!selection && peer.cursor) {
          const invalidKey = `${peer.clientId}:${peer.clock}`;
          if (invalidStickyKeysRef.current.has(invalidKey)) continue;
          try {
            selection = session.resolveSelection(peer.cursor);
          } catch {
            invalidStickyKeysRef.current.add(invalidKey);
            continue;
          }
          if (!selection) {
            invalidStickyKeysRef.current.add(invalidKey);
            continue;
          }
        }
        if (!selection) continue;
        const positions = positionFor(selection, locToDisplayPosition);
        if (!positions) continue;

        const caretRect =
          positions.head >= pagePositionRange.from && positions.head <= pagePositionRange.to
            ? displayListQueries.caretRect(positions.head)
            : null;
        const caret =
          caretRect && pageInRemotePresenceWindow(caretRect.pageIndex, pageWindow)
            ? projectPageLocalRect(
                host,
                overlayTarget,
                displayListQueries,
                caretRect.pageIndex,
                caretRect.x,
                caretRect.y,
                0,
                caretRect.height
              )
            : null;
        const from = Math.min(positions.anchor, positions.head);
        const to = Math.max(positions.anchor, positions.head);
        const selectionRange = clampRemoteSelectionRange(pagePositionRange, from, to);
        const selectionRects = selectionRange
          ? simplifyRemoteSelectionRects(
              displayListQueries.rangeRects(selectionRange.from, selectionRange.to),
              pageWindow
            ).flatMap((rect) => {
              const projected = projectPageLocalRect(
                host,
                overlayTarget,
                displayListQueries,
                rect.pageIndex,
                rect.x,
                rect.y,
                rect.width,
                rect.height
              );
              return projected ? [projected] : [];
            })
          : [];
        if (!caret && selectionRects.length === 0) continue;
        next.push({
          peer,
          caret: caret
            ? {
                left: caret.left,
                top: caret.top,
                width: 2,
                height: Math.max(1, caret.height),
              }
            : null,
          selection: selectionRects,
        });
      }
      setGeometry(next);
    };

    recompute();
    let disposed = false;
    let frameId: number | null = null;
    const scheduleRecompute = () => {
      if (frameId !== null) return;
      frameId = requestAnimationFrame(() => {
        frameId = null;
        if (!disposed) recompute();
      });
    };
    if (!displayListQueries.isReady()) {
      void displayListQueries
        .whenReady()
        .then(() => {
          if (!disposed) scheduleRecompute();
        })
        .catch(() => {});
    }
    const observer = new ResizeObserver(scheduleRecompute);
    observer.observe(host);
    observer.observe(overlayTarget);
    scrollTarget.addEventListener('scroll', scheduleRecompute, {
      passive: true,
    });
    window.addEventListener('resize', scheduleRecompute);
    host.addEventListener('transitionend', scheduleRecompute);
    return () => {
      disposed = true;
      if (frameId !== null) cancelAnimationFrame(frameId);
      observer.disconnect();
      scrollTarget.removeEventListener('scroll', scheduleRecompute);
      window.removeEventListener('resize', scheduleRecompute);
      host.removeEventListener('transitionend', scheduleRecompute);
    };
  }, [
    canvasHostRef,
    displayListFrameEpoch,
    displayListIdentity,
    displayListQueries,
    locToDisplayPosition,
    overlayTarget,
    peers,
    session,
    sidebarOpen,
    zoom,
  ]);

  return createPortal(
    <div
      aria-hidden="true"
      data-testid="remote-presence-overlay"
      style={{
        position: 'absolute',
        inset: 0,
        pointerEvents: 'none',
        overflow: 'visible',
        zIndex: 30,
      }}
    >
      {geometry.flatMap(({ peer, caret, selection }) => {
        const color = resolvePresenceColor(peer.clientId, peer.user.color);
        const flagVisible =
          peer.cursorMovedAt > 0 &&
          flagNow - peer.cursorMovedAt < FLAG_VISIBLE_MS;
        return [
          ...selection.map((rect, index) => (
            <div
              key={`${peer.clientId}:selection:${index}`}
              style={{
                position: 'absolute',
                left: rect.left,
                top: rect.top,
                width: rect.width,
                height: rect.height,
                boxSizing: 'border-box',
                backgroundColor: `color-mix(in srgb, ${color} 18%, transparent)`,
                borderColor: `color-mix(in srgb, ${color} 28%, transparent)`,
                borderStyle: 'solid',
                borderWidth: 1,
              }}
            />
          )),
          caret ? (
            <div
              key={`${peer.clientId}:caret`}
              style={{
                position: 'absolute',
                left: caret.left - 1,
                top: caret.top,
                width: 2,
                height: caret.height,
                borderRadius: 1,
                backgroundColor: color,
              }}
            >
              <div
                style={{
                  position: 'absolute',
                  left: 0,
                  bottom: '100%',
                  maxWidth: flagVisible ? 180 : 0,
                  padding: flagVisible ? '3px 6px' : 0,
                  borderRadius: '4px 4px 4px 0',
                  overflow: 'hidden',
                  opacity: flagVisible ? 1 : 0,
                  color: '#FFFFFF',
                  backgroundColor: color,
                  font: '600 11px/16px system-ui, sans-serif',
                  whiteSpace: 'nowrap',
                  transform: flagVisible ? 'translateY(0)' : 'translateY(2px)',
                  transition:
                    'opacity 160ms ease, transform 160ms ease, max-width 160ms ease, padding 160ms ease',
                }}
              >
                {peer.user.name}
              </div>
            </div>
          ) : null,
        ];
      })}
    </div>,
    overlayTarget
  );
}
