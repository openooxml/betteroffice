import { useEffect, useState } from 'react';
import { rangeRect } from '@betteroffice/xlsx';
import type { GridMeta } from '@betteroffice/xlsx';
import {
  AWARENESS_LABEL_DURATION_MS,
  resolveAwarenessCursor,
  type AwarenessPeer,
} from '@betteroffice/xlsx/collaboration';

interface PresenceStripProps {
  peers: readonly AwarenessPeer[];
  sheetIds: readonly string[];
  sheetNames: readonly string[];
  activeSheet: number;
}

interface RemoteSelectionsProps {
  peers: readonly AwarenessPeer[];
  grid: GridMeta | undefined;
  sheetIds: readonly string[];
  activeSheet: number;
  zoom: number;
}

function initials(name: string): string {
  const words = name.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return '?';
  return words
    .slice(0, 2)
    .map((word) => word[0])
    .join('')
    .toUpperCase();
}

function translucent(color: string): string {
  const red = Number.parseInt(color.slice(1, 3), 16);
  const green = Number.parseInt(color.slice(3, 5), 16);
  const blue = Number.parseInt(color.slice(5, 7), 16);
  return `rgba(${red}, ${green}, ${blue}, 0.1)`;
}

function useLabelClock(peers: readonly AwarenessPeer[]): number {
  const [now, setNow] = useState(() => Date.now());

  useEffect(() => {
    const current = Date.now();
    setNow(current);
    const deadlines = peers
      .map((peer) => peer.cursorMovedAt + AWARENESS_LABEL_DURATION_MS)
      .filter((deadline) => deadline > current);
    if (deadlines.length === 0) return;
    const timer = setTimeout(
      () => setNow(Date.now()),
      Math.min(...deadlines) - current + 1
    );
    const handle = timer as unknown as { unref?: () => void };
    handle.unref?.();
    return () => clearTimeout(timer);
  }, [peers]);

  return now;
}

export function PresenceStrip({
  peers,
  sheetIds,
  sheetNames,
  activeSheet,
}: PresenceStripProps) {
  if (peers.length === 0) return null;

  return (
    <div
      data-testid="xlsx-presence-strip"
      aria-label="Collaborators"
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: 4,
        padding: '0 6px',
        flex: '0 0 auto',
      }}
    >
      {peers.map((peer) => {
        const peerSheet = peer.cursor ? sheetIds.indexOf(peer.cursor.sheet) : activeSheet;
        const onOtherSheet = peer.cursor !== null && peerSheet !== activeSheet;
        const sheetLabel =
          peerSheet >= 0 ? sheetNames[peerSheet] : peer.cursor ? 'Unavailable sheet' : undefined;
        const title =
          onOtherSheet || peerSheet < 0
            ? `${peer.user.name} — ${sheetLabel}`
            : peer.user.name;
        return (
          <span
            key={peer.clientId}
            data-testid={`xlsx-presence-chip-${peer.clientId}`}
            title={title}
            aria-label={title}
            style={{
              display: 'inline-grid',
              placeItems: 'center',
              width: 24,
              height: 24,
              borderRadius: '50%',
              background: peer.user.color,
              color: '#ffffff',
              boxShadow: '0 0 0 2px #ffffff',
              fontSize: 9,
              fontWeight: 700,
              lineHeight: 1,
              letterSpacing: '-0.02em',
              opacity: onOtherSheet || peerSheet < 0 ? 0.42 : 1,
              transition: 'opacity 120ms ease',
            }}
          >
            {initials(peer.user.name)}
          </span>
        );
      })}
    </div>
  );
}

export function RemoteSelections({
  peers,
  grid,
  sheetIds,
  activeSheet,
  zoom,
}: RemoteSelectionsProps) {
  const now = useLabelClock(peers);
  if (!grid) return null;

  return peers.map((peer) => {
    if (!peer.cursor) return null;
    const resolved = resolveAwarenessCursor(peer.cursor, sheetIds);
    if (!resolved || resolved.sheetIndex !== activeSheet) return null;
    const rect = rangeRect(grid, resolved.range);
    if (!rect) return null;
    const scaled = {
      x: rect.x * zoom,
      y: rect.y * zoom,
      w: rect.w * zoom,
      h: rect.h * zoom,
    };
    const isRange =
      resolved.range.top !== resolved.range.bottom ||
      resolved.range.left !== resolved.range.right;
    const leftClipped = resolved.range.left < grid.startCol;
    const topClipped = resolved.range.top < grid.startRow || scaled.y < 18;
    const showLabel =
      peer.cursorMovedAt > 0 &&
      now - peer.cursorMovedAt < AWARENESS_LABEL_DURATION_MS;

    return (
      <div
        key={peer.clientId}
        data-testid={`xlsx-remote-selection-${peer.clientId}`}
        style={{
          position: 'absolute',
          left: scaled.x,
          top: scaled.y,
          width: scaled.w,
          height: scaled.h,
          boxSizing: 'border-box',
          border: `2px solid ${peer.user.color}`,
          background: isRange ? translucent(peer.user.color) : 'transparent',
        }}
      >
        {showLabel && (
          <span
            data-testid={`xlsx-remote-label-${peer.clientId}`}
            style={{
              position: 'absolute',
              top: topClipped ? 0 : -18,
              ...(leftClipped ? { right: -2 } : { left: -2 }),
              maxWidth: 180,
              height: 18,
              padding: '2px 6px',
              borderRadius: topClipped ? '0 0 3px 3px' : '3px 3px 0 0',
              boxSizing: 'border-box',
              overflow: 'hidden',
              background: peer.user.color,
              color: '#ffffff',
              fontSize: 10,
              fontWeight: 600,
              lineHeight: '14px',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap',
            }}
          >
            {peer.user.name}
          </span>
        )}
      </div>
    );
  });
}
