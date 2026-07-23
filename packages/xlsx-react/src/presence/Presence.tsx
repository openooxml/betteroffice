import { useEffect, useMemo, useState } from 'react';
import { rangeRect } from '@betteroffice/xlsx';
import type { CellRange, GridMeta, MergedRange } from '@betteroffice/xlsx';
import {
  AWARENESS_LABEL_DURATION_MS,
  normalizeAwarenessColor,
  resolveAwarenessCursor,
  type AwarenessPeer,
} from '@betteroffice/xlsx/collaboration';

const MAX_PRESENCE_CHIPS = 8;
const MAX_REMOTE_SELECTIONS = 32;

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
  mergedRanges: readonly MergedRange[];
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

function intersects(left: CellRange, right: CellRange): boolean {
  return (
    left.left <= right.right &&
    left.right >= right.left &&
    left.top <= right.bottom &&
    left.bottom >= right.top
  );
}

export function expandRangeToMergedCells(
  range: CellRange,
  mergedRanges: readonly MergedRange[]
): CellRange {
  const expanded = { ...range };
  let changed = true;
  while (changed) {
    changed = false;
    for (const merged of mergedRanges) {
      const mergedRange = {
        top: Math.min(merged.start.row, merged.end.row),
        left: Math.min(merged.start.col, merged.end.col),
        bottom: Math.max(merged.start.row, merged.end.row),
        right: Math.max(merged.start.col, merged.end.col),
      };
      if (!intersects(expanded, mergedRange)) continue;
      const top = Math.min(expanded.top, mergedRange.top);
      const left = Math.min(expanded.left, mergedRange.left);
      const bottom = Math.max(expanded.bottom, mergedRange.bottom);
      const right = Math.max(expanded.right, mergedRange.right);
      if (
        top === expanded.top &&
        left === expanded.left &&
        bottom === expanded.bottom &&
        right === expanded.right
      ) {
        continue;
      }
      Object.assign(expanded, { top, left, bottom, right });
      changed = true;
    }
  }
  return expanded;
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
  const visiblePeers = peers.slice(0, MAX_PRESENCE_CHIPS);
  const hiddenPeers = peers.length - visiblePeers.length;

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
      {visiblePeers.map((peer) => {
        const peerSheet = peer.cursor ? sheetIds.indexOf(peer.cursor.sheet) : activeSheet;
        const onOtherSheet = peer.cursor !== null && peerSheet !== activeSheet;
        const color = normalizeAwarenessColor(peer.user.color, peer.clientId);
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
              backgroundColor: color,
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
      {hiddenPeers > 0 && (
        <span
          data-testid="xlsx-presence-overflow"
          title={`${hiddenPeers} more collaborators`}
          aria-label={`${hiddenPeers} more collaborators`}
          style={{
            display: 'inline-grid',
            placeItems: 'center',
            minWidth: 24,
            height: 24,
            padding: '0 4px',
            borderRadius: 12,
            backgroundColor: '#5f6368',
            color: '#ffffff',
            fontSize: 9,
            fontWeight: 700,
            lineHeight: 1,
          }}
        >
          +{hiddenPeers}
        </span>
      )}
    </div>
  );
}

export function RemoteSelections({
  peers,
  grid,
  sheetIds,
  activeSheet,
  zoom,
  mergedRanges,
}: RemoteSelectionsProps) {
  const visiblePeers = useMemo(
    () => peers.slice(0, MAX_REMOTE_SELECTIONS),
    [peers]
  );
  const hiddenPeers = peers.length - visiblePeers.length;
  const now = useLabelClock(visiblePeers);
  if (!grid) return null;

  return (
    <>
      {visiblePeers.map((peer) => {
        if (!peer.cursor) return null;
        const resolved = resolveAwarenessCursor(peer.cursor, sheetIds);
        if (!resolved || resolved.sheetIndex !== activeSheet) return null;
        const renderedRange = expandRangeToMergedCells(resolved.range, mergedRanges);
        const rect = rangeRect(grid, renderedRange);
        if (!rect) return null;
        const color = normalizeAwarenessColor(peer.user.color, peer.clientId);
        const scaled = {
          x: rect.x * zoom,
          y: rect.y * zoom,
          w: rect.w * zoom,
          h: rect.h * zoom,
        };
        const isRange =
          resolved.range.top !== resolved.range.bottom ||
          resolved.range.left !== resolved.range.right;
        const leftClipped = renderedRange.left < grid.startCol;
        const topClipped = renderedRange.top < grid.startRow || scaled.y < 18;
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
              borderWidth: 2,
              borderStyle: 'solid',
              borderColor: color,
              backgroundColor: isRange ? translucent(color) : 'transparent',
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
                  backgroundColor: color,
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
      })}
      {hiddenPeers > 0 && (
        <span
          data-testid="xlsx-remote-selection-overflow"
          aria-hidden="true"
          style={{
            position: 'absolute',
            left: 4,
            top: 4,
            padding: '2px 5px',
            borderWidth: 1,
            borderStyle: 'solid',
            borderColor: '#c7cacf',
            borderRadius: 8,
            backgroundColor: '#ffffff',
            color: '#3c4043',
            fontSize: 9,
            lineHeight: 1,
          }}
        >
          +{hiddenPeers}
        </span>
      )}
    </>
  );
}
