import type {
  AwarenessCursor,
  AwarenessPayload,
  AwarenessPeer,
  AwarenessUpdate,
  CollaborationUser,
  CollaborationUserOptions,
} from './types';

export const AWARENESS_BROADCAST_INTERVAL_MS = 80;
export const AWARENESS_HEARTBEAT_INTERVAL_MS = 20_000;
export const AWARENESS_PEER_TIMEOUT_MS = 45_000;
export const AWARENESS_LABEL_DURATION_MS = 3_000;
export const XLSX_MAX_ROWS = 1_048_576;
export const XLSX_MAX_COLUMNS = 16_384;

export const AWARENESS_COLORS = [
  '#0B57D0',
  '#B3261E',
  '#146C2E',
  '#6B4EFF',
  '#8E24AA',
  '#00639B',
  '#8E4E00',
  '#7D5260',
] as const;

export interface AwarenessPeerEntry {
  clock: number;
  state: AwarenessPayload | null;
  lastSeen: number;
  cursorMovedAt: number;
}

export type AwarenessPeerStore = Map<number, AwarenessPeerEntry>;

export interface ResolvedAwarenessCursor {
  sheetIndex: number;
  range: {
    top: number;
    left: number;
    bottom: number;
    right: number;
  };
}

function cloneCursor(cursor: AwarenessCursor | null): AwarenessCursor | null {
  if (!cursor) return null;
  return {
    sheet: cursor.sheet,
    anchor: { ...cursor.anchor },
    head: { ...cursor.head },
  };
}

function clonePayload(payload: AwarenessPayload): AwarenessPayload {
  return {
    user: { ...payload.user },
    cursor: cloneCursor(payload.cursor),
  };
}

function cursorsEqual(left: AwarenessCursor | null, right: AwarenessCursor | null): boolean {
  if (!left || !right) return left === right;
  return (
    left.sheet === right.sheet &&
    left.anchor.row === right.anchor.row &&
    left.anchor.col === right.anchor.col &&
    left.head.row === right.head.row &&
    left.head.col === right.head.col
  );
}

function payloadsEqual(left: AwarenessPayload | null, right: AwarenessPayload | null): boolean {
  if (!left || !right) return left === right;
  return (
    left.user.name === right.user.name &&
    left.user.color === right.user.color &&
    cursorsEqual(left.cursor, right.cursor)
  );
}

export function colorForClientId(clientId: number): string {
  return AWARENESS_COLORS[clientId % AWARENESS_COLORS.length];
}

export function normalizeAwarenessColor(color: unknown, clientId: number): string {
  if (typeof color !== 'string' || !/^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i.test(color)) {
    return colorForClientId(clientId);
  }
  const hex = color.slice(1).toUpperCase();
  if (hex.length === 6) return `#${hex}`;
  return `#${hex[0]}${hex[0]}${hex[1]}${hex[1]}${hex[2]}${hex[2]}`;
}

export function normalizeCollaborationUser(
  user: CollaborationUserOptions | undefined,
  clientId: number
): CollaborationUser {
  const name = user?.name.trim().slice(0, 128) || 'Anonymous';
  return {
    name,
    color: normalizeAwarenessColor(user?.color, clientId),
  };
}

export function applyAwarenessUpdates(
  store: AwarenessPeerStore,
  updates: readonly AwarenessUpdate[],
  localClientId: number,
  now: number
): boolean {
  let changed = false;
  for (const update of updates) {
    if (update.clientId === localClientId) continue;
    const current = store.get(update.clientId);
    if (current && update.clock < current.clock) continue;
    if (current && update.clock === current.clock) {
      current.lastSeen = now;
      continue;
    }

    const cursorMoved =
      update.state?.cursor != null &&
      (!current?.state || !cursorsEqual(current.state.cursor, update.state.cursor));
    const nextState = update.state ? clonePayload(update.state) : null;
    changed ||= !payloadsEqual(current?.state ?? null, nextState);
    store.set(update.clientId, {
      clock: update.clock,
      state: nextState,
      lastSeen: now,
      cursorMovedAt: cursorMoved ? now : (current?.cursorMovedAt ?? 0),
    });
  }
  return changed;
}

export function expireAwarenessPeers(
  store: AwarenessPeerStore,
  now: number,
  timeout = AWARENESS_PEER_TIMEOUT_MS
): boolean {
  let changed = false;
  for (const [clientId, entry] of store) {
    if (now - entry.lastSeen < timeout) continue;
    changed ||= entry.state !== null;
    store.delete(clientId);
  }
  return changed;
}

export function awarenessPeers(store: AwarenessPeerStore): AwarenessPeer[] {
  const peers: AwarenessPeer[] = [];
  for (const [clientId, entry] of store) {
    if (!entry.state) continue;
    peers.push({
      clientId,
      clock: entry.clock,
      user: { ...entry.state.user },
      cursor: cloneCursor(entry.state.cursor),
      lastSeen: entry.lastSeen,
      cursorMovedAt: entry.cursorMovedAt,
    });
  }
  peers.sort((left, right) => left.clientId - right.clientId);
  return peers;
}

export function resolveAwarenessCursor(
  cursor: AwarenessCursor,
  sheetIds: readonly string[],
  rows = XLSX_MAX_ROWS,
  columns = XLSX_MAX_COLUMNS
): ResolvedAwarenessCursor | null {
  const sheetIndex = sheetIds.indexOf(cursor.sheet);
  if (sheetIndex < 0 || rows < 1 || columns < 1) return null;
  const cells = [cursor.anchor, cursor.head];
  if (
    cells.some(
      ({ row, col }) =>
        !Number.isSafeInteger(row) ||
        !Number.isSafeInteger(col) ||
        row < 0 ||
        col < 0 ||
        row >= rows ||
        col >= columns
    )
  ) {
    return null;
  }
  return {
    sheetIndex,
    range: {
      top: Math.min(cursor.anchor.row, cursor.head.row),
      left: Math.min(cursor.anchor.col, cursor.head.col),
      bottom: Math.max(cursor.anchor.row, cursor.head.row),
      right: Math.max(cursor.anchor.col, cursor.head.col),
    },
  };
}
