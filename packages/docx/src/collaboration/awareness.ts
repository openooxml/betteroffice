import type {
  CollaborationAwarenessState,
  CollaborationCursor,
  CollaborationPeer,
  CollaborationResolvedUser,
  CollaborationTextInsertion,
} from './types';

export interface AwarenessUpdateEntry {
  clientId: number;
  clock: number;
  state: {
    user: CollaborationResolvedUser;
    cursor: CollaborationCursor | null;
  } | null;
}

export interface AwarenessRecord {
  clientId: number;
  clock: number;
  state: AwarenessUpdateEntry['state'];
  lastSeenAt: number;
  cursorMovedAt: number;
  inferredCursor: CollaborationTextInsertion | null;
}

export interface AwarenessReduction {
  records: ReadonlyMap<number, AwarenessRecord>;
  peersChanged: boolean;
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  return (
    left.byteLength === right.byteLength &&
    left.every((value, index) => value === right[index])
  );
}

function sameCursor(
  left: CollaborationCursor | null,
  right: CollaborationCursor | null
): boolean {
  return (
    left === right ||
    Boolean(
      left &&
        right &&
        left.story === right.story &&
        sameBytes(left.anchor, right.anchor) &&
        sameBytes(left.head, right.head)
    )
  );
}

function sameInference(
  left: CollaborationTextInsertion | null,
  right: CollaborationTextInsertion
): boolean {
  return (
    left?.clientId === right.clientId &&
    left.story === right.story &&
    left.paraId === right.paraId &&
    left.endOffset === right.endOffset
  );
}

function cloneCursor(cursor: CollaborationCursor | null): CollaborationCursor | null {
  return cursor
    ? {
        story: cursor.story,
        anchor: cursor.anchor.slice(),
        head: cursor.head.slice(),
      }
    : null;
}

function cloneState(state: AwarenessUpdateEntry['state']): AwarenessUpdateEntry['state'] {
  return state
    ? {
        user: { ...state.user },
        cursor: cloneCursor(state.cursor),
      }
    : null;
}

export function reduceAwarenessEntries(
  current: ReadonlyMap<number, AwarenessRecord>,
  entries: readonly AwarenessUpdateEntry[],
  localClientId: number,
  now: number
): AwarenessReduction {
  let records = current;
  let peersChanged = false;
  for (const entry of entries) {
    if (entry.clientId === localClientId) continue;
    const previous = records.get(entry.clientId);
    if (previous && entry.clock <= previous.clock) {
      if (entry.clock === previous.clock && now > previous.lastSeenAt) {
        if (records === current) records = new Map(current);
        (records as Map<number, AwarenessRecord>).set(entry.clientId, {
          ...previous,
          lastSeenAt: now,
        });
      }
      continue;
    }

    if (records === current) records = new Map(current);
    const cursorMoved =
      entry.state?.cursor != null &&
      (previous?.inferredCursor != null ||
        previous?.state == null ||
        !sameCursor(previous.state.cursor, entry.state.cursor));
    (records as Map<number, AwarenessRecord>).set(entry.clientId, {
      clientId: entry.clientId,
      clock: entry.clock,
      state: cloneState(entry.state),
      lastSeenAt: now,
      cursorMovedAt: cursorMoved ? now : (previous?.cursorMovedAt ?? 0),
      inferredCursor: null,
    });
    peersChanged = true;
  }
  return { records, peersChanged };
}

export function reduceTypingInference(
  current: ReadonlyMap<number, AwarenessRecord>,
  inference: CollaborationTextInsertion,
  localClientId: number,
  now: number
): AwarenessReduction {
  if (inference.clientId === localClientId) {
    return { records: current, peersChanged: false };
  }
  const previous = current.get(inference.clientId);
  const records = new Map(current);
  records.set(inference.clientId, {
    clientId: inference.clientId,
    clock: previous?.clock ?? -1,
    state: previous?.state ?? null,
    lastSeenAt: previous?.lastSeenAt ?? now,
    cursorMovedAt: sameInference(previous?.inferredCursor ?? null, inference)
      ? (previous?.cursorMovedAt ?? now)
      : now,
    inferredCursor: { ...inference },
  });
  return { records, peersChanged: previous?.state != null };
}

export function expireAwarenessRecords(
  current: ReadonlyMap<number, AwarenessRecord>,
  now: number,
  maxAgeMs: number
): AwarenessReduction {
  let records = current;
  let peersChanged = false;
  for (const [clientId, record] of current) {
    if (now - record.lastSeenAt < maxAgeMs) continue;
    if (records === current) records = new Map(current);
    (records as Map<number, AwarenessRecord>).delete(clientId);
    if (record.state) peersChanged = true;
  }
  return { records, peersChanged };
}

export function awarenessPeers(
  records: ReadonlyMap<number, AwarenessRecord>
): CollaborationPeer[] {
  const peers: CollaborationPeer[] = [];
  for (const record of records.values()) {
    if (!record.state) continue;
    const state: CollaborationAwarenessState = {
      clientId: record.clientId,
      clock: record.clock,
      user: { ...record.state.user },
      cursor: cloneCursor(record.state.cursor),
    };
    peers.push({
      ...state,
      lastSeenAt: record.lastSeenAt,
      cursorMovedAt: record.cursorMovedAt,
      inferredCursor: record.inferredCursor ? { ...record.inferredCursor } : null,
    });
  }
  peers.sort((left, right) => left.clientId - right.clientId);
  return peers;
}
