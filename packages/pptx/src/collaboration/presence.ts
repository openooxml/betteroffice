import type {
  PptxPresenceCursor,
  PptxPresencePeer,
  PptxPresenceState,
  PptxPresenceUser,
} from './types';

export const PRESENCE_CURSOR_INTERVAL_MS = 80;
export const PRESENCE_HEARTBEAT_MS = 20_000;
export const PRESENCE_EXPIRY_MS = 45_000;
export const PRESENCE_LABEL_DURATION_MS = 3_000;

export const PRESENCE_COLORS = [
  '#B3261E',
  '#7B1FA2',
  '#3949AB',
  '#00796B',
  '#2E7D32',
  '#A84400',
  '#C2185B',
  '#00695C',
] as const;

export interface AwarenessUpdateEntry {
  clientId: number;
  clock: number;
  state: PptxPresenceState | null;
}

export function presenceColorForClientId(clientId: number): string {
  return PRESENCE_COLORS[Math.abs(clientId) % PRESENCE_COLORS.length];
}

export function presenceUser(
  clientId: number,
  user?: { name: string; color?: string }
): PptxPresenceUser {
  if (user && (typeof user.name !== 'string' || user.name.trim().length === 0)) {
    throw new TypeError('Collaboration user name must be a non-empty string');
  }
  if (user?.color !== undefined && typeof user.color !== 'string') {
    throw new TypeError('Collaboration user color must be a string');
  }
  const name = user?.name.trim() || `Guest ${clientId.toString(36).toUpperCase()}`;
  const color = user?.color?.trim() || presenceColorForClientId(clientId);
  return { name, color };
}

export function samePresenceCursor(
  left: PptxPresenceCursor | null,
  right: PptxPresenceCursor | null
): boolean {
  return left?.slideId === right?.slideId && left?.shapeId === right?.shapeId;
}

export class PresencePeers {
  private readonly states = new Map<number, PptxPresencePeer>();
  private readonly clocks = new Map<number, number>();

  constructor(private readonly localClientId: number) {}

  get peers(): readonly PptxPresencePeer[] {
    return [...this.states.values()]
      .sort((left, right) => left.state.clientId - right.state.clientId)
      .map(copyPeer);
  }

  apply(entries: readonly AwarenessUpdateEntry[], now: number): boolean {
    let changed = false;
    for (const entry of entries) {
      if (entry.clientId === this.localClientId) continue;
      const currentClock = this.clocks.get(entry.clientId) ?? -1;
      if (entry.clock <= currentClock) continue;
      this.clocks.set(entry.clientId, entry.clock);

      if (!entry.state) {
        changed = this.states.delete(entry.clientId) || changed;
        continue;
      }

      const current = this.states.get(entry.clientId);
      const cursorMovedAt =
        current && samePresenceCursor(current.state.cursor, entry.state.cursor)
          ? current.cursorMovedAt
          : now;
      this.states.set(entry.clientId, {
        state: copyState(entry.state),
        lastSeen: now,
        cursorMovedAt,
      });
      changed = true;
    }
    return changed;
  }

  expire(now: number, maxAge = PRESENCE_EXPIRY_MS): boolean {
    let changed = false;
    for (const [clientId, peer] of this.states) {
      if (now - peer.lastSeen < maxAge) continue;
      this.states.delete(clientId);
      changed = true;
    }
    return changed;
  }

  clear(): boolean {
    const changed = this.states.size > 0;
    this.states.clear();
    this.clocks.clear();
    return changed;
  }
}

function copyState(state: PptxPresenceState): PptxPresenceState {
  return {
    clientId: state.clientId,
    clock: state.clock,
    user: { ...state.user },
    cursor: state.cursor ? { ...state.cursor } : null,
  };
}

function copyPeer(peer: PptxPresencePeer): PptxPresencePeer {
  return {
    state: copyState(peer.state),
    lastSeen: peer.lastSeen,
    cursorMovedAt: peer.cursorMovedAt,
  };
}
