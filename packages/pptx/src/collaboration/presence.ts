import type {
  PptxPresenceCursor,
  PptxPresencePeer,
  PptxPresenceState,
  PptxPresenceUser,
} from './types';

export const PRESENCE_CURSOR_INTERVAL_MS = 80;
export const PRESENCE_HEARTBEAT_MS = 20_000;
export const PRESENCE_EXPIRY_MS = 45_000;
export const PRESENCE_CLOCK_RETENTION_MS = 5 * 60_000;
export const PRESENCE_LABEL_DURATION_MS = 3_000;
export const MAX_AWARENESS_STRING_LENGTH = 1024;
/** Full stores evict the least recently accepted client id first. */
export const MAX_TRACKED_PRESENCE_IDS = 4096;

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

const PRESENCE_COLOR_PATTERN = /^#[0-9a-fA-F]+$/;

export interface AwarenessUpdateEntry {
  clientId: number;
  clock: number;
  state: PptxPresenceState | null;
}

export interface PresencePeersOptions {
  maxTrackedIds?: number;
}

interface PresenceClock {
  clock: number;
  updatedAt: number;
}

export function presenceColorForClientId(clientId: number): string {
  return PRESENCE_COLORS[Math.abs(clientId) % PRESENCE_COLORS.length];
}

export function sanitizePresenceColor(clientId: number, color: unknown): string {
  return typeof color === 'string' &&
    (color.length === 4 || color.length === 7) &&
    PRESENCE_COLOR_PATTERN.test(color)
    ? color
    : presenceColorForClientId(clientId);
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
  const name =
    user?.name.trim().slice(0, MAX_AWARENESS_STRING_LENGTH) ||
    `Guest ${clientId.toString(36).toUpperCase()}`;
  const color = sanitizePresenceColor(clientId, user?.color);
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
  private readonly clocks = new Map<number, PresenceClock>();
  private readonly maxTrackedIds: number;

  constructor(
    private readonly localClientId: number,
    options: PresencePeersOptions = {}
  ) {
    this.maxTrackedIds = options.maxTrackedIds ?? MAX_TRACKED_PRESENCE_IDS;
    if (!Number.isSafeInteger(this.maxTrackedIds) || this.maxTrackedIds < 1) {
      throw new RangeError('Presence tracked id limit must be a positive integer');
    }
  }

  get peers(): readonly PptxPresencePeer[] {
    return [...this.states.values()]
      .sort((left, right) => left.state.clientId - right.state.clientId)
      .map(copyPeer);
  }

  get trackedIdCount(): number {
    return this.clocks.size;
  }

  get nextExpiryAt(): number | undefined {
    let next: number | undefined;
    for (const peer of this.states.values()) {
      const expiresAt = peer.lastSeen + PRESENCE_EXPIRY_MS;
      next = next === undefined ? expiresAt : Math.min(next, expiresAt);
    }
    for (const record of this.clocks.values()) {
      const expiresAt = record.updatedAt + PRESENCE_CLOCK_RETENTION_MS;
      next = next === undefined ? expiresAt : Math.min(next, expiresAt);
    }
    return next;
  }

  apply(entries: readonly AwarenessUpdateEntry[], now: number): boolean {
    let changed = this.expire(now);
    for (const entry of entries) {
      if (entry.clientId === this.localClientId) continue;
      const currentClock = this.clocks.get(entry.clientId);
      if (currentClock && entry.clock <= currentClock.clock) continue;
      if (!currentClock && this.clocks.size >= this.maxTrackedIds) {
        changed = this.evictOldest() || changed;
      }
      this.clocks.delete(entry.clientId);
      this.clocks.set(entry.clientId, { clock: entry.clock, updatedAt: now });

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
    for (const [clientId, record] of this.clocks) {
      if (now - record.updatedAt < PRESENCE_CLOCK_RETENTION_MS) continue;
      this.clocks.delete(clientId);
      changed = this.states.delete(clientId) || changed;
    }
    return changed;
  }

  clear(): boolean {
    const changed = this.states.size > 0;
    this.states.clear();
    this.clocks.clear();
    return changed;
  }

  private evictOldest(): boolean {
    const oldest = this.clocks.keys().next();
    if (oldest.done) return false;
    this.clocks.delete(oldest.value);
    return this.states.delete(oldest.value);
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
