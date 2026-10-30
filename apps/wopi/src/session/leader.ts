/** Minimal user shape needed to elect a save leader. */
export interface LeaderCandidate {
  id?: string;
  userId?: string;
  joinedAt: number;
  connected: boolean;
  canWrite: boolean;
  expiresAt?: number | null;
}

export const LEADER_MIN_TOKEN_MS = 2 * 60 * 1000;

/** Earliest joined connected writer whose token stays valid for two minutes. */
export function electLeader<T extends LeaderCandidate>(users: readonly T[], now: number): T | null {
  let best: T | null = null;
  for (const user of users) {
    if (!user.connected || !user.canWrite) continue;
    const expiresAt = user.expiresAt;
    if (expiresAt !== null && expiresAt !== undefined && expiresAt !== 0 && expiresAt - now < LEADER_MIN_TOKEN_MS) continue;
    if (best === null || user.joinedAt < best.joinedAt) best = user;
  }
  return best;
}
