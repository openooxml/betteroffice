import type { DisplayListQueries } from '@betteroffice/docx/layout/render';

export interface DisplayListQuerySnapshot {
  queries: DisplayListQueries;
  frameEpoch: number | null;
}

export type ResolveDisplayListQueries = (
  minimumFrameEpoch?: number | null
) => Promise<DisplayListQuerySnapshot | null>;

interface QueryWaiter {
  minimumFrameEpoch: number | null;
  resolve(snapshot: DisplayListQuerySnapshot | null): void;
}

function satisfies(
  snapshot: DisplayListQuerySnapshot,
  minimumFrameEpoch: number | null
): boolean {
  return (
    minimumFrameEpoch === null ||
    (snapshot.frameEpoch !== null && snapshot.frameEpoch >= minimumFrameEpoch)
  );
}

export class DisplayListQueryEpochGate {
  private current: DisplayListQuerySnapshot | null = null;
  private waiters: QueryWaiter[] = [];
  private state: 'unavailable' | 'pending' | 'ready' = 'unavailable';

  resolve(minimumFrameEpoch: number | null = null): Promise<DisplayListQuerySnapshot | null> {
    if (this.current && satisfies(this.current, minimumFrameEpoch)) {
      return Promise.resolve(this.current);
    }
    if (this.state === 'unavailable') return Promise.resolve(null);
    return new Promise((resolve) => {
      this.waiters.push({ minimumFrameEpoch, resolve });
    });
  }

  invalidate(): void {
    this.current = null;
    this.state = 'pending';
  }

  publish(snapshot: DisplayListQuerySnapshot): void {
    this.current = snapshot;
    this.state = 'ready';
    const pending = this.waiters;
    this.waiters = [];
    for (const waiter of pending) {
      if (satisfies(snapshot, waiter.minimumFrameEpoch)) waiter.resolve(snapshot);
      else this.waiters.push(waiter);
    }
  }

  clear(): void {
    this.current = null;
    this.state = 'unavailable';
    for (const waiter of this.waiters) waiter.resolve(null);
    this.waiters = [];
  }
}
