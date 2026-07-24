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
  timeout: ReturnType<typeof setTimeout>;
}

function satisfies(
  snapshot: DisplayListQuerySnapshot,
  minimumFrameEpoch: number | null
): boolean {
  return (
    minimumFrameEpoch === null ||
    snapshot.frameEpoch === null ||
    (snapshot.frameEpoch !== null && snapshot.frameEpoch >= minimumFrameEpoch)
  );
}

export class DisplayListQueryEpochGate {
  private current: DisplayListQuerySnapshot | null = null;
  private waiters: QueryWaiter[] = [];
  private state: 'unavailable' | 'pending' | 'ready' = 'unavailable';

  constructor(private readonly maximumWaitMs = 250) {}

  resolve(minimumFrameEpoch: number | null = null): Promise<DisplayListQuerySnapshot | null> {
    if (this.current && satisfies(this.current, minimumFrameEpoch)) {
      return Promise.resolve(this.current);
    }
    if (this.state === 'unavailable') return Promise.resolve(null);
    return new Promise((resolve) => {
      let waiter: QueryWaiter;
      waiter = {
        minimumFrameEpoch,
        resolve,
        timeout: setTimeout(() => {
          const index = this.waiters.indexOf(waiter);
          if (index < 0) return;
          this.waiters.splice(index, 1);
          resolve(this.current);
        }, this.maximumWaitMs),
      };
      this.waiters.push(waiter);
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
      if (satisfies(snapshot, waiter.minimumFrameEpoch)) {
        clearTimeout(waiter.timeout);
        waiter.resolve(snapshot);
      }
      else this.waiters.push(waiter);
    }
  }

  clear(): void {
    this.current = null;
    this.state = 'unavailable';
    for (const waiter of this.waiters) {
      clearTimeout(waiter.timeout);
      waiter.resolve(null);
    }
    this.waiters = [];
  }
}
