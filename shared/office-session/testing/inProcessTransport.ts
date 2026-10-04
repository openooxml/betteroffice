import type { SessionTransport } from '../transport';
import { SessionFailure } from '../types';

export function createInProcessPair(): { client: SessionTransport; host: SessionTransport } {
  let closed = false;
  const timers = new Set<ReturnType<typeof setTimeout>>();
  const listeners = [new Set<(message: unknown) => void>(), new Set<(message: unknown) => void>()];
  const errors = [new Set<(error: unknown) => void>(), new Set<(error: unknown) => void>()];
  const endpoint = (side: number): SessionTransport => ({
    post(message, transfer = []) {
      if (closed) throw new SessionFailure('disposed', 'Session transport is closed');
      const cloned = structuredClone(message, { transfer });
      const timer = setTimeout(() => {
        timers.delete(timer);
        const receiving = listeners[1 - side]!;
        for (const listener of [...receiving]) {
          if (closed) break;
          if (receiving.has(listener)) listener(cloned);
        }
      }, 0);
      timers.add(timer);
    },
    listen(listener) {
      if (!closed) listeners[side]!.add(listener);
      return () => { listeners[side]!.delete(listener); };
    },
    onError(listener) {
      if (!closed) errors[side]!.add(listener);
      return () => { errors[side]!.delete(listener); };
    },
    close() {
      closed = true;
      for (const timer of timers) clearTimeout(timer);
      timers.clear();
      for (const set of [...listeners, ...errors]) set.clear();
    },
  });
  return { client: endpoint(0), host: endpoint(1) };
}
