import type { CollaborationUpdateOrigin } from '../collaboration/types';
import type { EditSession } from './wasm/index';

/** @internal */
export interface EditSessionUpdateEvents {
  session: EditSession;
  subscribe(
    sink: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void;
  drain(): void;
}

type Sink = (update: Uint8Array, origin: CollaborationUpdateOrigin) => void;

const NO_DRAIN_AFTER = new Set([
  'free',
  'drain_update_event',
  'start_update_event_observation',
  'clear_update_event_observation',
]);

const hubs = new WeakMap<object, EditSessionUpdateEvents>();

/** @internal */
export function editSessionUpdateEvents(session: EditSession): EditSessionUpdateEvents {
  const existing = hubs.get(session);
  if (existing) return existing;

  const target = session as EditSession & Record<string, unknown>;
  const subscribers = new Set<{ sink: Sink }>();
  const wrappers = new Map<string, unknown>();
  let dead = false;
  let draining = false;

  const drain = (): void => {
    if (dead || draining || subscribers.size === 0) return;
    draining = true;
    try {
      while (!dead && subscribers.size > 0) {
        let event: Uint8Array;
        try {
          event = target.drain_update_event();
        } catch {
          return;
        }
        if (event.length === 0) return;
        const flag = event[0];
        if (flag !== 0 && flag !== 1) continue;
        const update = event.slice(1);
        const origin: CollaborationUpdateOrigin = flag === 1 ? 'remote' : 'local';
        for (const record of [...subscribers]) {
          if (dead || subscribers.size === 0) return;
          if (!subscribers.has(record)) continue;
          try {
            record.sink(update, origin);
          } catch {}
        }
      }
    } finally {
      draining = false;
    }
  };

  const subscribe = (sink: Sink): (() => void) => {
    if (dead) return () => {};
    if (subscribers.size === 0) target.start_update_event_observation();
    const record = { sink };
    subscribers.add(record);
    let active = true;
    return () => {
      if (!active) return;
      active = false;
      subscribers.delete(record);
      if (subscribers.size === 0 && !dead) target.clear_update_event_observation();
    };
  };

  const proxy = new Proxy(target, {
    get(raw, key) {
      if (typeof key !== 'string') return Reflect.get(raw, key, raw);
      const value = Reflect.get(raw, key, raw);
      if (typeof value !== 'function') return value;
      let wrapper = wrappers.get(key);
      if (wrapper) return wrapper;
      wrapper = function (this: unknown, ...args: unknown[]) {
        if (key === 'free') {
          if (dead) return undefined;
          dead = true;
          const free = Reflect.get(raw, key, raw) as (...a: unknown[]) => unknown;
          return free.apply(raw, args);
        }
        const method = Reflect.get(raw, key, raw) as (...a: unknown[]) => unknown;
        if (typeof method !== 'function') throw new TypeError(`${key} is not a function`);
        let result: unknown;
        try {
          result = method.apply(raw, args);
        } catch (error) {
          if (!NO_DRAIN_AFTER.has(key) && !key.endsWith('_update_observer')) {
            try { drain(); } catch {}
          }
          throw error;
        }
        if (!NO_DRAIN_AFTER.has(key) && !key.endsWith('_update_observer')) drain();
        return result;
      };
      wrappers.set(key, wrapper);
      return wrapper;
    },
  }) as EditSession;

  const hub: EditSessionUpdateEvents = { session: proxy, subscribe, drain };
  hubs.set(target, hub);
  hubs.set(proxy, hub);
  return hub;
}
