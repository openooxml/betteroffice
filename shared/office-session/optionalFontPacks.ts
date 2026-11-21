import { SessionFailure } from './types';
import { OFFICE_SESSION_SILENCE_MS } from './client';

export type OptionalFontPack = 'intl' | 'cjk';
export type FontPackUrls = Record<string, string>;
export type OptionalFontPackSource = (pack: OptionalFontPack) => Promise<FontPackUrls | undefined>;

export type FontPackRequest = { protocol: 1; kind: 'font-pack-request'; id: string; pack: OptionalFontPack };
export type FontPackReply = { protocol: 1; kind: 'font-pack-reply'; id: string; urls?: FontPackUrls; error?: string };

export function isFontPackRequest(message: unknown): message is FontPackRequest {
  const value = message as Partial<FontPackRequest> | null;
  return value?.protocol === 1 && value.kind === 'font-pack-request' && typeof value.id === 'string' &&
    (value.pack === 'intl' || value.pack === 'cjk');
}

export function createFontPackResponder(post: (message: FontPackReply) => void, source: OptionalFontPackSource) {
  const packs = new Map<OptionalFontPack, Promise<FontPackUrls | undefined>>();
  let disposed = false;
  return {
    receive(message: unknown): boolean {
      if (!isFontPackRequest(message)) return false;
      if (disposed) return true;
      let pending = packs.get(message.pack);
      if (!pending) {
        pending = Promise.resolve().then(() => source(message.pack));
        packs.set(message.pack, pending);
      }
      const reply = (value: FontPackReply) => { if (!disposed) { try { post(value); } catch {} } };
      pending.then(
        (urls) => reply({ protocol: 1, kind: 'font-pack-reply', id: message.id, urls }),
        (error: unknown) => reply({ protocol: 1, kind: 'font-pack-reply', id: message.id,
          error: error instanceof Error ? error.message : String(error) }),
      );
      return true;
    },
    dispose() { disposed = true; packs.clear(); },
  };
}

let nextChannel = 0;

export function installWorkerFontPacks(
  scope: Pick<Worker, 'postMessage' | 'addEventListener' | 'removeEventListener'>,
  install: (source: OptionalFontPackSource) => void,
) {
  const packs = createWorkerFontPacks((message) => scope.postMessage(message));
  const receive: EventListener = (event) => {
    if (packs.receive((event as MessageEvent).data)) event.stopImmediatePropagation();
  };
  scope.addEventListener('message', receive);
  install(packs.source);
  return () => {
    scope.removeEventListener('message', receive);
    packs.dispose();
  };
}

export function createWorkerFontPacks(post: (message: FontPackRequest) => void, waitMs = OFFICE_SESSION_SILENCE_MS) {
  const channel = ++nextChannel;
  let nextId = 0;
  let disposed = false;
  const packs = new Map<OptionalFontPack, Promise<FontPackUrls | undefined>>();
  const pending = new Map<string, { finish(urls?: FontPackUrls, error?: Error): void }>();
  const source: OptionalFontPackSource = (pack) => {
    if (disposed) return Promise.reject(new SessionFailure('disposed', 'Font pack channel is closed'));
    let promise = packs.get(pack);
    if (!promise) {
      const id = `font-pack:${channel}:${++nextId}`;
      promise = new Promise((resolve, reject) => {
        const timer = setTimeout(() => finish(), waitMs);
        const finish = (urls?: FontPackUrls, error?: Error) => {
          clearTimeout(timer);
          pending.delete(id);
          if (error) reject(error);
          else resolve(urls);
        };
        pending.set(id, { finish });
        try { post({ protocol: 1, kind: 'font-pack-request', id, pack }); }
        catch (error) { finish(undefined, new SessionFailure('message', String(error))); }
      });
      packs.set(pack, promise);
    }
    return promise;
  };
  return {
    source,
    receive(message: unknown): boolean {
      const value = message as Partial<FontPackReply> | null;
      if (value?.protocol !== 1 || value.kind !== 'font-pack-reply') return false;
      const request = typeof value.id === 'string' ? pending.get(value.id) : undefined;
      if (request) {
        const valid = value.urls === undefined || (value.urls !== null && typeof value.urls === 'object' &&
          Object.values(value.urls).every((href) => {
            try { return typeof href === 'string' && new URL(href).href === href; } catch { return false; }
          }));
        request.finish(valid ? value.urls : undefined, value.error || !valid
          ? new SessionFailure('message', value.error || 'Invalid optional font pack URLs') : undefined);
      }
      return true;
    },
    dispose() {
      disposed = true;
      for (const request of pending.values()) request.finish(undefined, new SessionFailure('disposed', 'Font pack channel is closed'));
      packs.clear();
    },
  };
}

