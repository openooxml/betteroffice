import { SessionFailure } from './types';

export type FontAssetKind = 'latin' | 'intl' | 'cjk' | 'decoder';
export type FontPackUrls = Record<string, string>;
export type FontPackSource = (pack: FontAssetKind) => Promise<FontPackUrls | undefined>;

type FontPackCdn = { prefix: string; timeoutMs: number };
export type FontPackRequest = { protocol: 1; kind: 'font-pack-request'; id: string; pack: FontAssetKind };
export type FontPackReply = { protocol: 1; kind: 'font-pack-reply'; id: string; urls?: FontPackUrls; error?: string; cdn?: FontPackCdn };

export function isFontPackRequest(message: unknown): message is FontPackRequest {
  const value = message as Partial<FontPackRequest> | null;
  return value?.protocol === 1 && value.kind === 'font-pack-request' && typeof value.id === 'string' &&
    (value.pack === 'latin' || value.pack === 'intl' || value.pack === 'cjk' || value.pack === 'decoder');
}

export function createFontPackResponder(
  post: (message: FontPackReply) => void,
  source: FontPackSource,
  fontCdn?: (pack: Exclude<FontAssetKind, 'decoder'>) => FontPackCdn | undefined,
) {
  const packs = new Map<FontAssetKind, Promise<Pick<FontPackReply, 'urls' | 'cdn'>>>();
  return {
    receive(message: unknown): boolean {
      if (!isFontPackRequest(message)) return false;
      const { pack } = message;
      let pending = packs.get(pack);
      if (!pending) {
        pending = Promise.resolve().then(async () => {
          const urls = await source(pack);
          const cdn = urls && pack !== 'decoder' ? fontCdn?.(pack) : undefined;
          return cdn ? { urls, cdn } : { urls };
        });
        packs.set(pack, pending);
        const forget = () => { packs.delete(pack); };
        pending.then(({ urls }) => { if (urls === undefined) forget(); }, forget);
      }
      const reply = (value: FontPackReply) => { try { post(value); } catch {} };
      pending.then(
        (assets) => reply({ protocol: 1, kind: 'font-pack-reply', id: message.id, ...assets }),
        (error: unknown) => reply({ protocol: 1, kind: 'font-pack-reply', id: message.id,
          error: error instanceof Error ? error.message : String(error) }),
      );
      return true;
    },
  };
}

let nextChannel = 0;

export function installWorkerFontPacks(
  scope: Pick<Worker, 'postMessage' | 'addEventListener' | 'removeEventListener'>,
  install: (source: FontPackSource, timeout: (url: URL) => number | undefined) => void,
) {
  const packs = createWorkerFontPacks((message) => scope.postMessage(message));
  const receive: EventListener = (event) => {
    if (packs.receive((event as MessageEvent).data)) event.stopImmediatePropagation();
  };
  scope.addEventListener('message', receive);
  install(packs.source, packs.timeout);
  return () => {
    scope.removeEventListener('message', receive);
    packs.dispose();
  };
}

export function createWorkerFontPacks(post: (message: FontPackRequest) => void, waitMs = 15_000) {
  const channel = ++nextChannel;
  let nextId = 0;
  let disposed = false;
  const packs = new Map<FontAssetKind, Promise<FontPackUrls | undefined>>();
  const pending = new Map<string, { finish(urls?: FontPackUrls, error?: Error): void }>();
  const timeouts = new Map<string, number>();
  const source: FontPackSource = (pack) => {
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
      const forget = () => { packs.delete(pack); };
      promise.then((urls) => { if (urls === undefined) forget(); }, forget);
    }
    return promise;
  };
  return {
    source,
    timeout: (url: URL) => timeouts.get(url.href),
    receive(message: unknown): boolean {
      const value = message as Partial<FontPackReply> | null;
      if (value?.protocol !== 1 || value.kind !== 'font-pack-reply') return false;
      const request = typeof value.id === 'string' ? pending.get(value.id) : undefined;
      if (request) {
        const valid = value.urls === undefined || (value.urls !== null && typeof value.urls === 'object' && !Array.isArray(value.urls) &&
          Object.values(value.urls).every((href) => {
            try { return typeof href === 'string' && new URL(href).href === href; } catch { return false; }
          }));
        if (!valid || value.error !== undefined) {
          request.finish(undefined, new SessionFailure('message', value.error ?? 'Invalid font asset URLs'));
        } else {
          const { cdn, urls } = value;
          if (urls && cdn && typeof cdn.prefix === 'string' && cdn.prefix && Number.isInteger(cdn.timeoutMs) &&
            cdn.timeoutMs > 0 && cdn.timeoutMs <= 4_294_967_295) {
            for (const href of Object.values(urls)) if (href.startsWith(cdn.prefix)) timeouts.set(href, cdn.timeoutMs);
          }
          request.finish(urls);
        }
      }
      return true;
    },
    dispose() {
      disposed = true;
      for (const request of pending.values()) request.finish(undefined, new SessionFailure('disposed', 'Font pack channel is closed'));
      packs.clear();
      timeouts.clear();
    },
  };
}

