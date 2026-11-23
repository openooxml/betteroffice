import { SessionFailure } from './types';
import type { FontCdnConfiguration } from './fontCdnRuntime';

export type FontAssetKind = 'base' | 'latin' | 'intl' | 'cjk' | 'decoder';
export type FontPackUrls = Record<string, string>;
export type FontPackSource = (pack: FontAssetKind) => Promise<FontPackUrls | undefined>;

type FontPackAssets = { urls?: FontPackUrls; cdn?: FontCdnConfiguration; pageUrl?: string };
export type WorkerFontPackSource = (pack: FontAssetKind, options?: { fresh?: boolean }) => Promise<FontPackAssets | undefined>;
export type FontPackRequest = { protocol: 1; kind: 'font-pack-request'; id: string; pack: FontAssetKind };
export type FontPackReply = FontPackAssets & { protocol: 1; kind: 'font-pack-reply'; id: string; error?: string };

let workerSource: WorkerFontPackSource = async () => undefined;

export function setWorkerFontAssets(source: WorkerFontPackSource): void {
  workerSource = source;
}

export const workerFontAssets: WorkerFontPackSource = (...args) => workerSource(...args);

export class WorkerFontModuleUnavailableError extends TypeError {
  constructor(specifier: string) {
    super(`Failed to resolve module specifier '${specifier}'`);
  }
}

function isAbsoluteHref(value: unknown): value is string {
  try {
    return typeof value === 'string' && new URL(value).href === value;
  } catch {
    return false;
  }
}

export function isFontPackRequest(message: unknown): message is FontPackRequest {
  const value = message as Partial<FontPackRequest> | null;
  return value?.protocol === 1 && value.kind === 'font-pack-request' && typeof value.id === 'string' &&
    (value.pack === 'base' || value.pack === 'latin' || value.pack === 'intl' || value.pack === 'cjk' || value.pack === 'decoder');
}

export function createFontPackResponder(
  post: (message: FontPackReply) => void,
  source: FontPackSource,
  fontCdn?: () => FontCdnConfiguration | undefined,
) {
  const packs = new Map<FontAssetKind, Promise<FontPackAssets>>();
  return {
    receive(message: unknown): boolean {
      if (!isFontPackRequest(message)) return false;
      const { pack } = message;
      const pageUrl = globalThis.location?.href;
      let pending = packs.get(pack);
      if (!pending) {
        pending = Promise.resolve().then(async () => {
          const assets = await source(pack);
          const urls = assets && Object.fromEntries(Object.entries(assets).map(([file, href]) => [file, new URL(href, pageUrl).href]));
          const cdn = pack === 'decoder' ? undefined : fontCdn?.();
          return cdn ? { urls, cdn } : { urls };
        });
        packs.set(pack, pending);
        const forget = () => { packs.delete(pack); };
        pending.then(forget, forget);
      }
      const reply = (value: FontPackReply) => { try { post(value); } catch {} };
      pending.then(
        (assets) => reply({ protocol: 1, kind: 'font-pack-reply', id: message.id, ...assets, ...(pack === 'base' && pageUrl ? { pageUrl } : {}) }),
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
  install: (source: WorkerFontPackSource) => void,
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

export function createWorkerFontPacks(post: (message: FontPackRequest) => void, waitMs = 15_000) {
  const channel = ++nextChannel;
  let nextId = 0;
  let disposed = false;
  let unresponsive: Set<string> | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const packs = new Map<FontAssetKind, Promise<FontPackAssets | undefined>>();
  const pending = new Map<string, { finish(assets?: FontPackAssets, error?: Error): void }>();

  function startTimer(): void {
    if (timer !== undefined) return;
    const scheduled = setTimeout(() => {
      if (timer !== scheduled) return;
      unresponsive = new Set(pending.keys());
      packs.clear();
      for (const request of pending.values()) request.finish();
    }, waitMs);
    timer = scheduled;
  }

  const source: WorkerFontPackSource = (pack, { fresh = false } = {}) => {
    if (disposed) return Promise.reject(new SessionFailure('disposed', 'Font pack channel is closed'));
    if (unresponsive) return Promise.resolve(undefined);
    let promise = fresh ? undefined : packs.get(pack);
    if (!promise) {
      const id = `font-pack:${channel}:${++nextId}`;
      promise = new Promise((resolve, reject) => {
        const finish = (assets?: FontPackAssets, error?: Error) => {
          pending.delete(id);
          if (pending.size === 0) {
            clearTimeout(timer);
            timer = undefined;
          }
          if (error) reject(error);
          else resolve(assets);
        };
        pending.set(id, { finish });
        startTimer();
        try { post({ protocol: 1, kind: 'font-pack-request', id, pack }); }
        catch (error) { finish(undefined, new SessionFailure('message', String(error))); }
      });
      if (!fresh) {
        packs.set(pack, promise);
        const forget = () => { if (packs.get(pack) === promise) packs.delete(pack); };
        promise.then(forget, forget);
      }
    }
    return promise;
  };
  return {
    source,
    receive(message: unknown): boolean {
      const value = message as Partial<FontPackReply> | null;
      if (value?.protocol !== 1 || value.kind !== 'font-pack-reply') return false;
      if (typeof value.id !== 'string') return true;
      if (unresponsive?.has(value.id)) unresponsive = undefined;
      const request = pending.get(value.id);
      if (request) {
        const valid = value.urls === undefined || (value.urls !== null && typeof value.urls === 'object' && !Array.isArray(value.urls) &&
          Object.values(value.urls).every(isAbsoluteHref));
        if (!valid || value.error !== undefined) {
          request.finish(undefined, new SessionFailure('message', value.error ?? 'Invalid font asset URLs'));
        } else {
          const cdn = value.cdn;
          const validCdn = cdn && Number.isInteger(cdn.timeoutMs) && cdn.timeoutMs > 0 && cdn.timeoutMs <= 4_294_967_295 &&
            isAbsoluteHref(cdn.baseUrl) && /^https?:\/\//.test(cdn.baseUrl);
          const pageUrl = isAbsoluteHref(value.pageUrl) ? value.pageUrl : undefined;
          request.finish({ urls: value.urls, ...(validCdn ? { cdn } : {}), ...(pageUrl ? { pageUrl } : {}) });
        }
      }
      return true;
    },
    dispose() {
      disposed = true;
      unresponsive = undefined;
      for (const request of pending.values()) request.finish(undefined, new SessionFailure('disposed', 'Font pack channel is closed'));
      packs.clear();
    },
  };
}

