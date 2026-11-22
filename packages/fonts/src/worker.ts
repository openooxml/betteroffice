import type { FontPackSource } from '../../../shared/office-session/optionalFontPacks';
import { createFontRuntime, type FontAssetUrls, type FontPackLookup } from './fontRuntime';
import { setWorkerFontAssets, workerFontAssets, workerFontTimeout } from './workerAssets';

export { setFontCdnConfiguration as configureFontCdn } from '../../../shared/office-session/fontCdn';
export * from './manifest';
export * from './progressive-entry';

const requested = new Map<Parameters<FontPackLookup>[0], Promise<FontAssetUrls | undefined>>();

export function setFontAssetSource(source: FontPackSource, timeout?: (url: URL) => number | undefined): void {
  setWorkerFontAssets(source, timeout);
  requested.clear();
}

const lookup: FontPackLookup = (pack) => {
  let pending = requested.get(pack);
  if (!pending) {
    pending = Promise.resolve().then(() => workerFontAssets(pack)).then((urls) => urls &&
      Object.fromEntries(Object.entries(urls).map(([file, href]) => [file, () => new URL(href)])));
    requested.set(pack, pending);
    const current = pending;
    const forget = () => { if (requested.get(pack) === current) requested.delete(pack); };
    pending.then((urls) => { if (urls === undefined) forget(); }, forget);
  }
  return pending;
};

export const {
  bundledFaceInstalled, isMissingFontAddon, loadBundledFontBytes, registerBundledFontFace,
  createFontProvider, decodeFontFaceOutsideEngine, checkFontPacks,
} = createFontRuntime(lookup, {
  diagnostics: (pack) => requested.get(pack) ?? Promise.resolve(undefined),
  timeout: workerFontTimeout,
});
