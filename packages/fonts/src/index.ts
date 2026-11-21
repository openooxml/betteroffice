import { createFontRuntime, type FontAssetUrls, type FontPackLookup } from './fontRuntime';

export { configureFontCdn, fontCdnConfiguration, type FontCdnOptions, type FontCdnConfiguration } from './fontCdn';
export * from './manifest';
export type { FontDecoder } from './decode';
export type { BundledFontLoader, BundledFontSource } from './provider';
export type { ProgressiveFontFace, SliceLocation } from './progressive';
export type { FontAssetOptions, FontPackDiagnostic } from './fontRuntime';

async function importIntlAssetUrls(): Promise<FontAssetUrls | undefined> {
  let addon: typeof import('@betteroffice/fonts-intl');
  try {
    addon = await import('@betteroffice/fonts-intl');
  } catch (error) {
    if (!missingAddon(error, 'intl')) throw error;
    return undefined;
  }
  return assetUrls(addon.INTL_FONT_ASSET_URLS, 'INTL');
}

async function importLatinAssetUrls(): Promise<FontAssetUrls | undefined> {
  let addon: typeof import('@betteroffice/fonts-latin');
  try {
    addon = await import('@betteroffice/fonts-latin');
  } catch (error) {
    if (!missingAddon(error, 'latin')) throw error;
    return undefined;
  }
  return assetUrls(addon.LATIN_FONT_ASSET_URLS, 'LATIN');
}

async function importCjkAssetUrls(): Promise<FontAssetUrls | undefined> {
  let addon: typeof import('@betteroffice/fonts-cjk');
  try {
    addon = await import('@betteroffice/fonts-cjk');
  } catch (error) {
    if (!missingAddon(error, 'cjk')) throw error;
    return undefined;
  }
  return assetUrls(addon.CJK_FONT_ASSET_URLS, 'CJK');
}

function missingAddon(error: unknown, pack: string): boolean {
  const message = (error as { message?: unknown } | null)?.message;
  return typeof message === 'string' && new RegExp(
    `^(?:Cannot find (?:module|package)|Failed to resolve module specifier|Error resolving module specifier|The specifier|Module name,) ["'“]@betteroffice/fonts-${pack}["'”]`,
  ).test(message);
}

function assetUrls(urls: FontAssetUrls, pack: string): FontAssetUrls {
  if (urls === null || typeof urls !== 'object') {
    throw new TypeError(`Invalid ${pack}_FONT_ASSET_URLS export from @betteroffice/fonts-${pack.toLowerCase()}`);
  }
  return urls;
}

const importers = { latin: importLatinAssetUrls, intl: importIntlAssetUrls, cjk: importCjkAssetUrls };
const pending = new Map<Parameters<FontPackLookup>[0], Promise<FontAssetUrls | undefined>>();
const lookup: FontPackLookup = (pack) => {
  let promise = pending.get(pack);
  if (!promise) {
    promise = importers[pack]();
    pending.set(pack, promise);
    const current = promise;
    const forget = () => { if (pending.get(pack) === current) pending.delete(pack); };
    promise.then((urls) => { if (urls === undefined) forget(); }, forget);
  }
  return promise;
};

export const {
  bundledFaceInstalled, isMissingFontAddon, loadBundledFontBytes, registerBundledFontFace,
  createFontProvider, decodeFontFaceOutsideEngine, checkFontPacks,
} = createFontRuntime(lookup);

/** The asset URLs of an installed optional font pack by file name, or undefined when the pack is not installed. */
export async function optionalFontPackUrls(pack: 'intl' | 'cjk'): Promise<Record<string, string> | undefined> {
  const urls = await lookup(pack);
  return urls && Object.fromEntries(Object.entries(urls).map(([file, url]) => [file, url().href]));
}
