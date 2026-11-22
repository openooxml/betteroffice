import { createFontRuntime, type FontAssetUrls, type FontPackLookup } from './fontRuntime';
import { absoluteHref } from './assetHref';
import { bundledDecoderUrl } from './decodeAssets';
import { fontCdnConfiguration } from './fontCdn';
import { BASE_LATIN_FONT_ASSET_URLS } from './base-latin';
import type { FontAssetKind } from '../../../shared/office-session/optionalFontPacks';

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
const pending = new Map<keyof typeof importers, Promise<FontAssetUrls | undefined>>();
function installedUrls(pack: Exclude<FontAssetKind, 'decoder'>): Promise<FontAssetUrls | undefined> {
  if (pack === 'base') return Promise.resolve(BASE_LATIN_FONT_ASSET_URLS);
  let promise = pending.get(pack);
  if (!promise) {
    promise = importers[pack]();
    pending.set(pack, promise);
    const current = promise;
    const forget = () => { if (pending.get(pack) === current) pending.delete(pack); };
    promise.then((urls) => { if (urls === undefined) forget(); }, forget);
  }
  return promise;
}

const lookup: FontPackLookup = async (pack) => ({ urls: await installedUrls(pack), cdn: fontCdnConfiguration() });

export const {
  bundledFaceInstalled, isMissingFontAddon, loadBundledFontBytes, registerBundledFontFace,
  createFontProvider, decodeFontFaceOutsideEngine, checkFontPacks,
} = createFontRuntime(lookup, { localCdn: fontCdnConfiguration });

/** Absolute installed font asset URLs by file name, or undefined when the package is absent. */
export async function fontAssetUrls(kind: FontAssetKind): Promise<Record<string, string> | undefined> {
  if (kind === 'decoder') {
    const url = bundledDecoderUrl();
    return url && { 'fontpack.wasm': url.href };
  }
  const installed = await installedUrls(kind);
  return installed && Object.fromEntries(Object.entries(installed).map(([file, url]) => [file, new URL(absoluteHref(url().href)).href]));
}
