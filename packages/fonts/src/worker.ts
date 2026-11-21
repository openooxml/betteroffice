import * as latinAddon from '@betteroffice/fonts-latin';
import { createFontRuntime, type FontAssetUrls } from './fontRuntime';

export { configureFontCdn, fontCdnConfiguration, type FontCdnOptions, type FontCdnConfiguration } from './fontCdn';
export * from './manifest';
export * from './progressive-entry';
export type { FontDecoder } from './decode';
export type { BundledFontLoader, BundledFontSource } from './provider';
export type { FontAssetOptions, FontPackDiagnostic } from './fontRuntime';

export type OptionalFontPack = 'intl' | 'cjk';
export type OptionalFontPackSource = (pack: OptionalFontPack) => Promise<Record<string, string> | undefined>;

let optionalFontPacks: OptionalFontPackSource = async () => undefined;
const requested = new Map<OptionalFontPack, Promise<FontAssetUrls | undefined>>();

export function setOptionalFontPacks(source: OptionalFontPackSource): void {
  optionalFontPacks = source;
  requested.clear();
}

const latin = () => Promise.resolve(latinAddon.LATIN_FONT_ASSET_URLS);
const lookup = (pack: 'latin' | OptionalFontPack): Promise<FontAssetUrls | undefined> => {
  if (pack === 'latin') return latin();
  let pending = requested.get(pack);
  if (!pending) {
    pending = Promise.resolve().then(() => optionalFontPacks(pack)).then((urls) => urls &&
      Object.fromEntries(Object.entries(urls).map(([file, href]) => [file, () => new URL(href)])));
    requested.set(pack, pending);
  }
  return pending;
};

export const {
  bundledFaceInstalled, isMissingFontAddon, loadBundledFontBytes, registerBundledFontFace,
  createFontProvider, decodeFontFaceOutsideEngine, checkFontPacks,
} = createFontRuntime(lookup, (pack) => pack === 'latin' ? latin() : requested.get(pack) ?? Promise.resolve(undefined));
