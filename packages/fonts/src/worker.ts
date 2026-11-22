import { getFontCdnConfiguration } from '../../../shared/office-session/fontCdn';
import { createFontRuntime, type FontPackLookup } from './fontRuntime';
import { workerFontAssets } from './workerAssets';

export { configureFontCdn, fontCdnConfiguration } from './fontCdn';
export * from './manifest';

const lookup: FontPackLookup = async (pack) => {
  const assets = await workerFontAssets(pack);
  return {
    urls: assets?.urls && Object.setPrototypeOf(
      Object.fromEntries(Object.entries(assets.urls).map(([file, href]) => [file, () => new URL(href)])),
      null,
    ),
    cdn: getFontCdnConfiguration() ?? assets?.cdn,
  };
};

export const {
  bundledFaceInstalled, isMissingFontAddon, loadBundledFontBytes, registerBundledFontFace,
  createFontProvider, decodeFontFaceOutsideEngine, checkFontPacks,
} = createFontRuntime(lookup, {
  inspectDisk: false,
  pageLocation: async () => (await workerFontAssets('base', { fresh: true }))?.pageUrl,
});
