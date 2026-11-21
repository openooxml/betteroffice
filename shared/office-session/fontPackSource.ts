import { createFontPackResponder, isFontPackRequest, type FontPackUrls, type OptionalFontPack, type OptionalFontPackSource } from './optionalFontPacks';
import type { SessionTransport } from './transport';

function missingAddon(error: unknown, pack: OptionalFontPack): boolean {
  const message = (error as { message?: unknown } | null)?.message;
  return typeof message === 'string' && new RegExp(
    `^(?:Cannot find (?:module|package)|Failed to resolve module specifier|Error resolving module specifier|The specifier|Module name,) ["'“]@betteroffice/fonts-${pack}["'”]`,
  ).test(message);
}

export const importOptionalFontPack: OptionalFontPackSource = async (pack) => {
  let urls: Record<string, () => URL>;
  if (pack === 'intl') {
    let addon: typeof import('@betteroffice/fonts-intl');
    try {
      addon = await import('@betteroffice/fonts-intl');
    } catch (error) {
      if (!missingAddon(error, pack)) throw error;
      return undefined;
    }
    urls = addon.INTL_FONT_ASSET_URLS;
  } else {
    let addon: typeof import('@betteroffice/fonts-cjk');
    try {
      addon = await import('@betteroffice/fonts-cjk');
    } catch (error) {
      if (!missingAddon(error, pack)) throw error;
      return undefined;
    }
    urls = addon.CJK_FONT_ASSET_URLS;
  }
  if (urls === null || typeof urls !== 'object') throw new TypeError(`Invalid @betteroffice/fonts-${pack} asset URLs`);
  const result: FontPackUrls = {};
  for (const [file, url] of Object.entries(urls)) result[file] = url().href;
  return result;
};

export function withFontPackSource(transport: SessionTransport, source = importOptionalFontPack): SessionTransport {
  const responder = createFontPackResponder((message) => transport.post(message), source);
  const unlisten = transport.listen((message) => { responder.receive(message); });
  return {
    ...transport,
    listen: (listener) => transport.listen((message) => { if (!isFontPackRequest(message)) listener(message); }),
    close() { responder.dispose(); unlisten(); transport.close(); },
  };
}
