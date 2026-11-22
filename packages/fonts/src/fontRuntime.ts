import type { FontDecoder } from './decode';
import { BUNDLED_FONTS, resolveMetricCompatFace, type BundledFontFace } from './manifest';
import { fontProvider, type BundledFontSource } from './provider';
import { loadFontBytes } from './bytes';
import { loadPackedFontBytes, readPackedFontBytes, isBaseLatinFont, isLatinFont, packFor, sliceDecoder, sliceRefOf } from './packed';
import type { SliceLocation } from './progressive';
import { checkInstalledPacks, type FontPackDiagnostic } from './packDiagnostics';
import { fontCdnAssetUrl, type FontCdnConfiguration } from './fontCdn';
import type { FontAssetKind } from '../../../shared/office-session/optionalFontPacks';

export interface FontAssetOptions {
  /**
   * Asset root. The default stays same-origin for privacy, offline use, and
   * strict CSP. Relative roots pin to the current document; server roots must
   * be absolute.
   */
  baseUrl?: string | URL;
}

const MISSING_ADDON = Symbol.for('@betteroffice/fonts/missing-addon');

/** @internal A face whose optional add-on package is not installed. */
class MissingFontAddonError extends Error {
  readonly [MISSING_ADDON] = true;
  constructor(message: string) {
    super(message);
    this.name = 'MissingFontAddonError';
  }
}

class MissingCjkAddonError extends MissingFontAddonError {}
class MissingIntlAddonError extends MissingFontAddonError {}

export type FontAssetUrls = Record<string, () => URL>;
type FontPack = Exclude<FontAssetKind, 'decoder'>;
type FontPackAssets = { urls?: FontAssetUrls; cdn?: FontCdnConfiguration };
export type FontPackLookup = (pack: FontPack) => Promise<FontPackAssets>;
type AssetBase = URL | Promise<URL> | undefined;

function resolveBase(baseUrl: string | URL, locationHref: string | undefined): URL {
  const href = typeof baseUrl === 'string' ? baseUrl : baseUrl.href;
  const base = href.endsWith('/') ? href : `${href}/`;
  try {
    return new URL(base, locationHref);
  } catch {
    throw new TypeError(`Font baseUrl must be absolute when no browser location exists: ${base}`);
  }
}

export function createFontRuntime(lookup: FontPackLookup, {
  inspectDisk = true, pageLocation, localCdn,
}: {
  inspectDisk?: boolean;
  pageLocation?: () => Promise<string | undefined>;
  localCdn?: () => FontCdnConfiguration;
} = {}) {
  const CJK_FILES = new Set(
    BUNDLED_FONTS.filter((face) => face.script?.startsWith('cjk-')).map((face) => face.file),
  );

  const INTL_FILES = new Set(
    BUNDLED_FONTS.filter((face) => face.file.endsWith('.ttf.zst') && !isLatinFont(face.file)).map((face) => face.file),
  );

  const cdnTimeouts = new WeakMap<URL, number>();

  function cdnAssetUrl(face: BundledFontFace, file: string, cdn: FontCdnConfiguration | undefined): URL | undefined {
    if (!cdn?.baseUrl || isBaseLatinFont(face.file)) return undefined;
    const pack = packFor(face);
    if (!pack.faces.some((entry) => entry.file === face.file)) return undefined;
    const url = fontCdnAssetUrl(pack, file, cdn.baseUrl);
    cdnTimeouts.set(url, cdn.timeoutMs);
    return url;
  }

  const requestTimeout = (url: URL) => cdnTimeouts.get(url);

  function assetBase(baseUrl: string | URL | undefined): AssetBase {
    if (baseUrl === undefined) return undefined;
    if (!pageLocation) return resolveBase(baseUrl, globalThis.location?.href);
    try {
      return resolveBase(baseUrl, undefined);
    } catch {
      const pinned = pageLocation().then((href) => resolveBase(baseUrl, href));
      void pinned.catch(() => {});
      return pinned;
    }
  }

  async function assetUrl(
    face: BundledFontFace,
    baseUrl: URL | undefined,
  ): Promise<URL> {
    const { file, script } = face;
    if (baseUrl !== undefined) return new URL(file, baseUrl);
    const pack = isBaseLatinFont(file) ? 'base' : isLatinFont(file) ? 'latin'
      : INTL_FILES.has(file) ? 'intl' : script?.startsWith('cjk-') || CJK_FILES.has(file) ? 'cjk' : undefined;
    if (pack) {
      const { urls, cdn } = await lookup(pack);
      const installed = urls?.[file];
      if (installed) return installed();
      const fallback = cdnAssetUrl(face, file, cdn);
      if (fallback) return fallback;
      if (pack === 'latin' && !urls) {
        throw new MissingFontAddonError(
          `Bundled font ${file} needs @betteroffice/fonts-latin - install @betteroffice/fonts-latin`,
        );
      }
      if (pack === 'intl' && !urls) {
        throw new MissingIntlAddonError(
          `Bundled font ${file} needs the optional intl add-on - install @betteroffice/fonts-intl`,
        );
      }
      if (pack === 'cjk') {
        throw new MissingCjkAddonError(urls
          ? `Bundled font ${file} needs a compressed @betteroffice/fonts-cjk release matching @betteroffice/fonts`
          : `Bundled font ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`);
      }
    }
    throw new Error(`Unknown bundled font asset: ${file}`);
  }

  function sliceAssetUrl(face: BundledFontFace, file: string, { urls, cdn }: FontPackAssets): URL {
    const installed = urls?.[file];
    if (installed) return installed();
    const fallback = cdnAssetUrl(face, file, cdn);
    if (fallback) return fallback;
    throw new MissingCjkAddonError(urls
      ? `Font slice ${file} needs a @betteroffice/fonts-cjk release matching @betteroffice/fonts`
      : `Font slice ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`);
  }

  function sliceLocation(
    face: BundledFontFace,
    base: AssetBase,
    runtime?: { decoder?: FontDecoder; retain?: boolean },
  ): SliceLocation | undefined {
    const ref = sliceRefOf(face.file);
    if (!ref) return undefined;
    const baseUrl = base instanceof URL ? base : undefined;
    const cdn = base === undefined ? localCdn?.() : undefined;
    const decoderUrl = (url: URL | undefined) => url && new URL('decoder/fontpack.wasm', url);
    return {
      ref,
      base: baseUrl?.href ?? (cdn?.baseUrl ? fontCdnAssetUrl(packFor(face), '', cdn.baseUrl).href : ''),
      timeoutMs: cdn?.baseUrl ? cdn.timeoutMs : undefined,
      url: async (file) => {
        const root = base instanceof Promise ? await base : base;
        if (root) return new URL(file, root);
        const assets = await lookup('cjk');
        return sliceAssetUrl(face, file, cdn ? { urls: assets.urls, cdn } : assets);
      },
      prepareUrls: baseUrl || cdn ? undefined : async () => {
        if (base instanceof Promise) {
          const root = await base;
          return { base: root.href, url: async (file: string) => new URL(file, root),
            decode: sliceDecoder(decoderUrl(root), runtime?.decoder) };
        }
        const assets = await lookup('cjk');
        return {
          base: assets.cdn?.baseUrl || '',
          url: async (file: string) => sliceAssetUrl(face, file, assets),
          timeoutMs: assets.cdn?.baseUrl ? assets.cdn.timeoutMs : undefined,
        };
      },
      requestTimeoutMs: requestTimeout,
      decode: base instanceof Promise
        ? async (encoded, maxOutput) => sliceDecoder(decoderUrl(await base), runtime?.decoder)(encoded, maxOutput)
        : sliceDecoder(decoderUrl(baseUrl), runtime?.decoder),
      decoder: runtime?.decoder,
      retain: runtime?.retain,
    };
  }

  /**
   * Whether an installed pack or the configured CDN can serve `face`.
   * @internal
   */
  async function bundledFaceInstalled(face: BundledFontFace): Promise<boolean> {
    try {
      await assetUrl(face, undefined);
      return true;
    } catch (error) {
      if (error instanceof MissingFontAddonError) return false;
      throw error;
    }
  }

  /**
   * Whether `error` is the rejection of a face whose optional pack is not installed, as opposed to an installed pack
   * failing to serve it.
   * @internal
   */
  function isMissingFontAddon(error: unknown): boolean {
    return (error as { [MISSING_ADDON]?: unknown } | null)?.[MISSING_ADDON] === true;
  }

  /** Loads font bytes into consumer-owned buffers with retryable failures. */
  function loadBundledFontBytes(
    face: BundledFontFace,
    options?: FontAssetOptions,
    runtime?: {
      decoder?: FontDecoder;
      /** @internal Entries this call creates are dropped once settled; concurrent requests still share one decode. */ retain?: boolean;
      /** @internal Reject when the CJK add-on is missing instead of loading a Latin face in its place. */ strict?: boolean;
    },
  ): Promise<ArrayBuffer> {
    return bundledFontBytes(face, loadPackedFontBytes, assetBase(options?.baseUrl), runtime);
  }

  async function bundledFontBytes(
    face: BundledFontFace,
    packed: typeof loadPackedFontBytes,
    base: AssetBase,
    runtime?: Parameters<typeof loadBundledFontBytes>[2],
  ): Promise<ArrayBuffer> {
    const baseUrl = base instanceof Promise ? await base : base;
    const bytes = isBaseLatinFont(face.file) || isLatinFont(face.file) || INTL_FILES.has(face.file) || CJK_FILES.has(face.file)
      ? packed(
        face,
        (dependency) => assetUrl(dependency, baseUrl),
        requestTimeout,
        baseUrl === undefined ? undefined : new URL('decoder/fontpack.wasm', baseUrl),
        undefined,
        runtime,
      )
      : assetUrl(face, baseUrl).then((url) => loadFontBytes(face, url, requestTimeout(url)));
    if (baseUrl !== undefined || !CJK_FILES.has(face.file) || runtime?.strict) return bytes;
    return bytes.catch((error) => {
      if (!(error instanceof MissingCjkAddonError)) throw error;
      return bundledFontBytes(
        resolveMetricCompatFace(
          face.family.includes('Serif') ? 'Times New Roman' : 'Arial',
          face.weight === 700,
          face.style === 'italic',
        )!,
        packed,
        baseUrl,
        runtime,
      );
    });
  }

  const registeredFaces = new Map<string, Promise<void>>();
  const engineRegistrations = new WeakMap<FontDecoder, Map<string, Promise<void>>>();

  /**
   * Register a face with the DOM via the `FontFace` API under an explicit CSS
   * family name (defaults to the face's real family), so browser measurement
   * uses the SAME bytes the wasm-side `FontStore` receives. Idempotent per
   * (cssFamily, weight, style); a failed registration is evicted so it can be
   * retried. Resolves as a no-op in non-DOM environments.
   */
  function registerBundledFontFace(
    face: BundledFontFace,
    cssFamily?: string,
    options?: FontAssetOptions,
    runtime?: { decoder?: FontDecoder; /** @internal Entries this call creates are dropped once settled; concurrent requests still share one decode. */ retain?: boolean },
  ): Promise<void> {
    if (
      typeof document === 'undefined' ||
      typeof FontFace === 'undefined' ||
      document.fonts === undefined
    ) {
      return Promise.resolve();
    }
    const family = cssFamily ?? face.family;
    let registrations = registeredFaces;
    if (runtime?.decoder !== undefined) {
      registrations = engineRegistrations.get(runtime.decoder) ?? new Map();
      engineRegistrations.set(runtime.decoder, registrations);
    }
    const key = `${family}|${face.weight}|${face.style}`;
    const existing = registrations.get(key);
    if (existing) return existing;
    const promise = (async () => {
      const bytes = await loadBundledFontBytes(face, options, runtime);
      // The family name goes through the FontFace API as a value, never
      // interpolated into a CSS string, so there is no CSS-injection sink here.
      const fontFace = new FontFace(family, bytes, {
        weight: String(face.weight),
        style: face.style,
      });
      await fontFace.load();
      document.fonts.add(fontFace);
    })();
    promise.catch(() => {
      if (registrations.get(key) === promise) registrations.delete(key);
    });
    registrations.set(key, promise);
    return promise;
  }

  /** Creates a lazy provider with optional custom asset URLs. */
  function createFontProvider(
    options?: FontAssetOptions,
    runtime?: { decoder?: FontDecoder; /** @internal Entries this call creates are dropped once settled; concurrent requests still share one decode. */ retain?: boolean },
  ): BundledFontSource {
    const base = assetBase(options?.baseUrl);
    return fontProvider(
      (face) => bundledFontBytes(face, loadPackedFontBytes, base, runtime),
      (face) => bundledFontBytes(face, readPackedFontBytes, base, runtime),
      (face) => sliceLocation(face, base, runtime),
    );
  }

  /**
   * Decodes a pack file in the package's own decoder and releases it afterwards, so a large face costs an engine's
   * memory nothing.
   * @internal
   */
  const decodeFontFaceOutsideEngine: FontDecoder = async (bytes, { codec, dict, maxOutput }) => {
    const decoder = await import('./decode');
    try {
      return await decoder.decodeFontFace(bytes, dict, { codec, maxOutput });
    } finally {
      decoder.releaseFontDecoder();
    }
  };

  /**
   * The installed font add-on packages that cannot be used (another release, missing files, an unreadable
   * manifest), each naming the pack and the reason. Their faces are substituted as when the pack is not installed;
   * a pack that is not installed is not reported.
   */
  function checkFontPacks(): Promise<FontPackDiagnostic[]> {
    const installed = async (pack: FontPack) => (await lookup(pack)).urls;
    return checkInstalledPacks({ latin: () => installed('latin'), intl: () => installed('intl'), cjk: () => installed('cjk') }, inspectDisk);
  }

  return { bundledFaceInstalled, isMissingFontAddon, loadBundledFontBytes, registerBundledFontFace, createFontProvider, decodeFontFaceOutsideEngine, checkFontPacks };
}

export type { FontPackDiagnostic } from './packDiagnostics';
