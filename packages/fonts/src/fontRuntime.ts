import { fontCdnConfiguration, fontCdnAssetUrl } from './fontCdn';
import type { FontDecoder } from './decode';
import { BUNDLED_FONTS, resolveMetricCompatFace, type BundledFontFace } from './manifest';
import { fontProvider, type BundledFontSource } from './provider';
import { loadFontBytes } from './bytes';
import { loadPackedFontBytes, readPackedFontBytes, isBaseLatinFont, isLatinFont, sliceDecoder, sliceRefOf, packFor } from './packed';
import { BASE_LATIN_FONT_ASSET_URLS } from './base-latin';
import type { SliceLocation } from './progressive';
import { checkInstalledPacks, type FontPackDiagnostic } from './packDiagnostics';

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

export type FontAssetUrls = Record<string, () => URL>;
export type FontPackLookup = (pack: 'latin' | 'intl' | 'cjk') => Promise<FontAssetUrls | undefined>;

export function createFontRuntime(lookup: FontPackLookup, diagnostics: FontPackLookup = lookup) {
  const loadLatinAssetUrls = () => lookup('latin');
  const loadIntlAssetUrls = () => lookup('intl');
  const loadCjkAssetUrls = () => lookup('cjk');

  function assetBase(baseUrl: string | URL): string {
    const href = typeof baseUrl === 'string' ? baseUrl : baseUrl.href;
    return href.endsWith('/') ? href : `${href}/`;
  }

  function resolvedAssetBase(baseUrl: string | URL): URL {
    const base = assetBase(baseUrl);
    try {
      return typeof location === 'undefined'
        ? new URL(base)
        : new URL(base, location.href);
    } catch {
      throw new TypeError(
        `Font baseUrl must be absolute when no browser location exists: ${base}`,
      );
    }
  }

  const CJK_FILES = new Set(
    BUNDLED_FONTS.filter((face) => face.script?.startsWith('cjk-')).map((face) => face.file),
  );

  const INTL_FILES = new Set(
    BUNDLED_FONTS.filter((face) => face.file.endsWith('.ttf.zst') && !isLatinFont(face.file)).map((face) => face.file),
  );

  const cdnTimeouts = new WeakMap<URL, number>();

  function cdnAssetUrl(face: BundledFontFace, file = face.file, configuration = fontCdnConfiguration()): URL | undefined {
    const { baseUrl, timeoutMs } = configuration;
    if (baseUrl === false || isBaseLatinFont(face.file)) return undefined;
    const pack = packFor(face);
    if (!pack.faces.some((entry) => entry.file === face.file)) return undefined;
    const url = fontCdnAssetUrl(pack, file, baseUrl);
    cdnTimeouts.set(url, timeoutMs);
    return url;
  }

  const cdnTimeout = (url: URL) => cdnTimeouts.get(url);

  async function assetUrl(
    face: BundledFontFace,
    baseUrl: URL | undefined,
  ): Promise<URL> {
    const { file, script } = face;
    if (baseUrl !== undefined) return new URL(file, baseUrl);
    const baseLatin = BASE_LATIN_FONT_ASSET_URLS[file];
    if (baseLatin) return baseLatin();
    if (isLatinFont(file)) {
      const latin = await loadLatinAssetUrls();
      const resolveLatin = latin?.[file];
      if (resolveLatin) return resolveLatin();
      const cdn = cdnAssetUrl(face);
      if (cdn) return cdn;
      if (!latin) {
        throw new MissingFontAddonError(
          `Bundled font ${file} needs @betteroffice/fonts-latin - install @betteroffice/fonts-latin`,
        );
      }
    }
    if (INTL_FILES.has(file)) {
      const intl = await loadIntlAssetUrls();
      const resolveIntl = intl?.[file];
      if (resolveIntl) return resolveIntl();
      const cdn = cdnAssetUrl(face);
      if (cdn) return cdn;
      if (!intl) {
        throw new MissingFontAddonError(
          `Bundled font ${file} needs the optional intl add-on - install @betteroffice/fonts-intl`,
        );
      }
    }
    if (script?.startsWith('cjk-') || CJK_FILES.has(file)) {
      const cjk = await loadCjkAssetUrls();
      const resolveCjk = cjk?.[file];
      if (resolveCjk) return resolveCjk();
      const cdn = cdnAssetUrl(face);
      if (cdn) return cdn;
      throw new MissingCjkAddonError(cjk
        ? `Bundled font ${file} needs a compressed @betteroffice/fonts-cjk release matching @betteroffice/fonts`
        : `Bundled font ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`);
    }
    throw new Error(`Unknown bundled font asset: ${file}`);
  }

  async function sliceAssetUrl(
    face: BundledFontFace, file: string, baseUrl: URL | undefined, configuration = fontCdnConfiguration(),
  ): Promise<URL> {
    if (baseUrl !== undefined) return new URL(file, baseUrl);
    const cjk = await loadCjkAssetUrls();
    const resolve = cjk?.[file];
    if (resolve) return resolve();
    const cdn = cdnAssetUrl(face, file, configuration);
    if (cdn) return cdn;
    throw new MissingCjkAddonError(cjk
      ? `Font slice ${file} needs a @betteroffice/fonts-cjk release matching @betteroffice/fonts`
      : `Font slice ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`);
  }

  function sliceLocation(
    face: BundledFontFace,
    baseUrl: URL | undefined,
    runtime?: { decoder?: FontDecoder; retain?: boolean },
  ): SliceLocation | undefined {
    const ref = sliceRefOf(face.file);
    const cdn = fontCdnConfiguration();
    return ref && {
      ref,
      base: baseUrl?.href ?? (cdn.baseUrl ? fontCdnAssetUrl(packFor(face), '', cdn.baseUrl).href : ''),
      url: (file) => sliceAssetUrl(face, file, baseUrl, cdn),
      timeoutMs: baseUrl === undefined && cdn.baseUrl ? cdn.timeoutMs : undefined,
      requestTimeoutMs: cdnTimeout,
      decode: sliceDecoder(baseUrl === undefined ? undefined : new URL('decoder/fontpack.wasm', baseUrl), runtime?.decoder),
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
    return bundledFontBytes(face, loadPackedFontBytes, options, runtime);
  }

  function bundledFontBytes(
    face: BundledFontFace,
    packed: typeof loadPackedFontBytes,
    options?: FontAssetOptions,
    runtime?: Parameters<typeof loadBundledFontBytes>[2],
  ): Promise<ArrayBuffer> {
    const baseUrl =
      options?.baseUrl === undefined
        ? undefined
        : resolvedAssetBase(options.baseUrl);
    const bytes = isBaseLatinFont(face.file) || isLatinFont(face.file) || INTL_FILES.has(face.file) || CJK_FILES.has(face.file)
      ? packed(
        face,
        (dependency) => assetUrl(dependency, baseUrl),
        cdnTimeout,
        baseUrl === undefined ? undefined : new URL('decoder/fontpack.wasm', baseUrl),
        undefined,
        runtime,
      )
      : assetUrl(face, baseUrl).then((url) => loadFontBytes(face, url, cdnTimeout(url)));
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
        options,
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
    const resolvedOptions =
      options?.baseUrl === undefined
        ? undefined
        : { baseUrl: resolvedAssetBase(options.baseUrl) };
    return fontProvider(
      (face) => loadBundledFontBytes(face, resolvedOptions, runtime),
      (face) => bundledFontBytes(face, readPackedFontBytes, resolvedOptions, runtime),
      (face) => sliceLocation(face, resolvedOptions?.baseUrl, runtime),
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
    return checkInstalledPacks({ latin: () => diagnostics('latin'), intl: () => diagnostics('intl'), cjk: () => diagnostics('cjk') });
  }

  return { bundledFaceInstalled, isMissingFontAddon, loadBundledFontBytes, registerBundledFontFace, createFontProvider, decodeFontFaceOutsideEngine, checkFontPacks };
}

export type { FontPackDiagnostic } from './packDiagnostics';
