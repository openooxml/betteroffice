import type { FontDecoder } from './decode';
import { BASE_LATIN_FONTS, BUNDLED_FONTS, resolveMetricCompatFace, type BundledFontFace } from './manifest';
import { fontProvider, type BundledFontSource } from './provider';
import { loadFontBytes } from './bytes';
import { loadPackedFontBytes, readPackedFontBytes, isBaseLatinFont, isLatinFont, packFor, sliceDecoder, sliceRefOf } from './packed';
import type { SliceLocation } from './progressive';
import { checkInstalledPacks } from './packDiagnostics';
import { isImplicitFontCdnError, reportImplicitCdnFailure } from '../../../shared/fontCdnDiagnostics';
import type { FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';
import {
  checkImplicitSliceCooldown, fontCdnAssetUrl, fontCdnCacheKey, implicitAssetCooldown, implicitFontRequest,
  type FontCdnConfiguration, type FontRequestPolicy,
} from './fontCdn';
import type { FontAssetKind } from '../../../shared/office-session/optionalFontPacks';

export interface FontAssetOptions {
  /** Relative asset roots resolve against the document. Server roots must be absolute. */
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
type FontLoadOptions = { decoder?: FontDecoder; retain?: boolean; strict?: boolean | 'except-optional-cdn-failure' };
type RuntimeFontSource = BundledFontSource & {
  /** @internal */
  faceUnavailableUntil?(face: BundledFontFace): Promise<number | undefined>;
};

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

  const cdnRequests = new WeakMap<URL, FontRequestPolicy>();
  const substitutedCdnFaces = new Set<string>();

  function cdnAssetUrl(
    face: BundledFontFace, file: string, kind: FontPack, cdn: FontCdnConfiguration | undefined,
  ): URL | undefined {
    if (!cdn?.baseUrl || kind === 'base') return undefined;
    const pack = packFor(face);
    if (!pack.faces.some((entry) => entry.file === face.file)) return undefined;
    const url = fontCdnAssetUrl(pack, file, cdn.implicit ? undefined : cdn.baseUrl);
    cdnRequests.set(url, cdn.implicit ? implicitFontRequest(kind) : cdn.timeoutMs);
    return url;
  }

  function requestPolicy(url: URL): FontRequestPolicy | undefined {
    return cdnRequests.get(url);
  }

  function requestTimeout(url: URL): number | undefined {
    const request = requestPolicy(url);
    return typeof request === 'number' ? request : undefined;
  }

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
    skipImplicit = false,
  ): Promise<URL> {
    const { file, script } = face;
    if (baseUrl !== undefined) return new URL(file, baseUrl);
    const pack = isBaseLatinFont(file) ? 'base' : isLatinFont(file) ? 'latin'
      : INTL_FILES.has(file) ? 'intl' : script?.startsWith('cjk-') || CJK_FILES.has(file) ? 'cjk' : undefined;
    if (pack) {
      const { urls, cdn } = await lookup(pack);
      const installed = urls?.[file];
      if (installed) return installed();
      const fallback = cdnAssetUrl(face, file, pack, skipImplicit && cdn?.implicit ? undefined : cdn);
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
    const fallback = cdnAssetUrl(face, file, 'cjk', cdn);
    if (fallback) return fallback;
    throw new MissingCjkAddonError(urls
      ? `Font slice ${file} needs a @betteroffice/fonts-cjk release matching @betteroffice/fonts`
      : `Font slice ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`);
  }

  function sliceLocation(
    face: BundledFontFace,
    base: AssetBase,
    runtime?: FontLoadOptions,
  ): SliceLocation | undefined {
    const ref = sliceRefOf(face.file);
    if (!ref) return undefined;
    const baseUrl = base instanceof URL ? base : undefined;
    const decoderUrl = (url: URL | undefined) => url && new URL('decoder/fontpack.wasm', url);
    return {
      ref,
      base: baseUrl?.href ?? '',
      url: async (file) => {
        const root = base instanceof Promise ? await base : base;
        if (root) return new URL(file, root);
        return sliceAssetUrl(face, file, await lookup('cjk'));
      },
      prepareUrls: baseUrl ? undefined : async () => {
        if (base instanceof Promise) {
          const root = await base;
          return { base: root.href, url: async (file: string) => new URL(file, root),
            decode: sliceDecoder(decoderUrl(root), runtime?.decoder) };
        }
        const assets = await lookup('cjk');
        const current = localCdn ? { urls: assets.urls, cdn: localCdn() } : assets;
        const indexUrl = sliceAssetUrl(face, ref.index, current);
        return {
          base: typeof requestPolicy(indexUrl) === 'object' ? `${indexUrl.href}#implicit` : indexUrl.href,
          url: Object.assign(
            async (file: string) => file === ref.index ? indexUrl : sliceAssetUrl(face, file, current),
            { implicitCdn: (file: string) => current.cdn?.implicit === true && !current.urls?.[file] },
          ),
          timeoutMs: requestTimeout(indexUrl),
        };
      },
      loadWholeForExport: () => bundledFontBytes(face, loadPackedFontBytes, base, { ...runtime, strict: true }),
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
      const url = await assetUrl(face, undefined);
      return typeof requestPolicy(url) !== 'object' || !substitutedCdnFaces.has(url.href);
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

  async function faceUnavailableUntil(face: BundledFontFace, strict: boolean): Promise<number | undefined> {
    const cooldown = async (dependency: BundledFontFace) => {
      const url = await assetUrl(dependency, undefined);
      if (typeof requestPolicy(url) !== 'object') return undefined;
      const slices = strict && sliceRefOf(dependency.file);
      return implicitAssetCooldown(url, slices ? new URL(slices.index, url) : undefined);
    };
    try {
      const dictionary = packFor(face).faces.find(({ file }) => file === face.file)?.dict;
      const faceUntil = await cooldown(face);
      const dictionaryUntil = dictionary ? await cooldown({ ...face, file: dictionary }) : undefined;
      return Math.max(faceUntil ?? 0, dictionaryUntil ?? 0) || undefined;
    } catch (error) {
      if (isMissingFontAddon(error)) return Infinity;
      throw error;
    }
  }

  /** Loads font bytes into consumer-owned buffers with retryable failures. */
  function loadBundledFontBytes(
    face: BundledFontFace,
    options?: FontAssetOptions,
    runtime?: FontLoadOptions,
  ): Promise<ArrayBuffer> {
    return bundledFontBytes(face, loadPackedFontBytes, assetBase(options?.baseUrl), runtime);
  }

  async function bundledFontBytes(
    face: BundledFontFace,
    packed: typeof loadPackedFontBytes,
    base: AssetBase,
    runtime?: FontLoadOptions,
    skipImplicit = false,
    onCdnFailure?: () => void,
  ): Promise<ArrayBuffer> {
    const baseUrl = base instanceof Promise ? await base : base;
    const substituteOptionalCdnFailure = runtime?.strict === 'except-optional-cdn-failure';
    let implicitCdnFace: string | undefined;
    const resolve = async (dependency: BundledFontFace) => {
      const url = await assetUrl(dependency, baseUrl, skipImplicit);
      const implicit = typeof requestPolicy(url) === 'object';
      if (substituteOptionalCdnFailure && implicit && dependency === face) implicitCdnFace = url.href;
      const slices = runtime?.strict && implicit && sliceRefOf(dependency.file);
      if (slices) checkImplicitSliceCooldown(new URL(slices.index, url));
      return url;
    };
    try {
      const bytes = await (isBaseLatinFont(face.file) || isLatinFont(face.file) || INTL_FILES.has(face.file) || CJK_FILES.has(face.file)
        ? packed(
          face, resolve, requestPolicy,
          baseUrl === undefined ? undefined : new URL('decoder/fontpack.wasm', baseUrl),
          undefined, runtime,
        )
        : resolve(face).then((url) => loadFontBytes(face, url, requestTimeout(url))));
      if (implicitCdnFace) substitutedCdnFaces.delete(implicitCdnFace);
      return bytes;
    } catch (error) {
      if (baseUrl !== undefined) throw error;
      if (isImplicitFontCdnError(error)) {
        const optional = error.diagnostic.pack !== 'latin';
        if (runtime?.strict && !(substituteOptionalCdnFailure && optional)) throw error;
        reportImplicitCdnFailure(error);
        onCdnFailure?.();
        const bytes = await substituteFontBytes(face, packed, runtime, optional);
        if (implicitCdnFace && optional) substitutedCdnFaces.add(implicitCdnFace);
        return bytes;
      }
      if (runtime?.strict) throw error;
      if (error instanceof MissingCjkAddonError || error instanceof MissingIntlAddonError) {
        return substituteFontBytes(face, packed, runtime, true);
      }
      throw error;
    }
  }

  async function substituteFontBytes(
    face: BundledFontFace,
    packed: typeof loadPackedFontBytes,
    runtime: FontLoadOptions | undefined,
    script: boolean,
  ): Promise<ArrayBuffer> {
    const heavy = face.weight >= 600;
    if (script) {
      const substitute = resolveMetricCompatFace(
        face.family.includes('Serif') ? 'Times New Roman' : 'Arial', heavy, face.style === 'italic',
      )!;
      try { return await bundledFontBytes(substitute, packed, undefined, runtime, true); }
      catch (error) { if (!(error instanceof MissingFontAddonError)) throw error; }
    }
    const lastResort = BASE_LATIN_FONTS.find(({ weight, style }) => (weight === 700) === heavy && style === face.style)!;
    return bundledFontBytes(lastResort, packed, undefined, runtime);
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
    const key = `${family}|${face.weight}|${face.style}|${fontCdnCacheKey()}`;
    const existing = registrations.get(key);
    if (existing) return existing;
    let substituted = false;
    const forget = () => {
      if (registrations.get(key) === promise) registrations.delete(key);
    };
    const promise = (async () => {
      const bytes = await bundledFontBytes(face, loadPackedFontBytes, assetBase(options?.baseUrl), runtime, false,
        () => { substituted = true; });
      // The family name goes through the FontFace API as a value, never
      // interpolated into a CSS string, so there is no CSS-injection sink here.
      const fontFace = new FontFace(family, bytes, {
        weight: String(face.weight),
        style: face.style,
      });
      await fontFace.load();
      document.fonts.add(fontFace);
    })();
    void promise.then(() => { if (substituted) forget(); }, forget);
    registrations.set(key, promise);
    return promise;
  }

  /** Creates a lazy provider with optional custom asset URLs. */
  function createFontProvider(
    options?: FontAssetOptions,
    runtime?: FontLoadOptions,
  ): RuntimeFontSource {
    const base = assetBase(options?.baseUrl);
    const source: RuntimeFontSource = fontProvider(
      (face) => bundledFontBytes(face, loadPackedFontBytes, base, runtime),
      (face) => bundledFontBytes(face, readPackedFontBytes, base, runtime),
      (face) => sliceLocation(face, base, runtime),
    );
    if (base === undefined) source.faceUnavailableUntil = (face) => faceUnavailableUntil(face, Boolean(runtime?.strict));
    return source;
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
  async function checkFontPacks(): Promise<FontPackDiagnostic[]> {
    const installed = async (pack: FontPack) => (await lookup(pack)).urls;
    return checkInstalledPacks({ latin: () => installed('latin'), intl: () => installed('intl'), cjk: () => installed('cjk') }, inspectDisk);
  }

  return { bundledFaceInstalled, isMissingFontAddon, loadBundledFontBytes, registerBundledFontFace, createFontProvider, decodeFontFaceOutsideEngine, checkFontPacks };
}

export type { FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';
