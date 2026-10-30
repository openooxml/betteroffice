export * from './manifest';
export type { FontDecoder } from './decode';
import type { FontDecoder } from './decode';
export type { BundledFontLoader, BundledFontSource } from './provider';
import { BUNDLED_FONTS, resolveMetricCompatFace, type BundledFontFace } from './manifest';
import { fontProvider, type BundledFontSource } from './provider';
import { loadFontBytes } from './bytes';
import { loadPackedFontBytes, isLatinFont } from './packed';

export interface FontAssetOptions {
  /**
   * Asset root. The default stays same-origin for privacy, offline use, and
   * strict CSP. Relative roots pin to the current document; server roots must
   * be absolute.
   */
  baseUrl?: string | URL;
}

let cjkAssetUrls: Promise<Record<string, () => URL> | undefined> | undefined;
let intlAssetUrls: Promise<Record<string, () => URL> | undefined> | undefined;

const UNRESOLVED_INTL_ADDON =
  /^(?:Cannot find (?:module|package)|Failed to resolve module specifier|Error resolving module specifier|The specifier|Module name,) ["'“]@betteroffice\/fonts-intl["'”]/;

async function importIntlAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  let addon: typeof import('@betteroffice/fonts-intl');
  try {
    addon = await import('@betteroffice/fonts-intl');
  } catch (error) {
    const message = (error as { message?: unknown } | null)?.message;
    if (typeof message !== 'string' || !UNRESOLVED_INTL_ADDON.test(message)) {
      throw error;
    }
    return undefined;
  }
  const urls = addon.INTL_FONT_ASSET_URLS;
  if (urls === null || typeof urls !== 'object') {
    throw new TypeError('Invalid INTL_FONT_ASSET_URLS export from @betteroffice/fonts-intl');
  }
  return urls;
}

function loadIntlAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  if (intlAssetUrls === undefined) {
    const promise = importIntlAssetUrls();
    promise.then(
      (urls) => {
        if (urls === undefined && intlAssetUrls === promise)
          intlAssetUrls = undefined;
      },
      () => {
        if (intlAssetUrls === promise) intlAssetUrls = undefined;
      },
    );
    intlAssetUrls = promise;
  }
  return intlAssetUrls;
}

let latinAssetUrls: Promise<Record<string, () => URL> | undefined> | undefined;

class MissingCjkAddonError extends Error {}

const UNRESOLVED_LATIN_ADDON =
  /^(?:Cannot find (?:module|package)|Failed to resolve module specifier|Error resolving module specifier|The specifier|Module name,) ["'“]@betteroffice\/fonts-latin["'”]/;

async function importLatinAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  let addon: typeof import('@betteroffice/fonts-latin');
  try {
    addon = await import('@betteroffice/fonts-latin');
  } catch (error) {
    const message = (error as { message?: unknown } | null)?.message;
    if (typeof message !== 'string' || !UNRESOLVED_LATIN_ADDON.test(message)) {
      throw error;
    }
    return undefined;
  }
  const urls = addon.LATIN_FONT_ASSET_URLS;
  if (urls === null || typeof urls !== 'object') {
    throw new TypeError('Invalid LATIN_FONT_ASSET_URLS export from @betteroffice/fonts-latin');
  }
  return urls;
}

function loadLatinAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  if (latinAssetUrls === undefined) {
    const promise = importLatinAssetUrls();
    promise.then(
      (urls) => {
        if (urls === undefined && latinAssetUrls === promise)
          latinAssetUrls = undefined;
      },
      () => {
        if (latinAssetUrls === promise) latinAssetUrls = undefined;
      },
    );
    latinAssetUrls = promise;
  }
  return latinAssetUrls;
}

const UNRESOLVED_CJK_ADDON =
  /^(?:Cannot find (?:module|package)|Failed to resolve module specifier|Error resolving module specifier|The specifier|Module name,) ["'“]@betteroffice\/fonts-cjk["'”]/;

async function importCjkAssetUrls(): Promise<
  Record<string, () => URL> | undefined
> {
  let addon: typeof import('@betteroffice/fonts-cjk');
  // Keep the SYNTACTIC try/catch with the await as its direct body. Rewriting
  // this as `import(…).catch()` or a two-argument `.then()` makes webpack (and
  // so `next build`) fail hard on the absent optional peer, and esbuild starts
  // resolving the specifier eagerly the moment it stops being that direct body.
  try {
    addon = await import('@betteroffice/fonts-cjk');
  } catch (error) {
    const message = (error as { message?: unknown } | null)?.message;
    if (typeof message !== 'string' || !UNRESOLVED_CJK_ADDON.test(message)) {
      throw error;
    }
    return undefined;
  }
  const urls = addon.CJK_FONT_ASSET_URLS;
  if (urls === null || typeof urls !== 'object') {
    throw new TypeError('Invalid CJK_FONT_ASSET_URLS export from @betteroffice/fonts-cjk');
  }
  return urls;
}

/** Shares in-flight or successful CJK imports while leaving misses retryable. */
function loadCjkAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  if (cjkAssetUrls === undefined) {
    const promise = importCjkAssetUrls();
    promise.then(
      (urls) => {
        if (urls === undefined && cjkAssetUrls === promise)
          cjkAssetUrls = undefined;
      },
      () => {
        if (cjkAssetUrls === promise) cjkAssetUrls = undefined;
      },
    );
    cjkAssetUrls = promise;
  }
  return cjkAssetUrls;
}

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

async function assetUrl(
  { file, script }: BundledFontFace,
  baseUrl: URL | undefined,
): Promise<URL> {
  if (baseUrl !== undefined) return new URL(file, baseUrl);
  if (isLatinFont(file)) {
    const latin = await loadLatinAssetUrls();
    const resolveLatin = latin?.[file];
    if (resolveLatin) return resolveLatin();
    if (!latin) {
      throw new Error(
        `Bundled font ${file} needs @betteroffice/fonts-latin - install @betteroffice/fonts-latin`,
      );
    }
  }
  if (INTL_FILES.has(file)) {
    const intl = await loadIntlAssetUrls();
    const resolveIntl = intl?.[file];
    if (resolveIntl) return resolveIntl();
    if (!intl) {
      throw new Error(
        `Bundled font ${file} needs the optional intl add-on - install @betteroffice/fonts-intl`,
      );
    }
  }
  if (script?.startsWith('cjk-') || CJK_FILES.has(file)) {
    const cjk = await loadCjkAssetUrls();
    const resolveCjk = cjk?.[file];
    if (resolveCjk) return resolveCjk();
    if (!cjk) {
      throw new MissingCjkAddonError(
        `Bundled font ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`,
      );
    }
  }
  throw new Error(`Unknown bundled font asset: ${file}`);
}

/** Loads font bytes into consumer-owned buffers with retryable failures. */
export function loadBundledFontBytes(
  face: BundledFontFace,
  options?: FontAssetOptions,
  runtime?: { decoder?: FontDecoder },
): Promise<ArrayBuffer> {
  const baseUrl =
    options?.baseUrl === undefined
      ? undefined
      : resolvedAssetBase(options.baseUrl);
  if (isLatinFont(face.file) || INTL_FILES.has(face.file)) {
    return loadPackedFontBytes(
      face,
      (dependency) => assetUrl(dependency, baseUrl),
      undefined,
      baseUrl === undefined ? undefined : new URL('decoder/fontpack.wasm', baseUrl),
      undefined,
      runtime,
    );
  }
  const bytes = assetUrl(face, baseUrl).then((url) => loadFontBytes(face, url));
  if (baseUrl !== undefined || !CJK_FILES.has(face.file)) return bytes;
  return bytes.catch((error) => {
    if (!(error instanceof MissingCjkAddonError)) throw error;
    return loadBundledFontBytes(
      resolveMetricCompatFace(
        face.family.includes('Serif') ? 'Times New Roman' : 'Arial',
        face.weight === 700,
        face.style === 'italic',
      )!,
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
export function registerBundledFontFace(
  face: BundledFontFace,
  cssFamily?: string,
  options?: FontAssetOptions,
  runtime?: { decoder?: FontDecoder },
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
export function createFontProvider(
  options?: FontAssetOptions,
  runtime?: { decoder?: FontDecoder },
): BundledFontSource {
  const resolvedOptions =
    options?.baseUrl === undefined
      ? undefined
      : { baseUrl: resolvedAssetBase(options.baseUrl) };
  return fontProvider((face) => loadBundledFontBytes(face, resolvedOptions, runtime));
}
