export * from './manifest';
export type { BundledFontSource } from './provider';
import type { BundledFontFace } from './manifest';
import { fontProvider, type BundledFontSource } from './provider';
import { loadFontBytes } from './bytes';

// Per-file LITERAL asset URLs. Bundlers only statically resolve `new URL()`
// when the specifier is a string literal — a template expression works under
// Vite's directory glob but collapses to a single (wrong) asset under
// webpack/Turbopack. Every face this package ships must have a row here; the
// CJK faces resolve through `@betteroffice/fonts-cjk` instead.
const FONT_ASSET_URLS: Record<string, () => URL> = {
  'Caladea-Bold.ttf': () =>
    new URL('../assets/Caladea-Bold.ttf', import.meta.url),
  'Caladea-BoldItalic.ttf': () =>
    new URL('../assets/Caladea-BoldItalic.ttf', import.meta.url),
  'Caladea-Italic.ttf': () =>
    new URL('../assets/Caladea-Italic.ttf', import.meta.url),
  'Caladea-Regular.ttf': () =>
    new URL('../assets/Caladea-Regular.ttf', import.meta.url),
  'Carlito-Bold.ttf': () =>
    new URL('../assets/Carlito-Bold.ttf', import.meta.url),
  'Carlito-BoldItalic.ttf': () =>
    new URL('../assets/Carlito-BoldItalic.ttf', import.meta.url),
  'Carlito-Italic.ttf': () =>
    new URL('../assets/Carlito-Italic.ttf', import.meta.url),
  'Carlito-Regular.ttf': () =>
    new URL('../assets/Carlito-Regular.ttf', import.meta.url),
  'LiberationMono-Bold.ttf': () =>
    new URL('../assets/LiberationMono-Bold.ttf', import.meta.url),
  'LiberationMono-BoldItalic.ttf': () =>
    new URL('../assets/LiberationMono-BoldItalic.ttf', import.meta.url),
  'LiberationMono-Italic.ttf': () =>
    new URL('../assets/LiberationMono-Italic.ttf', import.meta.url),
  'LiberationMono-Regular.ttf': () =>
    new URL('../assets/LiberationMono-Regular.ttf', import.meta.url),
  'LiberationSans-Bold.ttf': () =>
    new URL('../assets/LiberationSans-Bold.ttf', import.meta.url),
  'LiberationSans-BoldItalic.ttf': () =>
    new URL('../assets/LiberationSans-BoldItalic.ttf', import.meta.url),
  'LiberationSans-Italic.ttf': () =>
    new URL('../assets/LiberationSans-Italic.ttf', import.meta.url),
  'LiberationSans-Regular.ttf': () =>
    new URL('../assets/LiberationSans-Regular.ttf', import.meta.url),
  'LiberationSerif-Bold.ttf': () =>
    new URL('../assets/LiberationSerif-Bold.ttf', import.meta.url),
  'LiberationSerif-BoldItalic.ttf': () =>
    new URL('../assets/LiberationSerif-BoldItalic.ttf', import.meta.url),
  'LiberationSerif-Italic.ttf': () =>
    new URL('../assets/LiberationSerif-Italic.ttf', import.meta.url),
  'LiberationSerif-Regular.ttf': () =>
    new URL('../assets/LiberationSerif-Regular.ttf', import.meta.url),
  'NotoNaskhArabic-Regular.ttf': () =>
    new URL('../assets/NotoNaskhArabic-Regular.ttf', import.meta.url),
  'NotoSansArabic-Bold.ttf': () =>
    new URL('../assets/NotoSansArabic-Bold.ttf', import.meta.url),
  'NotoSansArabic-Regular.ttf': () =>
    new URL('../assets/NotoSansArabic-Regular.ttf', import.meta.url),
  'NotoSansHebrew-Bold.ttf': () =>
    new URL('../assets/NotoSansHebrew-Bold.ttf', import.meta.url),
  'NotoSansHebrew-Regular.ttf': () =>
    new URL('../assets/NotoSansHebrew-Regular.ttf', import.meta.url),
};

export interface FontAssetOptions {
  /**
   * Asset root. The default stays same-origin for privacy, offline use, and
   * strict CSP. Relative roots pin to the current document; server roots must
   * be absolute.
   */
  baseUrl?: string | URL;
}

let cjkAssetUrls: Promise<Record<string, () => URL> | undefined> | undefined;

async function importCjkAssetUrls(): Promise<
  Record<string, () => URL> | undefined
> {
  // Keep the SYNTACTIC try/catch with the await as its direct body. Rewriting
  // this as `import(…).catch()` or a two-argument `.then()` makes webpack (and
  // so `next build`) fail hard on the absent optional peer, and esbuild starts
  // resolving the specifier eagerly the moment it stops being that direct body.
  try {
    return (await import('@betteroffice/fonts-cjk')).CJK_FONT_ASSET_URLS;
  } catch {
    return undefined;
  }
}

/** Shares in-flight or successful CJK imports while leaving misses retryable. */
function loadCjkAssetUrls(): Promise<Record<string, () => URL> | undefined> {
  if (cjkAssetUrls === undefined) {
    const promise = importCjkAssetUrls();
    promise.then((urls) => {
      if (urls === undefined && cjkAssetUrls === promise)
        cjkAssetUrls = undefined;
    });
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

async function assetUrl(file: string, baseUrl: URL | undefined): Promise<URL> {
  if (baseUrl !== undefined) return new URL(file, baseUrl);
  const local = FONT_ASSET_URLS[file];
  if (local) return local();
  const cjk = await loadCjkAssetUrls();
  const resolveCjk = cjk?.[file];
  if (resolveCjk) return resolveCjk();
  if (!cjk) {
    throw new Error(
      `Bundled font ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`,
    );
  }
  throw new Error(`Unknown bundled font asset: ${file}`);
}

/** Loads font bytes with shared buffers and retryable failures. */
export function loadBundledFontBytes(
  face: BundledFontFace,
  options?: FontAssetOptions,
): Promise<ArrayBuffer> {
  const baseUrl =
    options?.baseUrl === undefined
      ? undefined
      : resolvedAssetBase(options.baseUrl);
  return assetUrl(face.file, baseUrl).then((url) => loadFontBytes(face, url));
}

const registeredFaces = new Map<string, Promise<void>>();

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
): Promise<void> {
  if (
    typeof document === 'undefined' ||
    typeof FontFace === 'undefined' ||
    document.fonts === undefined
  ) {
    return Promise.resolve();
  }
  const family = cssFamily ?? face.family;
  const key = `${family}|${face.weight}|${face.style}`;
  const existing = registeredFaces.get(key);
  if (existing) return existing;
  const promise = (async () => {
    const bytes = await loadBundledFontBytes(face, options);
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
    if (registeredFaces.get(key) === promise) registeredFaces.delete(key);
  });
  registeredFaces.set(key, promise);
  return promise;
}

/** Creates a lazy provider with optional custom asset URLs. */
export function createFontProvider(
  options?: FontAssetOptions,
): BundledFontSource {
  const resolvedOptions =
    options?.baseUrl === undefined
      ? undefined
      : { baseUrl: resolvedAssetBase(options.baseUrl) };
  return fontProvider((face) => loadBundledFontBytes(face, resolvedOptions));
}
