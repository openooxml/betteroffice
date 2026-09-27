export * from './manifest';
export type { BundledFontSource } from './provider';
import { BUNDLED_FONTS, type BundledFontFace } from './manifest';
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
  'ComicRelief-Bold.ttf': () =>
    new URL('../assets/ComicRelief-Bold.ttf', import.meta.url),
  'ComicRelief-Regular.ttf': () =>
    new URL('../assets/ComicRelief-Regular.ttf', import.meta.url),
  'DMSans-Bold.ttf': () =>
    new URL('../assets/DMSans-Bold.ttf', import.meta.url),
  'DMSans-BoldItalic.ttf': () =>
    new URL('../assets/DMSans-BoldItalic.ttf', import.meta.url),
  'DMSans-Italic.ttf': () =>
    new URL('../assets/DMSans-Italic.ttf', import.meta.url),
  'DMSans-Regular.ttf': () =>
    new URL('../assets/DMSans-Regular.ttf', import.meta.url),
  'DMSerifDisplay-Italic.ttf': () =>
    new URL('../assets/DMSerifDisplay-Italic.ttf', import.meta.url),
  'DMSerifDisplay-Regular.ttf': () =>
    new URL('../assets/DMSerifDisplay-Regular.ttf', import.meta.url),
  'Gelasio-Bold.ttf': () =>
    new URL('../assets/Gelasio-Bold.ttf', import.meta.url),
  'Gelasio-BoldItalic.ttf': () =>
    new URL('../assets/Gelasio-BoldItalic.ttf', import.meta.url),
  'Gelasio-Italic.ttf': () =>
    new URL('../assets/Gelasio-Italic.ttf', import.meta.url),
  'Gelasio-Regular.ttf': () =>
    new URL('../assets/Gelasio-Regular.ttf', import.meta.url),
  'Heebo-Bold.ttf': () =>
    new URL('../assets/Heebo-Bold.ttf', import.meta.url),
  'Heebo-Regular.ttf': () =>
    new URL('../assets/Heebo-Regular.ttf', import.meta.url),
  'Inter-Bold.ttf': () =>
    new URL('../assets/Inter-Bold.ttf', import.meta.url),
  'Inter-BoldItalic.ttf': () =>
    new URL('../assets/Inter-BoldItalic.ttf', import.meta.url),
  'Inter-Italic.ttf': () =>
    new URL('../assets/Inter-Italic.ttf', import.meta.url),
  'Inter-Regular.ttf': () =>
    new URL('../assets/Inter-Regular.ttf', import.meta.url),
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
  'Montserrat-Bold.ttf': () =>
    new URL('../assets/Montserrat-Bold.ttf', import.meta.url),
  'Montserrat-BoldItalic.ttf': () =>
    new URL('../assets/Montserrat-BoldItalic.ttf', import.meta.url),
  'Montserrat-Italic.ttf': () =>
    new URL('../assets/Montserrat-Italic.ttf', import.meta.url),
  'Montserrat-Regular.ttf': () =>
    new URL('../assets/Montserrat-Regular.ttf', import.meta.url),
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
  'OpenSans-Bold.ttf': () =>
    new URL('../assets/OpenSans-Bold.ttf', import.meta.url),
  'OpenSans-BoldItalic.ttf': () =>
    new URL('../assets/OpenSans-BoldItalic.ttf', import.meta.url),
  'OpenSans-Italic.ttf': () =>
    new URL('../assets/OpenSans-Italic.ttf', import.meta.url),
  'OpenSans-Regular.ttf': () =>
    new URL('../assets/OpenSans-Regular.ttf', import.meta.url),
  'Oswald-Bold.ttf': () =>
    new URL('../assets/Oswald-Bold.ttf', import.meta.url),
  'Oswald-Regular.ttf': () =>
    new URL('../assets/Oswald-Regular.ttf', import.meta.url),
  'Poppins-Bold.ttf': () =>
    new URL('../assets/Poppins-Bold.ttf', import.meta.url),
  'Poppins-BoldItalic.ttf': () =>
    new URL('../assets/Poppins-BoldItalic.ttf', import.meta.url),
  'Poppins-Italic.ttf': () =>
    new URL('../assets/Poppins-Italic.ttf', import.meta.url),
  'Poppins-Regular.ttf': () =>
    new URL('../assets/Poppins-Regular.ttf', import.meta.url),
  'Roboto-Bold.ttf': () =>
    new URL('../assets/Roboto-Bold.ttf', import.meta.url),
  'Roboto-BoldItalic.ttf': () =>
    new URL('../assets/Roboto-BoldItalic.ttf', import.meta.url),
  'Roboto-Italic.ttf': () =>
    new URL('../assets/Roboto-Italic.ttf', import.meta.url),
  'Roboto-Regular.ttf': () =>
    new URL('../assets/Roboto-Regular.ttf', import.meta.url),
  'SourceSans3-Bold.ttf': () =>
    new URL('../assets/SourceSans3-Bold.ttf', import.meta.url),
  'SourceSans3-BoldItalic.ttf': () =>
    new URL('../assets/SourceSans3-BoldItalic.ttf', import.meta.url),
  'SourceSans3-Italic.ttf': () =>
    new URL('../assets/SourceSans3-Italic.ttf', import.meta.url),
  'SourceSans3-Regular.ttf': () =>
    new URL('../assets/SourceSans3-Regular.ttf', import.meta.url),
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

const CJK_FILES = new Set(
  BUNDLED_FONTS.filter((face) => face.script?.startsWith('cjk-')).map((face) => face.file),
);

async function assetUrl(
  { file, script }: BundledFontFace,
  baseUrl: URL | undefined,
): Promise<URL> {
  if (baseUrl !== undefined) return new URL(file, baseUrl);
  const local = FONT_ASSET_URLS[file];
  if (local) return local();
  if (script?.startsWith('cjk-') || CJK_FILES.has(file)) {
    const cjk = await loadCjkAssetUrls();
    const resolveCjk = cjk?.[file];
    if (resolveCjk) return resolveCjk();
    if (!cjk) {
      throw new Error(
        `Bundled font ${file} needs the optional CJK add-on — install @betteroffice/fonts-cjk`,
      );
    }
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
  return assetUrl(face, baseUrl).then((url) => loadFontBytes(face, url));
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
