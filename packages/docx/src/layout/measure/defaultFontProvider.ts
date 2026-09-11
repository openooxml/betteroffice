/** Opt-in bundled-font provider resolution. @packageDocumentation */

import type { BundledFontProvider } from './fontRegistry';

/**
 * The slice of `@betteroffice/fonts` this module consumes. Structural on
 * purpose: the engine never names that package, so no bundler ever has to
 * resolve it.
 *
 * @public
 */
export interface BundledFontModule {
  createFontProvider(options?: { baseUrl?: string | URL }): BundledFontProvider;
}

/**
 * How to obtain bundled fonts. Nothing is loaded until one of `fonts` or
 * `load` is given — the engine never reaches for a font package on its own.
 *
 * @public
 */
export interface DefaultFontOptions {
  /** The imported `@betteroffice/fonts` module. */
  fonts?: BundledFontModule;
  /** Lazy alternative to {@link DefaultFontOptions.fonts}, resolving to the same module. */
  load?: () => Promise<unknown>;
  /** Fetch the faces from this base URL instead of the module's own assets. */
  baseUrl?: string | URL;
}

let options: DefaultFontOptions = {};
let resolved: Promise<BundledFontProvider | undefined> | undefined;

function configuredBaseUrl(baseUrl: string | URL | undefined): URL | undefined {
  if (baseUrl === undefined) return undefined;
  const href = typeof baseUrl === 'string' ? baseUrl : baseUrl.href;
  const locationHref = typeof location === 'undefined' ? undefined : location.href;
  try {
    return new URL(href, locationHref);
  } catch {
    throw new TypeError(
      `configureDefaultFonts baseUrl must be absolute when no browser location exists: ${href}`
    );
  }
}

function hasFontSource(): boolean {
  return options.fonts !== undefined || options.load !== undefined;
}

/** Configure the default provider and reset its memoized resolution. @public */
export function configureDefaultFonts(next: DefaultFontOptions): void {
  options = { ...next, baseUrl: configuredBaseUrl(next.baseUrl) };
  resolved = undefined;
  if (options.baseUrl !== undefined && !hasFontSource()) {
    console.warn(
      '[configureDefaultFonts] a baseUrl on its own loads nothing — the face manifest lives in ' +
        '@betteroffice/fonts. Pass the module too: configureDefaultFonts({ fonts, baseUrl }).'
    );
  }
}

async function load(): Promise<BundledFontProvider | undefined> {
  const { baseUrl, fonts, load: loader } = options;
  try {
    const module = fonts ?? ((await loader!()) as BundledFontModule);
    return module.createFontProvider(baseUrl === undefined ? undefined : { baseUrl });
  } catch (error) {
    console.warn('[configureDefaultFonts] the configured font source failed to load', error);
    return undefined;
  }
}

/** Resolve the configured provider, or `undefined` when no font source is set. @public */
export function resolveDefaultFontProvider(): Promise<BundledFontProvider | undefined> {
  if (!hasFontSource()) return Promise.resolve(undefined);
  if (resolved === undefined) {
    const promise = load();
    promise.then((provider) => {
      if (provider === undefined && resolved === promise) resolved = undefined;
    });
    resolved = promise;
  }
  return resolved;
}
