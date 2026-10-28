/** Lazy default-font provider resolution. @packageDocumentation */

import type { BundledFontProvider } from './fontRegistry';

/** The consumed font-manager contract. @public */
export interface BundledFontModule {
  createFontProvider(options?: { baseUrl?: string | URL }): BundledFontProvider;
}

/** Override the lazy `@betteroffice/fonts` default. @public */
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

/** Whether a default font source is available. */
export function hasDefaultFontSource(): boolean {
  return true;
}

/** Whether {@link configureDefaultFonts} was given a font source. */
export function hasConfiguredFontSource(): boolean {
  return options.fonts !== undefined || options.load !== undefined;
}

/** Configure the default provider and reset its memoized resolution. @public */
export function configureDefaultFonts(next: DefaultFontOptions): void {
  options = { ...next, baseUrl: configuredBaseUrl(next.baseUrl) };
  resolved = undefined;
}

async function load(): Promise<BundledFontProvider | undefined> {
  const { baseUrl, fonts, load: loader } = options;
  try {
    const module = fonts ?? (
      (loader ? await loader() : await import('@betteroffice/fonts')) as BundledFontModule
    );
    return module.createFontProvider(baseUrl === undefined ? undefined : { baseUrl });
  } catch (error) {
    console.warn('[configureDefaultFonts] the configured font source failed to load', error);
    return undefined;
  }
}

/** Resolve the default provider, or `undefined` after a retryable failure. @public */
export function resolveDefaultFontProvider(): Promise<BundledFontProvider | undefined> {
  if (resolved === undefined) {
    const promise = load();
    promise.then((provider) => {
      if (provider === undefined && resolved === promise) resolved = undefined;
    });
    resolved = promise;
  }
  return resolved;
}
