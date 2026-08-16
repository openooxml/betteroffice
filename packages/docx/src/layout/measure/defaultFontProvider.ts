/** Optional bundled-font provider resolution. @packageDocumentation */

import type { BundledFontProvider } from './fontRegistry';

type BundledFontsModule = typeof import('@betteroffice/fonts');

/** @public */
export interface DefaultFontOptions {
  /** Serve faces from a base URL pinned when this configuration is applied. */
  baseUrl?: string | URL;
  /** Override loading of the optional `@betteroffice/fonts` peer. */
  load?: () => Promise<unknown>;
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

/** Configure the default provider and reset its memoized resolution. @public */
export function configureDefaultFonts(next: DefaultFontOptions): void {
  options = { ...next, baseUrl: configuredBaseUrl(next.baseUrl) };
  resolved = undefined;
}

async function load(): Promise<BundledFontProvider | undefined> {
  const { baseUrl, load: loader } = options;
  // Keep this syntactic try/catch: webpack tolerates a missing optional peer
  // here but fails the build when the import uses `.catch()`.
  try {
    const imported = await (loader ? loader() : import('@betteroffice/fonts'));
    const module = imported as BundledFontsModule;
    return module.createFontProvider(baseUrl === undefined ? undefined : { baseUrl });
  } catch {
    return undefined;
  }
}

/** Resolve the bundled provider, or `undefined` when its optional peer is absent. @public */
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
