/**
 * Resolves the default provider through a memoized dynamic import of the
 * optional `@betteroffice/fonts` peer.
 * @packageDocumentation
 */

import type { BundledFontProvider } from './fontRegistry';

/** Compile-time contract for the optional peer. */
type BundledFontsModule = typeof import('@betteroffice/fonts');

/** @public */
export interface DefaultFontOptions {
  /** Serve bundled faces from this base URL instead of package assets. */
  baseUrl?: string | URL;
  /** Override loading of the optional `@betteroffice/fonts` peer. */
  load?: () => Promise<unknown>;
}

let options: DefaultFontOptions = {};
let resolved: Promise<BundledFontProvider | undefined> | undefined;

/** Configure the default provider and reset its memoized resolution. @public */
export function configureDefaultFonts(next: DefaultFontOptions): void {
  options = { ...next };
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
  resolved ??= load();
  return resolved;
}
