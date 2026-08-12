/**
 * The default measurement font provider — how core gets real font bytes when
 * a host injects no provider of its own.
 *
 * Without this, a family the document does not embed has no bytes to measure
 * with and the engine falls back to synthetic metrics. Measured across 813
 * real-world documents scored against Word's own `docProps/app.xml <Pages>`,
 * that costs 15.4 points of exact page-count accuracy (61.9% -> 46.5%).
 *
 * `@betteroffice/fonts` is resolved through an OPTIONAL DYNAMIC `import()`,
 * never a static one. The bundled faces are 7.9 MB (33 MB more with the CJK
 * add-on); a static edge would put them in every consumer's dependency graph,
 * including consumers who inject their own provider. Dynamic keeps the cost on
 * the hosts that actually use it, and the package is declared an optional peer
 * so npm never installs it uninvited.
 *
 * Resolution is memoized including the miss, so an absent package costs one
 * failed import per session, not one per family.
 *
 * @packageDocumentation
 */

import type { BundledFontProvider } from './fontRegistry';

/**
 * The real module type, so a signature change in `@betteroffice/fonts` breaks
 * this file at compile time rather than at a consumer's first layout pass. The
 * package is a devDependency for exactly this; it stays an optional peer at
 * runtime, and this type is local so it never reaches the published `.d.ts`.
 */
type BundledFontsModule = typeof import('@betteroffice/fonts');

/** @public */
export interface DefaultFontOptions {
  /**
   * Serve the bundled faces from here instead of the package's own assets.
   *
   * Same-origin is the default on purpose: a CDN default would leak
   * document-font usage to a third party and break offline and strict-CSP
   * deployments. Point this at a CDN deliberately, or not at all.
   */
  baseUrl?: string | URL;
  /**
   * Override how `@betteroffice/fonts` is loaded. The escape hatch for hosts
   * whose bundler cannot leave an uninstalled optional peer unresolved, and
   * the seam tests use to simulate the package's absence.
   */
  load?: () => Promise<unknown>;
}

let options: DefaultFontOptions = {};
let resolved: Promise<BundledFontProvider | undefined> | undefined;

/**
 * Configure the default provider. Call before the first layout pass; it resets
 * the memoized resolution, so a later call re-resolves.
 *
 * @public
 */
export function configureDefaultFonts(next: DefaultFontOptions): void {
  options = { ...next };
  resolved = undefined;
}

async function load(): Promise<BundledFontProvider | undefined> {
  const { baseUrl, load: loader } = options;
  try {
    const imported = await (loader ? loader() : import('@betteroffice/fonts'));
    const module = imported as BundledFontsModule;
    return module.createFontProvider(baseUrl === undefined ? undefined : { baseUrl });
  } catch {
    return undefined;
  }
}

/**
 * The bundled provider, or `undefined` when `@betteroffice/fonts` is not
 * installed. Never throws — an absent package degrades to synthetic metrics
 * (loudly, see the registry's chain warning), it does not break rendering.
 *
 * @public
 */
export function resolveDefaultFontProvider(): Promise<BundledFontProvider | undefined> {
  resolved ??= load();
  return resolved;
}
