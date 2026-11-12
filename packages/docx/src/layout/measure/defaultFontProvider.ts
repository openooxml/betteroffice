/** Lazy default-font provider resolution. @packageDocumentation */

import type { BundledFontLoader, BundledFontProvider } from './fontRegistry';
import type { FontDecoder } from '../../../../../shared/office-session/fontDecoder';
import { decodeFontFace } from '../../wasm/fontDecoder';

export const IMPLICIT_FONT_COVERAGE = '@betteroffice/fonts/coverage';
export const IMPLICIT_SYMBOL_SHAPING = '@betteroffice/fonts/coverage-symbols';

export class ImplicitFontAbsentError extends Error {}

const implicitProviders = new WeakSet<BundledFontProvider>();

export function isImplicitFontProvider(provider: BundledFontProvider | undefined): boolean {
  return provider !== undefined && implicitProviders.has(provider);
}

const hostedFamiliesKey = Symbol.for('@betteroffice/docx/hostedFontFamilies');
const standInKey = Symbol.for('@betteroffice/docx/standInFamily');
const STAND_IN_TIMEOUT_MS = 30_000;

interface StandInHook {
  standsIn(family: string): boolean;
  report(family: string): boolean;
}

export function hostedFontFamilies(source: object | undefined): ReadonlySet<string> | undefined {
  return (source as { [hostedFamiliesKey]?: ReadonlySet<string> } | undefined)?.[hostedFamiliesKey];
}

export function excludeHostedFontFamilies(
  provider: BundledFontProvider | undefined,
  families: ReadonlySet<string> | undefined
): BundledFontProvider | undefined {
  if (!provider || !isImplicitFontProvider(provider) || !families?.size) return provider;
  const hosted = (family: string) => families.has(family.trim().toLowerCase());
  const filtered: BundledFontProvider = {
    ...provider,
    resolve: (family, bold, italic) => hosted(family) ? undefined : provider.resolve(family, bold, italic),
    resolveFamily: (family, bold, italic) => hosted(family) ? undefined : provider.resolveFamily?.(family, bold, italic),
    resolveGlyphFallback: (family, bold, italic, office) => hosted(family) ? undefined : provider.resolveGlyphFallback?.(family, bold, italic, office),
    resolveLastResort: (family, bold, italic, office) => hosted(family) ? undefined : provider.resolveLastResort?.(family, bold, italic, office),
  };
  implicitProviders.add(filtered);
  return filtered;
}

function implicitFontProvider(
  manager: typeof import('@betteroffice/fonts'),
  runtime: { decoder?: FontDecoder },
  standIns: boolean
): BundledFontProvider {
  const provider = manager.createFontProvider(undefined, runtime);
  const providesFamily = (family: string) =>
    manager.resolveMetricCompatFamily(family) !== undefined ||
    manager.resolveBundledFamilyFace(family, false, false) !== undefined;
  const fallsBack = (family: string) => standIns || providesFamily(family);
  const standsIn = (family: string) => standIns && !providesFamily(family);
  const bounded = new WeakMap<BundledFontLoader, BundledFontLoader>();
  const standIn = (loader: BundledFontLoader | undefined): BundledFontLoader | undefined => {
    if (!loader) return undefined;
    let wrapped = bounded.get(loader);
    if (!wrapped) {
      const due = (load: () => Promise<ArrayBuffer>) => {
        let timer: ReturnType<typeof setTimeout> | undefined;
        return Promise.race([
          load(),
          new Promise<never>((_, reject) => {
            timer = setTimeout(() => reject(new Error('the stand-in face did not load in time')), STAND_IN_TIMEOUT_MS);
          }),
        ]).finally(() => clearTimeout(timer));
      };
      const read = loader.readOnlyBytes;
      wrapped = Object.assign(() => due(loader), loader, read && { readOnlyBytes: () => due(read) });
      bounded.set(loader, wrapped);
    }
    return wrapped;
  };
  const checked = (
    loader: BundledFontLoader | undefined,
    face: import('@betteroffice/fonts').BundledFontFace | undefined
  ): BundledFontLoader | undefined => {
    if (!face) return undefined;
    if (!loader || !face.script) return loader;
    const verify = async (load: () => Promise<ArrayBuffer>) => {
      let bytes: ArrayBuffer;
      try { bytes = await load(); }
      catch { throw new ImplicitFontAbsentError(face.family); }
      if (bytes.byteLength !== face.byteLength) throw new ImplicitFontAbsentError(face.family);
      return bytes;
    };
    const read = loader.readOnlyBytes;
    return Object.assign(() => verify(loader), loader, read && { readOnlyBytes: () => verify(read) });
  };
  const checkedFamilyFallback = (
    loader: BundledFontLoader | undefined,
    family: string,
    bold: boolean,
    italic: boolean
  ): BundledFontLoader | undefined => {
    const face = manager.resolveMetricCompatFace(family, bold, italic) ??
      manager.resolveBundledFamilyFace(family, bold, italic) ??
      manager.resolveBundledFamilyFace(family, false, false);
    if (!loader || !face?.script) return loader;
    const primary = checked(
      provider.resolve(family, bold, italic) ?? provider.resolveFamily(family, bold, italic) ??
        provider.resolveFamily(family, false, false),
      face
    );
    const after = async <T>(load: () => Promise<T>) => {
      if (!primary) throw new ImplicitFontAbsentError(face.family);
      await (primary.readOnlyBytes ?? primary)();
      return load();
    };
    const read = loader.readOnlyBytes;
    return Object.assign(() => after(loader), loader, read && { readOnlyBytes: () => after(read) });
  };
  const implicit: BundledFontProvider = {
    ...provider,
    resolve: (family, bold, italic) => checked(
      provider.resolve(family, bold, italic) ?? provider.resolveFamily(family, bold, italic),
      manager.resolveMetricCompatFace(family, bold, italic) ?? manager.resolveBundledFamilyFace(family, bold, italic)
    ),
    resolveFamily: (family, bold, italic) => checked(
      provider.resolveFamily(family, bold, italic), manager.resolveBundledFamilyFace(family, bold, italic)
    ),
    resolveScriptFallback: (script, bold, italic, family) => checked(
      provider.resolveScriptFallback(script, bold, italic, family),
      manager.resolveScriptFallbackFace(script, bold, italic, family)
    ),
    resolveGlyphFallback: (family, bold, italic, office) => {
      if (!fallsBack(family)) return undefined;
      const loader = checkedFamilyFallback(
        checked(
          provider.resolveGlyphFallback?.(family, bold, italic, office),
          manager.resolveGlyphFallbackFace(family, bold, italic, office)
        ), family, bold, italic
      );
      return standsIn(family) ? standIn(loader) : loader;
    },
    resolveLastResort: (family, bold, italic, office) => {
      if (!fallsBack(family)) return undefined;
      const loader = checkedFamilyFallback(
        checked(
          provider.resolveLastResort(family, bold, italic, office),
          manager.resolveLastResortFace(family, bold, italic, office)
        ), family, bold, italic
      );
      return standsIn(family) ? standIn(loader) : loader;
    },
  };
  const reported = new Set<string>();
  const standInHook: StandInHook = {
    standsIn,
    report: (family) => {
      const key = family.trim().toLowerCase();
      if (reported.has(key)) return false;
      reported.add(key);
      return true;
    },
  };
  Object.assign(implicit, { [standInKey]: standInHook });
  implicitProviders.add(implicit);
  return implicit;
}

/** The consumed font-manager contract. @public */
export interface BundledFontModule {
  createFontProvider(options?: { baseUrl?: string | URL }, runtime?: { decoder?: FontDecoder }): BundledFontProvider;
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
const engineRuntime = { decoder: decodeFontFace };
let resolved = new WeakMap<object, Promise<BundledFontProvider | undefined>>();
let resolvedWithStandIns = new WeakMap<object, Promise<BundledFontProvider | undefined>>();
let baseLatinResolved = new WeakMap<object, Promise<BundledFontProvider>>();

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
  resolved = new WeakMap();
  resolvedWithStandIns = new WeakMap();
  baseLatinResolved = new WeakMap();
}

async function load(runtime: { decoder?: FontDecoder }, standIns: boolean): Promise<BundledFontProvider | undefined> {
  const { baseUrl, fonts, load: loader } = options;
  try {
    if (fonts === undefined && loader === undefined && baseUrl === undefined) {
      return implicitFontProvider(await import('@betteroffice/fonts'), runtime, standIns);
    }
    const module = fonts ?? (
      (loader ? await loader() : await import('@betteroffice/fonts')) as BundledFontModule
    );
    return module.createFontProvider(baseUrl === undefined ? undefined : { baseUrl }, runtime);
  } catch (error) {
    console.warn('[configureDefaultFonts] the configured font source failed to load', error);
    return undefined;
  }
}

/** Resolve the default provider, or `undefined` after a retryable failure. @public */
export function resolveDefaultFontProvider(
  runtime: { decoder?: FontDecoder } = engineRuntime,
  standIns = false
): Promise<BundledFontProvider | undefined> {
  const cache = standIns && !hasConfiguredFontSource() && options.baseUrl === undefined ? resolvedWithStandIns : resolved;
  let pending = cache.get(runtime);
  if (pending === undefined) {
    const promise = load(runtime, standIns);
    promise.then((provider) => {
      if (provider === undefined && cache.get(runtime) === promise) cache.delete(runtime);
    });
    cache.set(runtime, promise);
    pending = promise;
  }
  return pending;
}

/**
 * The configured provider, else every bundled face including the last-resort
 * ones, for engines with no browser fonts to fall back on.
 */
export async function resolveCompleteFontProvider(
  runtime: { decoder?: FontDecoder } = engineRuntime
): Promise<BundledFontProvider | undefined> {
  return (await resolveCompleteFontSelection(runtime)).provider;
}

/** @internal {@link resolveCompleteFontProvider}, with whether the provider is the configured one. */
export async function resolveCompleteFontSelection(
  runtime: { decoder?: FontDecoder } = engineRuntime
): Promise<{ provider: BundledFontProvider | undefined; configured: boolean }> {
  if (hasConfiguredFontSource() || options.baseUrl !== undefined) {
    return { provider: await resolveDefaultFontProvider(runtime), configured: true };
  }
  return { provider: (await import('@betteroffice/fonts')).createFontProvider(undefined, runtime), configured: false };
}

/** @internal State the lazily loaded backstop provider resolves against. */
export function baseLatinFontState(): {
  baseUrl: string | URL | undefined;
  cache: WeakMap<object, Promise<BundledFontProvider>>;
  runtime: { decoder?: FontDecoder };
} {
  return { baseUrl: options.baseUrl, cache: baseLatinResolved, runtime: engineRuntime };
}
