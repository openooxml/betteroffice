// @betteroffice/docx declaration layout/measure/defaultFontProvider.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Lazy default-font provider resolution. @packageDocumentation */
import type { BundledFontProvider } from './fontRegistry';
import type { FontDecoder } from '../../../../../shared/office-session/fontDecoder';
export declare const IMPLICIT_FONT_COVERAGE = "@betteroffice/fonts/coverage";
export declare const IMPLICIT_SYMBOL_SHAPING = "@betteroffice/fonts/coverage-symbols";
export declare class ImplicitFontAbsentError extends Error {
}
export declare function isImplicitFontProvider(provider: BundledFontProvider | undefined): boolean;
export declare function hostedFontFamilies(source: object | undefined): ReadonlySet<string> | undefined;
export declare function excludeHostedFontFamilies(provider: BundledFontProvider | undefined, families: ReadonlySet<string> | undefined): BundledFontProvider | undefined;
/** The consumed font-manager contract. @public */
export interface BundledFontModule {
    createFontProvider(options?: {
        baseUrl?: string | URL;
    }, runtime?: {
        decoder?: FontDecoder;
    }): BundledFontProvider;
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
/** Whether a default font source is available. */
export declare function hasDefaultFontSource(): boolean;
/** Whether {@link configureDefaultFonts} was given a font source. */
export declare function hasConfiguredFontSource(): boolean;
/** Configure the default provider and reset its memoized resolution. @public */
export declare function configureDefaultFonts(next: DefaultFontOptions): void;
/** Resolve the default provider, or `undefined` after a retryable failure. @public */
export declare function resolveDefaultFontProvider(runtime?: {
    decoder?: FontDecoder;
}, standIns?: boolean): Promise<BundledFontProvider | undefined>;
/**
 * The configured provider, else every bundled face including the last-resort
 * ones, for engines with no browser fonts to fall back on.
 */
export declare function resolveCompleteFontProvider(runtime?: {
    decoder?: FontDecoder;
}): Promise<BundledFontProvider | undefined>;
/** @internal State the lazily loaded backstop provider resolves against. */
export declare function baseLatinFontState(): {
    baseUrl: string | URL | undefined;
    cache: WeakMap<object, Promise<BundledFontProvider>>;
    runtime: {
        decoder?: FontDecoder;
    };
};
