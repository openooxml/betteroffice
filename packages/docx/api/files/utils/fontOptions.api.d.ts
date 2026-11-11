// @betteroffice/docx declaration utils/fontOptions.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Shared FontOption shape + normaliser used by FontPicker components
 * in both adapters. Lifted from packages/docx-react/src/components/ui/
 * normalizeFontFamilies.ts so the type definition has a single home.
 * @packageDocumentation
 * @public
 */
export interface FontOption {
    name: string;
    fontFamily: string;
    category?: 'sans-serif' | 'serif' | 'monospace' | 'other';
}
/**
 * Normalize a `fontFamilies` prop (mix of strings and FontOption
 * objects) into a uniform `FontOption[]`. Returns `undefined` for
 * `undefined` input so callers fall back to their built-in defaults.
 * Strings expand into the `'other'` group with no CSS fallback chain.
 */
export declare function normalizeFontFamilies(fontFamilies: ReadonlyArray<string | FontOption> | undefined): FontOption[] | undefined;
