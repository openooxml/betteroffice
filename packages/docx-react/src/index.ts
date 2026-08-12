/**
 * @betteroffice/docx-react
 *
 * Curated root entry for the documented React editor API.
 *
 * @packageDocumentation
 * @public
 */

export const VERSION = '0.0.2';

// Main editor contract
export {
  DocxEditor,
  type DocxEditorProps,
  type DocxEditorRef,
  type DocxEditorCollaborationOptions,
  type EditorMode,
} from './components/DocxEditor';

// Rust measurement — the `measurementFontProvider` prop's interface, plus the
// knob that points the default bundled faces at a CDN. Re-exported from
// `@betteroffice/docx` so consumers need not add the core package to their
// dependency tree.
export type { BundledFontProvider } from '@betteroffice/docx/layout';
export { configureDefaultFonts, type DefaultFontOptions } from '@betteroffice/docx/layout';

// i18n contract — runtime only. Locale string types (LocaleStrings,
// Translations, PartialLocaleStrings, TranslationKey) live in
// `@betteroffice/docx-i18n`; import them from there.
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';
