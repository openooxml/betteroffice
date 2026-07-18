/**
 * @betteroffice/docx-react
 *
 * Curated root entry for the documented React editor API. Advanced surfaces
 * stay public through explicit subpaths:
 * - `@betteroffice/docx-react/ui`
 * - `@betteroffice/docx-react/dialogs`
 * - `@betteroffice/docx-react/hooks`
 * - `@betteroffice/docx-react/plugin-api`
 *
 * Framework-agnostic document utilities live in `@betteroffice/docx`.
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
  type EditorMode,
} from './components/DocxEditor';
export { renderAsync, type RenderAsyncOptions, type DocxEditorHandle } from './renderAsync';

// Rust measurement — the `measurementFontProvider` prop's interface,
// re-exported from `@betteroffice/docx` so consumers can implement a
// provider without adding the core package to their dependency tree.
export type { BundledFontProvider } from '@betteroffice/docx/layout';

// Document factory helpers — re-exported from `@betteroffice/docx` so
// the common "spawn a blank editor" affordance is available without forcing
// consumers to add `-core` to their dependency tree alongside `-react`.
export {
  createEmptyDocument,
  createDocumentWithText,
  type CreateEmptyDocumentOptions,
} from '@betteroffice/docx';

// i18n contract — runtime only. Locale string types (LocaleStrings,
// Translations, PartialLocaleStrings, TranslationKey) live in
// `@betteroffice/docx-i18n`; import them from there.
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';
