/**
 * @betteroffice/docx-react
 *
 * Curated root entry for the documented React editor API.
 *
 * @packageDocumentation
 * @public
 */

import { version as packageVersion } from '../package.json';

export const VERSION: string = packageVersion;

// Main editor contract
export {
  DocxEditor,
  type DocxEditorProps,
  type DocxEditorRef,
  type DocxEditorCollaborationOptions,
  type EditorMode,
} from './components/DocxEditor';

export type { BundledFontProvider } from '@betteroffice/docx/layout';
export {
  configureDefaultFonts,
  type BundledFontModule,
  type DefaultFontOptions,
} from '@betteroffice/docx/layout';

// i18n contract — runtime only. Locale string types (LocaleStrings,
// Translations, PartialLocaleStrings, TranslationKey) live in
// `@betteroffice/docx-i18n`; import them from there.
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';
