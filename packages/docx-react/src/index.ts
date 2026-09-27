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

// Commands: one authority for built-in and host chrome
export { DocxCommandProvider, type DocxCommandProviderProps } from './commands/DocxCommandProvider';
export {
  useDocxCommands,
  useDocxCommandState,
  useDocxCommand,
  type DocxBoundCommand,
} from './commands/hooks';
export type {
  CommandReason,
  CommandState,
  JsonValue,
  DocxCommandArgs,
  DocxCommandDescriptor,
  DocxCommandDisabledCode,
  DocxCommandFailureCode,
  DocxCommandId,
  DocxCommandOption,
  DocxCommandOptionPreview,
  DocxCommandResult,
  DocxCommandShortcut,
  DocxCommandState,
  DocxCommandStatus,
  DocxCommandStore,
  DocxCommandValues,
  DocxImageTransform,
  DocxSelectCommandId,
  DocxTableAction,
  DocxTableValue,
  DocxTextColor,
} from './commands/types';

// Composable chrome
export {
  EditorToolbar,
  type EditorToolbarProps,
  type TitleBarProps,
  type LogoProps,
  type DocumentNameProps,
  type TitleBarRightProps,
  type ToolbarProps,
  type ToolbarReviewControlsProps,
} from './components/EditorToolbar';
export {
  ToolbarButton,
  ToolbarGroup,
  ToolbarSeparator,
  type ToolbarButtonProps,
  type ToolbarGroupProps,
} from './components/toolbar/ToolbarPrimitives';
export { ToolbarOverflow, type ToolbarOverflowProps } from './components/toolbar/ToolbarOverflow';
export {
  ToolbarCommand,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  type ToolbarCommandArgs,
  type ToolbarCommandButtonProps,
  type ToolbarCommandProps,
  type ToolbarCommandSelectProps,
} from './components/toolbar/ToolbarCommand';

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
