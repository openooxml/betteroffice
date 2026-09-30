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
  type DocxBoundPluginCommand,
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

// Plugins: host-owned extensions with granted access to the editor
export { defineDocxPlugin } from './plugins/defineDocxPlugin';
export { DocxPluginToolbar } from './plugins/DocxPluginToolbar';
export type {
  DocxEditorPluginProps,
  DocxAnchorGeometryResult,
  DocxAnchorRect,
  DocxGeometryTarget,
  DocxPlugin,
  DocxPluginCommand,
  DocxPluginCommandClient,
  DocxPluginContext,
  DocxPluginDefinition,
  DocxPluginEditClient,
  DocxPluginError,
  DocxPluginErrorPhase,
  DocxPluginEvent,
  DocxPluginFailureCode,
  DocxPluginGeometry,
  DocxPluginGrant,
  DocxPluginLayout,
  DocxPluginNavigation,
  DocxPluginNavigationFailureCode,
  DocxPluginPanel,
  DocxPluginPointPosition,
  DocxPluginReadClient,
  DocxPluginRect,
  DocxPluginRefusal,
  DocxPluginSelection,
  DocxPluginSidebarItem,
  DocxPluginSnapshot,
  MaybePromise,
  PluginCleanupReason,
  PluginGrant,
  PluginLoadReason,
} from './plugins/types';
export type {
  DocxPluginCommandDescriptor,
  DocxPluginCommandId,
  DocxPluginCommandResult,
  DocxPluginCommandState,
} from './commands/types';
export type { DocxPointPosition, SelectionState } from './components/DocxEditor/types';
export type {
  RenderedDomContext,
  PositionCoordinates,
  PointPosition,
} from '@betteroffice/docx/plugin-api';

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

// Host-proposal types — re-exported so hosts don't need to import
// `@betteroffice/docx/yrs` directly for `proposeChanges`/`setProposalStates`/`getProposals`.
export type {
  DocxOccurrence,
  DocxProposalFailure,
  DocxProposalInput,
  DocxProposalRecord,
  DocxProposalRequest,
  DocxProposalResult,
  DocxProposalSnapshot,
  DocxProposalState,
  DocxProposalStateRequest,
} from '@betteroffice/docx/yrs';
