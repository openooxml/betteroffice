/**
 * `@betteroffice/xlsx-react` — the React chrome for the xlsx editor. Framework
 * glue only; all compute lives in `@betteroffice/xlsx`.
 */

export { XlsxEditor, XlsxSaveRefusedError } from './XlsxEditor';
export type {
  XlsxEditorProps,
  XlsxEditorApi,
  XlsxEditorCollaborationOptions,
} from './XlsxEditor';
export { EditorToolbar } from './components/EditorToolbar';
export {
  Toolbar,
  XlsxToolbar,
  type ToolbarProps,
  type ToolbarMode,
  type SelectionFormatting,
  type SelectionShape,
  type FormattingAction,
  type NumberFormat,
  type BorderPreset,
  type BorderStyle,
  type HorizontalAlignment,
  type VerticalAlignment,
  type TextWrapping,
  type MergeAction,
} from './components/Toolbar';
export {
  EditorToolbarContext,
  useEditorToolbar,
  type EditorToolbarProps,
} from './components/EditorToolbarContext';
export {
  ToolbarButton,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  ToolbarMenuSeparator,
  ToolbarSeparator,
  type ToolbarButtonProps,
  type ToolbarDropdownProps,
  type ToolbarMenuItemProps,
} from './components/ui/ToolbarPrimitives';
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';

export {
  XlsxCommandProvider,
  type XlsxCommandProviderProps,
} from './commands/XlsxCommandProvider';
export {
  useXlsxCommands,
  useXlsxCommandState,
  useXlsxCommand,
  type XlsxBoundCommand,
} from './commands/hooks';
export type {
  CommandReason,
  CommandState,
  JsonValue,
  XlsxCommandArgs,
  XlsxCommandDescriptor,
  XlsxCommandDisabledCode,
  XlsxCommandFailureCode,
  XlsxCommandId,
  XlsxCommandOption,
  XlsxCommandResult,
  XlsxCommandShortcut,
  XlsxCommandState,
  XlsxCommandStatus,
  XlsxCommandStore,
  XlsxCommandValues,
  XlsxNumberFormatValue,
  XlsxSelectCommandId,
} from './commands/types';
export {
  ToolbarCommand,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  type ToolbarCommandArgs,
  type ToolbarCommandButtonProps,
  type ToolbarCommandProps,
  type ToolbarCommandSelectProps,
} from './components/toolbar/ToolbarCommand';
export { ToolbarOverflow, type ToolbarOverflowProps } from './components/toolbar/ToolbarOverflow';
export type { FormulaBarProps } from './components/toolbar/FormulaBar';
