/**
 * `@betteroffice/xlsx-react` — the React chrome for the xlsx editor. Framework
 * glue only; all compute lives in `@betteroffice/xlsx`.
 */

export { XlsxEditor } from './XlsxEditor';
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
} from './components/ui/ToolbarPrimitives';
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';
