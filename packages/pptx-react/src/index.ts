export { PptxEditor } from './PptxEditor';
export type {
  PptxEditorApi,
  PptxEditorCollaborationOptions,
  PptxEditorProps,
  PptxTextSelection,
} from './PptxEditor';
export { EditorToolbar } from './components/EditorToolbar';
export {
  Toolbar,
  PptxToolbar,
  type ToolbarProps,
  type SelectionFormatting,
  type FormattingAction,
  type SlideLayoutOption,
  type PptxEditorTool,
  type PptxZoom,
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
