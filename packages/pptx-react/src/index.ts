export { PptxEditor } from './PptxEditor';
export type {
  PptxEditorApi,
  PptxEditorCollaborationOptions,
  PptxEditorProps,
  PptxPointPosition,
  PptxTextSelection,
  PptxTextSelectionTarget,
} from './PptxEditor';
export {
  EditorToolbar,
  type EditorToolbarCommandProps,
} from './components/EditorToolbar';
export {
  Toolbar,
  PptxToolbar,
  SHAPE_PRESETS,
  type CommandToolbarProps,
  type ToolbarProps,
  type SelectionFormatting,
  type FormattingAction,
  type ShapeFormatting,
  type ShapeFormattingAction,
  type SlideLayoutOption,
  type PptxEditorTool,
  type PptxShapePreset,
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
  type ToolbarButtonProps,
  type ToolbarDropdownProps,
  type ToolbarMenuItemProps,
} from './components/ui/ToolbarPrimitives';
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';

export { PptxCommandProvider, type PptxCommandProviderProps } from './commands/PptxCommandProvider';
export {
  usePptxCommands,
  usePptxCommandState,
  usePptxCommand,
  type PptxBoundCommand,
  type PptxBoundPluginCommand,
} from './commands/hooks';
export type {
  CommandReason,
  CommandState,
  JsonValue,
  PptxCommandArgs,
  PptxCommandDescriptor,
  PptxCommandDisabledCode,
  PptxCommandFailureCode,
  PptxCommandId,
  PptxCommandOption,
  PptxCommandResult,
  PptxCommandShortcut,
  PptxCommandState,
  PptxCommandStatus,
  PptxCommandStore,
  PptxCommandValues,
  PptxPluginCommandDescriptor,
  PptxPluginCommandId,
  PptxPluginCommandResult,
  PptxPluginCommandState,
  PptxSelectCommandId,
  PptxZOrderMove,
} from './commands/types';
export {
  ToolbarCommand,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  type ToolbarCommandArgs,
  type ToolbarCommandButtonProps,
  type ToolbarCommandProps,
  type ToolbarCommandSelectProps,
  type PptxControlCommandId,
} from './components/toolbar/ToolbarCommand';
export { ToolbarOverflow, type ToolbarOverflowProps } from './components/toolbar/ToolbarOverflow';

export { definePptxPlugin } from './plugins/definePptxPlugin';
export { PptxPluginToolbar } from './plugins/PptxPluginToolbar';
export type {
  MaybePromise,
  PluginCleanupReason,
  PluginGrant,
  PluginLoadReason,
  PptxEditorPluginProps,
  PptxPlugin,
  PptxPluginCommand,
  PptxPluginCommandClient,
  PptxPluginContext,
  PptxPluginDefinition,
  PptxPluginEditClient,
  PptxPluginError,
  PptxPluginErrorPhase,
  PptxPluginEvent,
  PptxPluginFailureCode,
  PptxPluginGeometry,
  PptxPluginGrant,
  PptxPluginLayout,
  PptxPluginNavigation,
  PptxPluginNavigationFailureCode,
  PptxPluginNavigationOptions,
  PptxPluginNavigationResult,
  PptxPluginPanel,
  PptxPluginPointPosition,
  PptxPluginReadClient,
  PptxPluginRect,
  PptxPluginRefusal,
  PptxPluginSelection,
  PptxPluginSnapshot,
} from './plugins/types';
