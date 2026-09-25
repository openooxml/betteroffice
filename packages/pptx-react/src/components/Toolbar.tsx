import { useContext } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { PptxCommandContext } from '../commands/PptxCommandProvider';
import { useTranslation } from '../i18n';
import { EditorToolbarContext, ToolbarModeContext } from './EditorToolbarContext';
import { useLegacyToolbarStore } from './toolbar/legacyCommands';
import { ShapeToolMenu, ToolbarCommand, ToolbarCommandButton } from './toolbar/ToolbarCommand';
import { LegacyToolbarContent, ToolbarRail } from './toolbar/ToolbarOverflow';
import type {
  FormattingAction,
  PptxEditorTool,
  PptxZoom,
  SelectionFormatting,
  ShapeFormatting,
  ShapeFormattingAction,
  SlideLayoutOption,
} from './toolbarTypes';
import { ToolbarGroup, ToolbarSeparator } from './ui/ToolbarPrimitives';

export {
  SHAPE_PRESETS,
  type FormattingAction,
  type PptxEditorTool,
  type PptxShapePreset,
  type PptxZoom,
  type SelectionFormatting,
  type ShapeFormatting,
  type ShapeFormattingAction,
  type ShapeZOrder,
  type SlideLayoutOption,
} from './toolbarTypes';

/**
 * Props of the legacy toolbar, which shows the state it is given and reports
 * actions through callbacks. Prefer `mode="commands"`, which binds controls to
 * the editor's command store.
 */
export interface ToolbarProps {
  /** `legacy` (the default outside a command-mode `EditorToolbar`) binds controls to these props. */
  mode?: 'legacy';
  currentFormatting?: SelectionFormatting;
  textSelectionActive?: boolean;
  onFormat?: (action: FormattingAction) => void;
  currentShapeFormatting?: ShapeFormatting;
  shapeSelectionActive?: boolean;
  /** Enables the arrange (z-order) menu for any selected object. */
  shapeArrangeActive?: boolean;
  onShapeFormat?: (action: ShapeFormattingAction) => void;
  onInsertSlide?: (layoutPartPath?: string | null) => void;
  onInsertImage?: () => void;
  slideLayouts?: readonly SlideLayoutOption[];
  currentLayoutPartPath?: string | null;
  onSave?: () => void;
  onExportPng?: () => void;
  onUndo?: () => void;
  onRedo?: () => void;
  canUndo?: boolean;
  canRedo?: boolean;
  zoom?: PptxZoom;
  onZoomChange?: (zoom: PptxZoom) => void;
  activeTool?: PptxEditorTool;
  onToolChange?: (tool: PptxEditorTool) => void;
  fontFamilies?: readonly string[];
  fontSizes?: readonly number[];
  disabled?: boolean;
  className?: string;
  style?: CSSProperties;
  /** Host controls appended after the default ones. */
  children?: ReactNode;
}

type LegacyStateKey = Exclude<keyof ToolbarProps, 'mode' | 'className' | 'style' | 'children'>;

/** Props of a toolbar bound to the nearest editor's commands. */
export type CommandToolbarProps = {
  /** Binds every control to the command store of the nearest editor or `PptxCommandProvider`. */
  mode: 'commands';
  /** The complete arrangement; omit to render the default controls. */
  children?: ReactNode;
  className?: string;
  style?: CSSProperties;
} & { [K in LegacyStateKey]?: never };

const LEGACY_STATE_KEYS: readonly LegacyStateKey[] = [
  'currentFormatting',
  'textSelectionActive',
  'onFormat',
  'currentShapeFormatting',
  'shapeSelectionActive',
  'shapeArrangeActive',
  'onShapeFormat',
  'onInsertSlide',
  'onInsertImage',
  'slideLayouts',
  'currentLayoutPartPath',
  'onSave',
  'onExportPng',
  'onUndo',
  'onRedo',
  'canUndo',
  'canRedo',
  'zoom',
  'onZoomChange',
  'activeTool',
  'onToolChange',
  'fontFamilies',
  'fontSizes',
  'disabled',
];

/** Throws when command-mode chrome is also given legacy state or callbacks. */
export function rejectLegacyState(props: object, component: string): void {
  const conflicting = LEGACY_STATE_KEYS.filter(
    (key) => (props as Record<string, unknown>)[key] !== undefined
  );
  if (conflicting.length > 0) {
    throw new Error(
      `${component} in command mode reads state from the editor's commands; remove ${conflicting.join(
        ', '
      )} or use mode="legacy".`
    );
  }
}

/** The default arrangement of built-in controls. */
function DefaultToolbarItems() {
  const { t } = useTranslation();
  return (
    <>
      <ToolbarGroup label={t('toolbar.groups.file')}>
        <ToolbarCommandButton id="save" />
        <ToolbarCommandButton id="exportPng" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.slides')}>
        <ToolbarCommand id="insertSlide" />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.history')}>
        <ToolbarCommandButton id="undo" />
        <ToolbarCommandButton id="redo" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.zoom')}>
        <ToolbarCommand id="zoom" />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.tools')}>
        <ToolbarCommandButton id="tool" args={{ value: 'select' }} />
        <ToolbarCommandButton id="tool" args={{ value: 'textBox' }} />
        <ToolbarCommandButton id="insertImage" />
        <ShapeToolMenu />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.font')}>
        <ToolbarCommand id="fontFamily" />
        <ToolbarCommandButton id="fontSizeStep" args={{ direction: 'decrease' }} />
        <ToolbarCommand id="fontSize" />
        <ToolbarCommandButton id="fontSizeStep" args={{ direction: 'increase' }} />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.text')}>
        <ToolbarCommandButton id="bold" />
        <ToolbarCommandButton id="italic" />
        <ToolbarCommandButton id="underline" />
        <ToolbarCommand id="textColor" />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.alignment')}>
        <ToolbarCommand id="alignment" />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.shape')}>
        <ToolbarCommand id="shapeFill" />
        <ToolbarCommand id="shapeStrokeColor" />
        <ToolbarCommand id="shapeStrokeWidth" />
        <ToolbarCommand id="shapeAdjustment" />
        <ToolbarCommand id="zOrder" />
      </ToolbarGroup>
    </>
  );
}

function stripUndefined<T extends object>(value: T): Partial<T> {
  const result: Partial<T> = {};
  for (const key of Object.keys(value) as Array<keyof T>) {
    if (value[key] !== undefined) result[key] = value[key];
  }
  return result;
}

function LegacyToolbar(explicitProps: ToolbarProps) {
  const context = useContext(EditorToolbarContext);
  const props = context ? { ...context, ...stripUndefined(explicitProps) } : explicitProps;
  const store = useLegacyToolbarStore(props);
  return (
    <PptxCommandContext.Provider value={store}>
      <ToolbarRail className={props.className} style={props.style} testId="pptx-formatting-toolbar">
        <DefaultToolbarItems />
        {props.children ? <LegacyToolbarContent>{props.children}</LegacyToolbarContent> : null}
      </ToolbarRail>
    </PptxCommandContext.Provider>
  );
}

function CommandToolbar(props: CommandToolbarProps) {
  rejectLegacyState(props, 'Toolbar');
  return (
    <ToolbarRail className={props.className} style={props.style} testId="pptx-formatting-toolbar">
      {props.children ?? <DefaultToolbarItems />}
    </ToolbarRail>
  );
}

/**
 * The formatting rail, with overflow into an accessible "More" menu. In
 * command mode (explicit, or inherited from `EditorToolbar mode="commands"`)
 * children replace the default controls; in legacy mode they follow them.
 */
export function Toolbar(props: ToolbarProps | CommandToolbarProps) {
  const inherited = useContext(ToolbarModeContext);
  const mode = props.mode ?? inherited ?? 'legacy';
  return mode === 'commands' ? (
    <CommandToolbar {...(props as CommandToolbarProps)} mode="commands" />
  ) : (
    <LegacyToolbar {...(props as ToolbarProps)} />
  );
}

export { Toolbar as PptxToolbar };
