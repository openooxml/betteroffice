import { createContext, useContext, useMemo, useRef } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { useXlsxCommand } from '../commands/hooks';
import { XlsxCommandContext } from '../commands/XlsxCommandProvider';
import { useTranslation } from '../i18n';
import { XlsxPluginToolbar } from '../plugins/XlsxPluginToolbar';
import { EditorToolbarContext, ToolbarModeContext } from './EditorToolbarContext';
import { createLegacyToolbarStore } from './toolbar/legacyToolbarStore';
import { ToolbarCommand } from './toolbar/ToolbarCommand';
import { ToolbarLegacyContent, ToolbarRail } from './toolbar/ToolbarOverflow';
import { ToolbarButtonBase, ToolbarGroup, ToolbarSeparator } from './ui/ToolbarPrimitives';

/**
 * How a toolbar gets its state: `'legacy'` from props and callbacks,
 * `'commands'` from the editor's command store.
 * @experimental
 */
export type ToolbarMode = 'legacy' | 'commands';

export type NumberFormat =
  | 'automatic'
  | 'plainText'
  | 'number'
  | 'percent'
  | 'scientific'
  | 'currency'
  | 'date'
  | 'time'
  | 'custom';

export type BorderPreset =
  | 'all'
  | 'inner'
  | 'horizontal'
  | 'vertical'
  | 'outer'
  | 'left'
  | 'top'
  | 'right'
  | 'bottom'
  | 'none';

export type BorderStyle = 'solid' | 'dashed' | 'dotted' | 'double';
export type HorizontalAlignment = 'left' | 'center' | 'right';
export type VerticalAlignment = 'top' | 'middle' | 'bottom';
export type TextWrapping = 'overflow' | 'wrap' | 'clip';
export type MergeAction = 'all' | 'horizontal' | 'vertical' | 'unmerge';

export interface SelectionFormatting {
  paintFormat?: boolean;
  numberFormat?: NumberFormat;
  numberFormatPattern?: string;
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  textColor?: string;
  fillColor?: string;
  borderPreset?: BorderPreset;
  borderStyle?: BorderStyle;
  borderColor?: string;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  textWrapping?: TextWrapping;
}

export type FormattingAction =
  | 'paintFormat'
  | 'currency'
  | 'percent'
  | 'decreaseDecimal'
  | 'increaseDecimal'
  | 'bold'
  | 'italic'
  | 'strikethrough'
  | { type: 'numberFormat'; value: NumberFormat }
  | { type: 'fontFamily'; value: string }
  | { type: 'fontSize'; value: number }
  | { type: 'textColor'; value: string }
  | { type: 'fillColor'; value: string }
  | { type: 'borderPreset'; value: BorderPreset }
  | { type: 'borderStyle'; value: BorderStyle }
  | { type: 'borderColor'; value: string }
  | { type: 'horizontalAlignment'; value: HorizontalAlignment }
  | { type: 'verticalAlignment'; value: VerticalAlignment }
  | { type: 'textWrapping'; value: TextWrapping };

export interface SelectionShape {
  rows: number;
  columns: number;
  canUnmerge?: boolean;
}

export interface ToolbarProps {
  currentFormatting?: SelectionFormatting;
  selectionShape?: SelectionShape;
  onFormat?: (action: FormattingAction) => void;
  onMerge?: (action: MergeAction) => void;
  onSearchMenus?: () => void;
  onUndo?: () => void;
  onRedo?: () => void;
  canUndo?: boolean;
  canRedo?: boolean;
  onPrint?: () => void;
  zoom?: number;
  onZoomChange?: (zoom: number) => void;
  fontFamilies?: readonly string[];
  fontSizes?: readonly number[];
  disabled?: boolean;
  className?: string;
  style?: CSSProperties;
  /** Legacy mode appends these after the built-in controls; commands mode renders only these. */
  children?: ReactNode;
  /**
   * `'commands'` binds the built-in controls to the editor's commands, and
   * rejects the state and callback props above. Defaults to the mode of the
   * surrounding `EditorToolbar`, else `'legacy'`.
   * @experimental
   */
  mode?: ToolbarMode;
}


const LEGACY_PROPS = [
  'currentFormatting',
  'selectionShape',
  'onFormat',
  'onMerge',
  'onSearchMenus',
  'onUndo',
  'onRedo',
  'canUndo',
  'canRedo',
  'onPrint',
  'zoom',
  'onZoomChange',
  'fontFamilies',
  'fontSizes',
  'disabled',
] as const satisfies readonly (keyof ToolbarProps)[];

function stripUndefined<T extends object>(value: T): Partial<T> {
  const result: Partial<T> = {};
  for (const key of Object.keys(value) as Array<keyof T>) {
    if (value[key] !== undefined) result[key] = value[key];
  }
  return result;
}

/** The legacy `onFormat` callback, whose currency and percent buttons send shorthand actions. */
const LegacyFormatContext = createContext<ToolbarProps['onFormat'] | null>(null);

/** A number-format shortcut button; legacy toolbars send `'currency'` or `'percent'`. */
function NumberShortcut({ value }: { value: 'currency' | 'percent' }) {
  const legacyFormat = useContext(LegacyFormatContext);
  const command = useXlsxCommand('numberFormat', { value });
  const { t } = useTranslation();
  if (!legacyFormat) return <ToolbarCommand id="numberFormat" args={{ value }} />;
  const state = command.state;
  return (
    <ToolbarButtonBase
      title={command.label}
      disabled={!state.enabled}
      description={state.enabled ? undefined : state.disabledReason.message}
      onClick={() => legacyFormat(value)}
    >
      <span style={{ fontSize: value === 'currency' ? 16 : 15 }}>
        {value === 'currency' ? t('toolbar.currencySymbol') : '%'}
      </span>
    </ToolbarButtonBase>
  );
}

/** The built-in controls in their default order. */
function DefaultToolbarItems() {
  const { t } = useTranslation();
  return (
    <>
      <ToolbarGroup label={t('toolbar.groups.search')}>
        <ToolbarCommand id="searchMenus" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.history')}>
        <ToolbarCommand id="undo" />
        <ToolbarCommand id="redo" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.print')}>
        <ToolbarCommand id="print" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.paintFormat')}>
        <ToolbarCommand id="paintFormat" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.zoom')}>
        <ToolbarCommand id="zoom" />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.number')}>
        <NumberShortcut value="currency" />
        <NumberShortcut value="percent" />
        <ToolbarCommand id="decimalPlaces" args={{ direction: 'decrease' }} />
        <ToolbarCommand id="decimalPlaces" args={{ direction: 'increase' }} />
        <ToolbarCommand id="numberFormat" />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.font')}>
        <ToolbarCommand id="fontFamily" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.font')}>
        <ToolbarCommand id="fontSizeStep" args={{ direction: 'decrease' }} />
        <ToolbarCommand id="fontSize" />
        <ToolbarCommand id="fontSizeStep" args={{ direction: 'increase' }} />
      </ToolbarGroup>
      <ToolbarSeparator />
      <ToolbarGroup label={t('toolbar.groups.text')}>
        <ToolbarCommand id="bold" />
        <ToolbarCommand id="italic" />
        <ToolbarCommand id="strikethrough" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.colors')}>
        <ToolbarCommand id="textColor" />
        <ToolbarCommand id="fillColor" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.borders')}>
        <ToolbarCommand id="borderPreset" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.merge')}>
        <ToolbarCommand id="merge" />
      </ToolbarGroup>
      <ToolbarGroup label={t('toolbar.groups.alignment')}>
        <ToolbarCommand id="horizontalAlignment" />
        <ToolbarCommand id="verticalAlignment" />
        <ToolbarCommand id="textWrapping" />
      </ToolbarGroup>
      <XlsxPluginToolbar />
    </>
  );
}

/** Rejects legacy state and callbacks where commands are the only authority. */
export function assertCommandProps(props: ToolbarProps, component: string): void {
  const conflicting = LEGACY_PROPS.filter((key) => props[key] !== undefined);
  if (conflicting.length === 0) return;
  throw new Error(
    `A commands-mode ${component} reads state and actions from the editor's commands; remove ${conflicting.join(', ')} or use mode="legacy".`
  );
}

function CommandToolbar(props: ToolbarProps) {
  assertCommandProps(props, 'Toolbar');
  return (
    <ToolbarRail className={props.className} style={props.style}>
      {props.children ?? <DefaultToolbarItems />}
    </ToolbarRail>
  );
}

function LegacyToolbar(explicitProps: ToolbarProps) {
  const { t } = useTranslation();
  const context = useContext(EditorToolbarContext);
  const props: ToolbarProps = context ? { ...context, ...stripUndefined(explicitProps) } : explicitProps;
  const latest = useRef(props);
  latest.current = props;
  const key = JSON.stringify([
    props.currentFormatting,
    props.selectionShape,
    props.canUndo,
    props.canRedo,
    props.zoom,
    props.fontFamilies,
    props.fontSizes,
    props.disabled,
    LEGACY_PROPS.filter((name) => name.startsWith('on') && props[name] !== undefined),
  ]);
  const store = useMemo(() => createLegacyToolbarStore(props, () => latest.current, t), [key, t]);
  const format = props.onFormat ? (action: FormattingAction) => latest.current.onFormat?.(action) : null;
  return (
    <XlsxCommandContext.Provider value={store}>
      <LegacyFormatContext.Provider value={format}>
        <ToolbarRail className={props.className} style={props.style}>
          <DefaultToolbarItems />
          {props.children != null && <ToolbarLegacyContent>{props.children}</ToolbarLegacyContent>}
        </ToolbarRail>
      </LegacyFormatContext.Provider>
    </XlsxCommandContext.Provider>
  );
}

/**
 * The formatting rail. In legacy mode it renders the default controls from
 * props and appends `children`; in commands mode, `children` replace the
 * default arrangement. Narrow rails move trailing groups into a "More" menu.
 */
export function Toolbar(props: ToolbarProps) {
  const inherited = useContext(ToolbarModeContext);
  const mode = props.mode ?? inherited ?? 'legacy';
  return mode === 'commands' ? <CommandToolbar {...props} /> : <LegacyToolbar {...props} />;
}

export { Toolbar as XlsxToolbar };
