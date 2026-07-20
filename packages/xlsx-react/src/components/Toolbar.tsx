import { useContext, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { useTranslation } from '../i18n';
import { EditorToolbarContext } from './EditorToolbarContext';
import { ColorPicker } from './ui/ColorPicker';
import { EditableCombobox } from './ui/EditableCombobox';
import { ToolbarIcon } from './ui/ToolbarIcon';
import {
  ToolbarButton,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  ToolbarMenuSeparator,
  ToolbarSeparator,
  toolbarColors,
} from './ui/ToolbarPrimitives';

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
  children?: ReactNode;
}

interface ToolbarSection {
  key: string;
  width: number;
  node: ReactNode;
}

const ZOOM_LEVELS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2] as const;
const DEFAULT_FONT_FAMILIES = [
  'Arial',
  'Calibri',
  'Cambria',
  'Georgia',
  'Roboto',
  'Times New Roman',
  'Verdana',
] as const;
const DEFAULT_FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72] as const;
const NUMBER_FORMATS: readonly NumberFormat[] = [
  'automatic',
  'plainText',
  'number',
  'percent',
  'scientific',
  'currency',
  'date',
  'time',
  'custom',
];
const BORDER_PRESETS: readonly BorderPreset[] = [
  'all',
  'inner',
  'horizontal',
  'vertical',
  'outer',
  'left',
  'top',
  'right',
  'bottom',
  'none',
];
const BORDER_STYLES: readonly BorderStyle[] = ['solid', 'dashed', 'dotted', 'double'];
const HORIZONTAL_ALIGNMENTS: readonly HorizontalAlignment[] = ['left', 'center', 'right'];
const VERTICAL_ALIGNMENTS: readonly VerticalAlignment[] = ['top', 'middle', 'bottom'];
const WRAPPING_OPTIONS: readonly TextWrapping[] = ['overflow', 'wrap', 'clip'];

function stripUndefined<T extends object>(value: T): Partial<T> {
  const result: Partial<T> = {};
  for (const key of Object.keys(value) as Array<keyof T>) {
    if (value[key] !== undefined) result[key] = value[key];
  }
  return result;
}

function useToolbarProps(props: ToolbarProps): ToolbarProps {
  const context = useContext(EditorToolbarContext);
  return context ? { ...context, ...stripUndefined(props) } : props;
}

function BorderGlyph({ preset }: { preset: BorderPreset }) {
  const edge = (side: BorderPreset) =>
    preset === side || preset === 'all' || preset === 'outer' ? 2 : 0.6;
  return (
    <svg
      width="20"
      height="20"
      viewBox="0 0 20 20"
      fill="none"
      stroke="currentColor"
      aria-hidden="true"
    >
      {preset !== 'none' && (
        <>
          <path d="M3 3h14" strokeWidth={edge('top')} />
          <path d="M17 3v14" strokeWidth={edge('right')} />
          <path d="M3 17h14" strokeWidth={edge('bottom')} />
          <path d="M3 3v14" strokeWidth={edge('left')} />
          {(preset === 'all' || preset === 'inner' || preset === 'horizontal') && (
            <path d="M3 10h14" strokeWidth="1.5" />
          )}
          {(preset === 'all' || preset === 'inner' || preset === 'vertical') && (
            <path d="M10 3v14" strokeWidth="1.5" />
          )}
        </>
      )}
      {preset === 'none' && <path d="M4 4h12v12H4zM3 17 17 3" strokeWidth="1.5" />}
    </svg>
  );
}

function HorizontalAlignmentGlyph({ value }: { value: HorizontalAlignment }) {
  const x1 = value === 'left' ? 3 : value === 'center' ? 5 : 7;
  const x2 = value === 'left' ? 15 : value === 'center' ? 17 : 19;
  return (
    <svg
      width="20"
      height="20"
      viewBox="0 0 20 20"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      aria-hidden="true"
    >
      <path d="M3 4h14M3 10h14M3 16h14" />
      <path d={`M${x1} 7h${x2 - x1}M${x1} 13h${x2 - x1}`} />
    </svg>
  );
}

function VerticalAlignmentGlyph({ value }: { value: VerticalAlignment }) {
  const y = value === 'top' ? 5 : value === 'middle' ? 10 : 15;
  return (
    <svg
      width="20"
      height="20"
      viewBox="0 0 20 20"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      aria-hidden="true"
    >
      <path d="M3 3h14M3 17h14" />
      <path d={`M6 ${y}h8`} />
      <path
        d={
          value === 'top'
            ? 'm8 8 2-3 2 3'
            : value === 'bottom'
            ? 'm8 12 2 3 2-3'
            : 'm8 7 2 3 2-3m-4 6 2-3 2 3'
        }
      />
    </svg>
  );
}

function BorderStyleGlyph({ value }: { value: BorderStyle }) {
  return (
    <svg
      width="44"
      height="12"
      viewBox="0 0 44 12"
      fill="none"
      stroke="currentColor"
      aria-hidden="true"
    >
      {value === 'double' ? (
        <>
          <path d="M2 4h40M2 8h40" />
        </>
      ) : (
        <path
          d="M2 6h40"
          strokeWidth="2"
          strokeDasharray={value === 'dashed' ? '7 4' : value === 'dotted' ? '1 4' : undefined}
        />
      )}
    </svg>
  );
}

function nextFontSize(value: number, sizes: readonly number[], direction: -1 | 1): number {
  if (direction > 0) return sizes.find((size) => size > value) ?? value + 1;
  return [...sizes].reverse().find((size) => size < value) ?? Math.max(1, value - 1);
}

export function Toolbar(explicitProps: ToolbarProps) {
  const { t } = useTranslation();
  const {
    currentFormatting = {},
    selectionShape,
    onFormat,
    onMerge,
    onSearchMenus,
    onUndo,
    onRedo,
    canUndo = false,
    canRedo = false,
    onPrint,
    zoom = 1,
    onZoomChange,
    fontFamilies = DEFAULT_FONT_FAMILIES,
    fontSizes = DEFAULT_FONT_SIZES,
    disabled = false,
    className,
    style,
    children,
  } = useToolbarProps(explicitProps);
  const rootRef = useRef<HTMLDivElement>(null);
  const [rootWidth, setRootWidth] = useState(Number.POSITIVE_INFINITY);
  const formattingEnabled = !disabled && Boolean(onFormat);
  const mergeEnabled = !disabled && Boolean(onMerge);
  const fontSize = currentFormatting.fontSize ?? 10;
  const hasRows = (selectionShape?.rows ?? 1) > 1;
  const hasColumns = (selectionShape?.columns ?? 1) > 1;
  const canMergeAll = hasRows || hasColumns;

  useEffect(() => {
    const root = rootRef.current;
    if (!root) return;
    const update = () => setRootWidth(root.clientWidth);
    update();
    const observer = new ResizeObserver(update);
    observer.observe(root);
    return () => observer.disconnect();
  }, []);

  const zoomOptions = useMemo(
    () =>
      ZOOM_LEVELS.map((level) => ({
        value: `${level * 100}%`,
        label: `${level * 100}%`,
      })),
    []
  );
  const fontSizeOptions = useMemo(
    () => fontSizes.map((size) => ({ value: String(size), label: String(size) })),
    [fontSizes]
  );

  const apply = (action: FormattingAction) => {
    if (formattingEnabled) onFormat?.(action);
  };

  const renderNumberFormats = (close: () => void) => (
    <>
      {NUMBER_FORMATS.map((format) => (
        <ToolbarMenuItem
          key={format}
          label={t(`toolbar.numberFormats.${format}`)}
          selected={currentFormatting.numberFormat === format}
          disabled={!formattingEnabled}
          onClick={() => apply({ type: 'numberFormat', value: format })}
          close={close}
        />
      ))}
    </>
  );

  const renderBorders = (close: () => void) => (
    <>
      {BORDER_PRESETS.map((preset) => (
        <ToolbarMenuItem
          key={preset}
          icon={<BorderGlyph preset={preset} />}
          label={t(`toolbar.borderPresets.${preset}`)}
          selected={currentFormatting.borderPreset === preset}
          disabled={!formattingEnabled}
          onClick={() => apply({ type: 'borderPreset', value: preset })}
          close={close}
        />
      ))}
      <ToolbarMenuSeparator />
      {BORDER_STYLES.map((borderStyle) => (
        <ToolbarMenuItem
          key={borderStyle}
          icon={<BorderStyleGlyph value={borderStyle} />}
          label={t(`toolbar.borderStyles.${borderStyle}`)}
          selected={currentFormatting.borderStyle === borderStyle}
          disabled={!formattingEnabled}
          onClick={() => apply({ type: 'borderStyle', value: borderStyle })}
          close={close}
        />
      ))}
      <ToolbarMenuSeparator />
      <label
        style={{
          position: 'relative',
          display: 'flex',
          alignItems: 'center',
          gap: 10,
          minHeight: 32,
          padding: '5px 9px',
          color: toolbarColors.text,
          font: '400 13px ui-sans-serif, system-ui, sans-serif',
          cursor: formattingEnabled ? 'pointer' : 'default',
        }}
      >
        <span
          style={{
            width: 20,
            height: 5,
            borderRadius: 2,
            background: currentFormatting.borderColor ?? '#000000',
          }}
        />
        <span>{t('toolbar.borderColor')}</span>
        <input
          type="color"
          value={currentFormatting.borderColor ?? '#000000'}
          disabled={!formattingEnabled}
          aria-label={t('toolbar.borderColor')}
          onChange={(event) => apply({ type: 'borderColor', value: event.target.value })}
          style={{
            position: 'absolute',
            inset: 0,
            width: '100%',
            height: '100%',
            opacity: 0,
          }}
        />
      </label>
    </>
  );

  const renderMerge = (close: () => void) => (
    <>
      <ToolbarMenuItem
        label={t('toolbar.merge.all')}
        disabled={!mergeEnabled || !canMergeAll}
        onClick={() => onMerge?.('all')}
        close={close}
      />
      <ToolbarMenuItem
        label={t('toolbar.merge.horizontal')}
        disabled={!mergeEnabled || !hasColumns}
        onClick={() => onMerge?.('horizontal')}
        close={close}
      />
      <ToolbarMenuItem
        label={t('toolbar.merge.vertical')}
        disabled={!mergeEnabled || !hasRows}
        onClick={() => onMerge?.('vertical')}
        close={close}
      />
      <ToolbarMenuSeparator />
      <ToolbarMenuItem
        label={t('toolbar.merge.unmerge')}
        disabled={!mergeEnabled || !selectionShape?.canUnmerge}
        onClick={() => onMerge?.('unmerge')}
        close={close}
      />
    </>
  );

  const renderHorizontalAlignment = (close: () => void) => (
    <>
      {HORIZONTAL_ALIGNMENTS.map((alignment) => (
        <ToolbarMenuItem
          key={alignment}
          icon={<HorizontalAlignmentGlyph value={alignment} />}
          label={t(`toolbar.horizontalAlign.${alignment}`)}
          selected={currentFormatting.horizontalAlignment === alignment}
          disabled={!formattingEnabled}
          onClick={() => apply({ type: 'horizontalAlignment', value: alignment })}
          close={close}
        />
      ))}
    </>
  );

  const renderVerticalAlignment = (close: () => void) => (
    <>
      {VERTICAL_ALIGNMENTS.map((alignment) => (
        <ToolbarMenuItem
          key={alignment}
          icon={<VerticalAlignmentGlyph value={alignment} />}
          label={t(`toolbar.verticalAlign.${alignment}`)}
          selected={currentFormatting.verticalAlignment === alignment}
          disabled={!formattingEnabled}
          onClick={() => apply({ type: 'verticalAlignment', value: alignment })}
          close={close}
        />
      ))}
    </>
  );

  const renderWrapping = (close: () => void) => (
    <>
      {WRAPPING_OPTIONS.map((wrapping) => (
        <ToolbarMenuItem
          key={wrapping}
          label={t(`toolbar.wrapping.${wrapping}`)}
          selected={currentFormatting.textWrapping === wrapping}
          disabled={!formattingEnabled}
          onClick={() => apply({ type: 'textWrapping', value: wrapping })}
          close={close}
        />
      ))}
    </>
  );

  const sections: ToolbarSection[] = [
    {
      key: 'search',
      width: 29,
      node: (
        <ToolbarGroup label={t('toolbar.groups.search')}>
          <ToolbarButton
            title={t('toolbar.searchMenus')}
            onClick={onSearchMenus}
            testId="xlsx-search-menus"
          >
            <ToolbarIcon name="search" />
          </ToolbarButton>
        </ToolbarGroup>
      ),
    },
    {
      key: 'history',
      width: 58,
      node: (
        <ToolbarGroup label={t('toolbar.groups.history')}>
          <ToolbarButton
            title={t('toolbar.undo')}
            disabled={disabled || !canUndo || !onUndo}
            onClick={onUndo}
            testId="xlsx-undo"
          >
            <ToolbarIcon name="undo" />
          </ToolbarButton>
          <ToolbarButton
            title={t('toolbar.redo')}
            disabled={disabled || !canRedo || !onRedo}
            onClick={onRedo}
            testId="xlsx-redo"
          >
            <ToolbarIcon name="redo" />
          </ToolbarButton>
        </ToolbarGroup>
      ),
    },
    {
      key: 'print',
      width: 29,
      node: (
        <ToolbarGroup label={t('toolbar.groups.print')}>
          <ToolbarButton
            title={t('toolbar.print')}
            disabled={disabled || !onPrint}
            onClick={onPrint}
          >
            <ToolbarIcon name="print" />
          </ToolbarButton>
        </ToolbarGroup>
      ),
    },
    {
      key: 'paint',
      width: 29,
      node: (
        <ToolbarGroup label={t('toolbar.groups.paintFormat')}>
          <ToolbarButton
            title={t('toolbar.paintFormat')}
            active={currentFormatting.paintFormat}
            disabled={!formattingEnabled}
            onClick={() => apply('paintFormat')}
          >
            <ToolbarIcon name="formatPaint" />
          </ToolbarButton>
        </ToolbarGroup>
      ),
    },
    {
      key: 'zoom',
      width: 82,
      node: (
        <ToolbarGroup label={t('toolbar.groups.zoom')}>
          <EditableCombobox
            value={`${Math.round(zoom * 100)}%`}
            options={zoomOptions}
            label={t('toolbar.zoomValue', {
              value: `${Math.round(zoom * 100)}%`,
            })}
            disabled={disabled || !onZoomChange}
            onCommit={(value) => {
              const percent = Number.parseFloat(value.replace('%', ''));
              if (Number.isFinite(percent) && percent >= 25 && percent <= 400)
                onZoomChange?.(percent / 100);
            }}
            width={76}
            testId="xlsx-zoom"
          />
        </ToolbarGroup>
      ),
    },
    {
      key: 'number',
      width: 202,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.number')}>
            <ToolbarButton
              title={t('toolbar.currency')}
              disabled={!formattingEnabled}
              onClick={() => apply('currency')}
            >
              <span style={{ fontSize: 16 }}>{t('toolbar.currencySymbol')}</span>
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.percent')}
              disabled={!formattingEnabled}
              onClick={() => apply('percent')}
            >
              <span style={{ fontSize: 15 }}>%</span>
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.decreaseDecimal')}
              disabled={!formattingEnabled}
              onClick={() => apply('decreaseDecimal')}
            >
              <ToolbarIcon name="decimalDecrease" />
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.increaseDecimal')}
              disabled={!formattingEnabled}
              onClick={() => apply('increaseDecimal')}
            >
              <ToolbarIcon name="decimalIncrease" />
            </ToolbarButton>
            <ToolbarDropdown
              title={t('toolbar.moreNumberFormats')}
              disabled={!formattingEnabled}
              trigger={
                <>
                  <span>123</span>
                  <ToolbarIcon name="chevronDown" size={13} />
                </>
              }
            >
              {renderNumberFormats}
            </ToolbarDropdown>
          </ToolbarGroup>
        </>
      ),
    },
    {
      key: 'font-family',
      width: 137,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.font')}>
            <ToolbarDropdown
              title={t('toolbar.fontFamily')}
              disabled={!formattingEnabled}
              menuWidth={210}
              style={{ width: 120, justifyContent: 'space-between' }}
              trigger={
                <>
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
                    {currentFormatting.fontFamily ?? 'Calibri'}
                  </span>
                  <ToolbarIcon name="chevronDown" size={13} />
                </>
              }
            >
              {(close) => (
                <>
                  {fontFamilies.map((font) => (
                    <ToolbarMenuItem
                      key={font}
                      label={font}
                      selected={currentFormatting.fontFamily === font}
                      disabled={!formattingEnabled}
                      onClick={() => apply({ type: 'fontFamily', value: font })}
                      close={close}
                    />
                  ))}
                </>
              )}
            </ToolbarDropdown>
          </ToolbarGroup>
        </>
      ),
    },
    {
      key: 'font-size',
      width: 116,
      node: (
        <ToolbarGroup label={t('toolbar.groups.font')}>
          <ToolbarButton
            title={t('toolbar.decreaseFontSize')}
            disabled={!formattingEnabled}
            onClick={() =>
              apply({
                type: 'fontSize',
                value: nextFontSize(fontSize, fontSizes, -1),
              })
            }
          >
            <ToolbarIcon name="remove" />
          </ToolbarButton>
          <EditableCombobox
            value={String(fontSize)}
            options={fontSizeOptions}
            label={t('toolbar.fontSize')}
            disabled={!formattingEnabled}
            onCommit={(value) => {
              const size = Number.parseFloat(value);
              if (Number.isFinite(size) && size >= 1 && size <= 400)
                apply({ type: 'fontSize', value: size });
            }}
            width={50}
            inputStyle={{ textAlign: 'center' }}
          />
          <ToolbarButton
            title={t('toolbar.increaseFontSize')}
            disabled={!formattingEnabled}
            onClick={() =>
              apply({
                type: 'fontSize',
                value: nextFontSize(fontSize, fontSizes, 1),
              })
            }
          >
            <ToolbarIcon name="add" />
          </ToolbarButton>
        </ToolbarGroup>
      ),
    },
    {
      key: 'text',
      width: 98,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.text')}>
            <ToolbarButton
              title={t('toolbar.bold')}
              active={currentFormatting.bold}
              disabled={!formattingEnabled}
              onClick={() => apply('bold')}
            >
              <ToolbarIcon name="bold" />
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.italic')}
              active={currentFormatting.italic}
              disabled={!formattingEnabled}
              onClick={() => apply('italic')}
            >
              <ToolbarIcon name="italic" />
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.strikethrough')}
              active={currentFormatting.strikethrough}
              disabled={!formattingEnabled}
              onClick={() => apply('strikethrough')}
            >
              <ToolbarIcon name="strikethrough" />
            </ToolbarButton>
          </ToolbarGroup>
        </>
      ),
    },
    {
      key: 'colors',
      width: 58,
      node: (
        <ToolbarGroup label={t('toolbar.groups.colors')}>
          <ColorPicker
            mode="text"
            value={currentFormatting.textColor ?? '#000000'}
            label={t('toolbar.textColor')}
            disabled={!formattingEnabled}
            onChange={(value) => apply({ type: 'textColor', value })}
          />
          <ColorPicker
            mode="fill"
            value={currentFormatting.fillColor ?? '#ffffff'}
            label={t('toolbar.fillColor')}
            disabled={!formattingEnabled}
            onChange={(value) => apply({ type: 'fillColor', value })}
          />
        </ToolbarGroup>
      ),
    },
    {
      key: 'borders',
      width: 36,
      node: (
        <ToolbarGroup label={t('toolbar.groups.borders')}>
          <ToolbarDropdown
            title={t('toolbar.borders')}
            disabled={!formattingEnabled}
            menuWidth={240}
            trigger={
              <>
                <BorderGlyph preset={currentFormatting.borderPreset ?? 'all'} />
                <ToolbarIcon name="chevronDown" size={12} />
              </>
            }
          >
            {renderBorders}
          </ToolbarDropdown>
        </ToolbarGroup>
      ),
    },
    {
      key: 'merge',
      width: 58,
      node: (
        <ToolbarGroup label={t('toolbar.groups.merge')}>
          <ToolbarButton
            title={t('toolbar.merge.all')}
            disabled={!mergeEnabled || !canMergeAll}
            onClick={() => onMerge?.('all')}
            style={{ borderRadius: '4px 0 0 4px' }}
            testId="xlsx-merge-all"
          >
            <ToolbarIcon name="merge" />
          </ToolbarButton>
          <ToolbarDropdown
            title={t('toolbar.mergeCells')}
            disabled={!mergeEnabled || (!canMergeAll && !selectionShape?.canUnmerge)}
            menuWidth={220}
            style={{
              minWidth: 20,
              width: 20,
              padding: 0,
              borderRadius: '0 4px 4px 0',
            }}
            trigger={<ToolbarIcon name="chevronDown" size={13} />}
          >
            {renderMerge}
          </ToolbarDropdown>
        </ToolbarGroup>
      ),
    },
    {
      key: 'alignment',
      width: 118,
      node: (
        <ToolbarGroup label={t('toolbar.groups.alignment')}>
          <ToolbarDropdown
            title={t('toolbar.horizontalAlignment')}
            disabled={!formattingEnabled}
            menuWidth={180}
            trigger={
              <>
                <HorizontalAlignmentGlyph value={currentFormatting.horizontalAlignment ?? 'left'} />
                <ToolbarIcon name="chevronDown" size={12} />
              </>
            }
          >
            {renderHorizontalAlignment}
          </ToolbarDropdown>
          <ToolbarDropdown
            title={t('toolbar.verticalAlignment')}
            disabled={!formattingEnabled}
            menuWidth={180}
            trigger={
              <>
                <VerticalAlignmentGlyph value={currentFormatting.verticalAlignment ?? 'middle'} />
                <ToolbarIcon name="chevronDown" size={12} />
              </>
            }
          >
            {renderVerticalAlignment}
          </ToolbarDropdown>
          <ToolbarDropdown
            title={t('toolbar.textWrapping')}
            disabled={!formattingEnabled}
            menuWidth={180}
            trigger={
              <>
                <ToolbarIcon name="wrap" />
                <ToolbarIcon name="chevronDown" size={12} />
              </>
            }
          >
            {renderWrapping}
          </ToolbarDropdown>
        </ToolbarGroup>
      ),
    },
  ];

  if (children) sections.push({ key: 'custom', width: 40, node: children });

  const availableWidth = Math.max(0, rootWidth - 48);
  let usedWidth = 0;
  let visibleCount = sections.length;
  for (let index = 0; index < sections.length; index++) {
    usedWidth += sections[index].width;
    if (usedWidth > availableWidth) {
      visibleCount = index;
      break;
    }
  }
  const visibleSections = sections.slice(0, visibleCount);
  const overflowSections = sections.slice(visibleCount);

  return (
    <div
      ref={rootRef}
      className={className}
      role="toolbar"
      aria-label={t('toolbar.actionsLabel')}
      data-testid="xlsx-formatting-toolbar"
      style={{
        display: 'flex',
        alignItems: 'center',
        minWidth: 0,
        minHeight: 36,
        margin: '0 8px 5px',
        padding: '4px 7px',
        borderRadius: 18,
        background: toolbarColors.rail,
        color: toolbarColors.text,
        overflow: 'hidden',
        boxSizing: 'border-box',
        ...style,
      }}
    >
      {visibleSections.map((section) => (
        <span key={section.key} style={{ display: 'contents' }}>
          {section.node}
        </span>
      ))}
      <span style={{ marginLeft: 'auto', flex: '0 0 auto' }}>
        <ToolbarDropdown
          title={t('toolbar.more')}
          disabled={overflowSections.length === 0}
          menuWidth={390}
          testId="xlsx-toolbar-more"
          trigger={<ToolbarIcon name="more" />}
        >
          {() => (
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                flexWrap: 'wrap',
                gap: 2,
              }}
            >
              {overflowSections.map((section) => (
                <span key={section.key} style={{ display: 'contents' }}>
                  {section.node}
                </span>
              ))}
            </div>
          )}
        </ToolbarDropdown>
      </span>
    </div>
  );
}

export { Toolbar as XlsxToolbar };
