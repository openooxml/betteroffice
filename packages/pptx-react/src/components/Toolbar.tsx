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
  ToolbarSeparator,
  toolbarColors,
} from './ui/ToolbarPrimitives';

export type PptxEditorTool = 'select' | 'textBox';
export type PptxZoom = number | 'fit';

export interface SelectionFormatting {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  textColor?: string;
}

export type FormattingAction =
  | 'bold'
  | 'italic'
  | 'underline'
  | { type: 'fontFamily'; value: string }
  | { type: 'fontSize'; value: number }
  | { type: 'textColor'; value: string };

export interface SlideLayoutOption {
  partPath: string | null;
  label?: string;
}

export interface ToolbarProps {
  currentFormatting?: SelectionFormatting;
  textSelectionActive?: boolean;
  onFormat?: (action: FormattingAction) => void;
  onInsertSlide?: (layoutPartPath?: string | null) => void;
  slideLayouts?: readonly SlideLayoutOption[];
  currentLayoutPartPath?: string | null;
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
  children?: ReactNode;
}

interface ToolbarSection {
  key: string;
  width: number;
  node: ReactNode;
}

const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2] as const;
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

function nextFontSize(value: number, sizes: readonly number[], direction: -1 | 1): number {
  if (direction > 0) return sizes.find((size) => size > value) ?? value + 1;
  return [...sizes].reverse().find((size) => size < value) ?? Math.max(1, value - 1);
}

export function Toolbar(explicitProps: ToolbarProps) {
  const { t } = useTranslation();
  const {
    currentFormatting = {},
    textSelectionActive = false,
    onFormat,
    onInsertSlide,
    slideLayouts = [],
    currentLayoutPartPath,
    onUndo,
    onRedo,
    canUndo = false,
    canRedo = false,
    zoom = 'fit',
    onZoomChange,
    activeTool = 'select',
    onToolChange,
    fontFamilies = DEFAULT_FONT_FAMILIES,
    fontSizes = DEFAULT_FONT_SIZES,
    disabled = false,
    className,
    style,
    children,
  } = useToolbarProps(explicitProps);
  const rootRef = useRef<HTMLDivElement>(null);
  const [rootWidth, setRootWidth] = useState(Number.POSITIVE_INFINITY);
  const formattingEnabled = !disabled && textSelectionActive && Boolean(onFormat);
  const slideEnabled = !disabled && Boolean(onInsertSlide);
  const toolEnabled = !disabled && Boolean(onToolChange);
  const fontSize = currentFormatting.fontSize ?? 24;
  const fitLabel = t('toolbar.fit');
  const zoomValue = zoom === 'fit' ? fitLabel : `${Math.round(zoom * 100)}%`;

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
    () => [
      { value: fitLabel, label: fitLabel },
      ...ZOOM_LEVELS.map((level) => ({
        value: `${level * 100}%`,
        label: `${level * 100}%`,
      })),
    ],
    [fitLabel]
  );
  const fontSizeOptions = useMemo(
    () => fontSizes.map((size) => ({ value: String(size), label: String(size) })),
    [fontSizes]
  );

  const apply = (action: FormattingAction) => {
    if (formattingEnabled) onFormat?.(action);
  };

  const sections: ToolbarSection[] = [
    {
      key: 'new-slide',
      width: 59,
      node: (
        <ToolbarGroup label={t('toolbar.groups.slides')}>
          <ToolbarButton
            title={t('toolbar.newSlide')}
            disabled={!slideEnabled}
            onClick={() => onInsertSlide?.()}
            style={{ borderRadius: '4px 0 0 4px' }}
            testId="pptx-new-slide"
          >
            <ToolbarIcon name="newSlide" />
          </ToolbarButton>
          <ToolbarDropdown
            title={t('toolbar.newSlideWithLayout')}
            disabled={!slideEnabled || slideLayouts.length === 0}
            menuWidth={230}
            testId="pptx-new-slide-layout"
            style={{
              minWidth: 20,
              width: 20,
              padding: 0,
              borderRadius: '0 4px 4px 0',
            }}
            trigger={<ToolbarIcon name="chevronDown" size={13} />}
          >
            {(close) => (
              <>
                {slideLayouts.map((layout, index) => (
                  <ToolbarMenuItem
                    key={layout.partPath ?? `default-${index}`}
                    label={layout.label ?? t('toolbar.layoutOption', { number: index + 1 })}
                    selected={(layout.partPath ?? null) === (currentLayoutPartPath ?? null)}
                    onClick={() => onInsertSlide?.(layout.partPath)}
                    close={close}
                  />
                ))}
              </>
            )}
          </ToolbarDropdown>
        </ToolbarGroup>
      ),
    },
    {
      key: 'history',
      width: 75,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.history')}>
            <ToolbarButton
              title={t('toolbar.undoShortcut')}
              disabled={disabled || !canUndo || !onUndo}
              onClick={onUndo}
              testId="pptx-undo"
            >
              <ToolbarIcon name="undo" />
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.redoShortcut')}
              disabled={disabled || !canRedo || !onRedo}
              onClick={onRedo}
              testId="pptx-redo"
            >
              <ToolbarIcon name="redo" />
            </ToolbarButton>
          </ToolbarGroup>
        </>
      ),
    },
    {
      key: 'zoom',
      width: 82,
      node: (
        <ToolbarGroup label={t('toolbar.groups.zoom')}>
          <EditableCombobox
            value={zoomValue}
            options={zoomOptions}
            label={t('toolbar.zoomValue', { value: zoomValue })}
            disabled={disabled || !onZoomChange}
            onCommit={(value) => {
              if (value === fitLabel) {
                onZoomChange?.('fit');
                return;
              }
              const percent = Number.parseFloat(value.replace('%', ''));
              if (Number.isFinite(percent) && percent >= 25 && percent <= 400) {
                onZoomChange?.(percent / 100);
              }
            }}
            width={76}
            testId="pptx-zoom"
          />
        </ToolbarGroup>
      ),
    },
    {
      key: 'tools',
      width: 75,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.tools')}>
            <ToolbarButton
              title={t('toolbar.selectToolShortcut')}
              active={activeTool === 'select'}
              disabled={!toolEnabled}
              onClick={() => onToolChange?.('select')}
              testId="pptx-tool-select"
            >
              <ToolbarIcon name="select" />
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.textBoxTool')}
              active={activeTool === 'textBox'}
              disabled={!toolEnabled}
              onClick={() => onToolChange?.('textBox')}
              testId="pptx-tool-text-box"
            >
              <ToolbarIcon name="textBox" />
            </ToolbarButton>
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
              testId="pptx-font-family"
              style={{ width: 120, justifyContent: 'space-between' }}
              trigger={
                <>
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
                    {currentFormatting.fontFamily ?? t('toolbar.mixed')}
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
            value={currentFormatting.fontSize === undefined ? '' : String(fontSize)}
            options={fontSizeOptions}
            label={t('toolbar.fontSize')}
            disabled={!formattingEnabled}
            onCommit={(value) => {
              const size = Number.parseFloat(value);
              if (Number.isFinite(size) && size >= 1 && size <= 400) {
                apply({ type: 'fontSize', value: size });
              }
            }}
            width={50}
            inputStyle={{ textAlign: 'center' }}
            testId="pptx-font-size"
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
      width: 133,
      node: (
        <>
          <ToolbarSeparator />
          <ToolbarGroup label={t('toolbar.groups.text')}>
            <ToolbarButton
              title={t('toolbar.boldShortcut')}
              active={currentFormatting.bold}
              disabled={!formattingEnabled}
              onClick={() => apply('bold')}
              testId="pptx-bold"
            >
              <ToolbarIcon name="bold" />
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.italicShortcut')}
              active={currentFormatting.italic}
              disabled={!formattingEnabled}
              onClick={() => apply('italic')}
              testId="pptx-italic"
            >
              <ToolbarIcon name="italic" />
            </ToolbarButton>
            <ToolbarButton
              title={t('toolbar.underlineShortcut')}
              active={currentFormatting.underline}
              disabled={!formattingEnabled}
              onClick={() => apply('underline')}
              testId="pptx-underline"
            >
              <ToolbarIcon name="underline" />
            </ToolbarButton>
            <ColorPicker
              value={currentFormatting.textColor ?? '#000000'}
              label={t('toolbar.textColor')}
              disabled={!formattingEnabled}
              onChange={(value) => apply({ type: 'textColor', value })}
              testId="pptx-text-color"
            />
          </ToolbarGroup>
        </>
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
      aria-label={t('toolbar.label')}
      data-testid="pptx-formatting-toolbar"
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
          testId="pptx-toolbar-more"
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

export { Toolbar as PptxToolbar };
