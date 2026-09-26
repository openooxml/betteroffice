import { useCallback, useContext, useRef, useState } from 'react';
import type { ComponentType, ReactNode } from 'react';
import type { TFunction, TranslationKey } from '@betteroffice/docx-i18n';
import { toolbarValueToLayoutTarget } from '@betteroffice/docx/docx';
import type { ColorValue, Theme } from '@betteroffice/docx/types/document';
import {
  generateThemeTintShadeMatrix,
  resolveColorToHex,
  resolveHighlightColor,
  type ThemeMatrixCell,
} from '@betteroffice/docx/utils';
import type { OverflowMenuEntry } from '../../../../../shared/react-toolbar/OverflowMenu';
import {
  docxCommandController,
  type DocxPendingCommand,
} from '../../commands/createDocxCommandStore';
import { commandShortcut } from '../../commands/descriptors';
import {
  useDocxChrome,
  useDocxCommand,
  useDocxCommands,
  useDocxCommandState,
} from '../../commands/hooks';
import type {
  DocxCommandArgs,
  DocxCommandId,
  DocxCommandResult,
  DocxCommandState,
  DocxCommandStore,
  DocxSelectCommandId,
  DocxTableAction,
} from '../../commands/types';
import { useTranslation } from '../../i18n';
import { EditingModeDropdown } from '../DocxEditor/EditingModeDropdown';
import { AlignmentButtons } from '../ui/AlignmentButtons';
import { ColorPicker, STANDARD_COLORS } from '../ui/ColorPicker';
import { FontPicker, type FontOption } from '../ui/FontPicker';
import { FontSizePicker } from '../ui/FontSizePicker';
import { ImageTransformDropdown } from '../ui/ImageTransformDropdown';
import { ImageWrapDropdown } from '../ui/ImageWrapDropdown';
import { LineSpacingPicker } from '../ui/LineSpacingPicker';
import { MaterialSymbol } from '../ui/MaterialSymbol';
import { StylePicker, type StyleOption } from '../ui/StylePicker';
import { TableBorderColorPicker } from '../ui/TableBorderColorPicker';
import { TableBorderPicker } from '../ui/TableBorderPicker';
import { TableBorderWidthPicker } from '../ui/TableBorderWidthPicker';
import { TableCellFillPicker } from '../ui/TableCellFillPicker';
import { TableGridInline } from '../ui/TableGridInline';
import { TableMoreDropdown } from '../ui/TableMoreDropdown';
import { ZoomControl } from '../ui/ZoomControl';
import { useFixedDropdown } from '../../hooks/useFixedDropdown';
import { Button } from '../ui/Button';
import { Tooltip } from '../ui/Tooltip';
import { cn } from '../../lib/utils';
import { restoreEditorFocus, restoreFocusAfterPointer } from './activation';
import { ToolbarOverflowContext, useOverflowSource, type OverflowPrompt } from './overflowRegistry';
import { ToolbarButton } from './ToolbarPrimitives';

const ICON_SIZE = 18;

const COMMAND_ICONS: Partial<Record<DocxCommandId, string>> = {
  undo: 'undo',
  redo: 'redo',
  bold: 'format_bold',
  italic: 'format_italic',
  underline: 'format_underlined',
  strikethrough: 'strikethrough_s',
  superscript: 'superscript',
  subscript: 'subscript',
  clearFormatting: 'format_clear',
  textColor: 'format_color_text',
  highlightColor: 'ink_highlighter',
  alignment: 'format_align_left',
  lineSpacing: 'format_line_spacing',
  bulletList: 'format_list_bulleted',
  numberedList: 'format_list_numbered',
  indent: 'format_indent_increase',
  outdent: 'format_indent_decrease',
  setLtr: 'format_textdirection_l_to_r',
  setRtl: 'format_textdirection_r_to_l',
  insertLink: 'link',
  insertImage: 'image',
  insertTable: 'grid_on',
  insertPageBreak: 'page_break',
  insertSectionBreakNextPage: 'horizontal_rule',
  insertSectionBreakContinuous: 'border_horizontal',
  insertTOC: 'format_list_numbered',
  imageWrap: 'wrap_text',
  imageTransform: 'rotate_right',
  imageProperties: 'tune',
  tableAction: 'table',
  pageSetup: 'settings',
  watermark: 'branding_watermark',
  editingMode: 'edit_note',
  reviewAccept: 'check',
  reviewReject: 'close',
  reviewPrevious: 'keyboard_arrow_up',
  reviewNext: 'keyboard_arrow_down',
  commentsSidebar: 'comment',
  open: 'file_upload',
  save: 'file_download',
  print: 'print',
};

const ALIGNMENT_ICONS: Record<string, string> = {
  left: 'format_align_left',
  center: 'format_align_center',
  right: 'format_align_right',
  both: 'format_align_justify',
};

/**
 * Arguments a control binds; required when the command takes arguments.
 * @experimental
 */
export type ToolbarCommandArgs<K extends DocxCommandId> = null extends DocxCommandArgs[K]
  ? { args?: DocxCommandArgs[K] }
  : { args: DocxCommandArgs[K] };

/** @experimental */
export type ToolbarCommandButtonProps<K extends DocxCommandId> = {
  id: K;
  /** Button content; defaults to the command's icon, or its label when it has none. */
  children?: ReactNode;
  /** Accessible name; defaults to the localized command label. */
  label?: string;
  className?: string;
} & ToolbarCommandArgs<K>;

/** @experimental */
export interface ToolbarCommandSelectProps<K extends DocxSelectCommandId> {
  id: K;
  className?: string;
}

/** @experimental */
export interface ToolbarCommandProps<K extends DocxCommandId> {
  id: K;
  /** Binds arguments; omit to render the command's full built-in control. */
  args?: DocxCommandArgs[K];
  className?: string;
}

function reasonOf(state: DocxCommandState): string | undefined {
  return state.enabled ? undefined : state.disabledReason.message;
}

function iconFor(id: DocxCommandId, args: unknown): string | undefined {
  if (id === 'alignment' && args && typeof args === 'object') {
    return ALIGNMENT_ICONS[(args as DocxCommandArgs['alignment']).value] ?? COMMAND_ICONS[id];
  }
  return COMMAND_ICONS[id];
}

function run<K extends DocxCommandId>(
  store: DocxCommandStore,
  id: K,
  args: DocxCommandArgs[K],
  afterPointer = false
): void {
  void store.execute(id, args).then((result) => {
    if (afterPointer) restoreFocusAfterPointer(store, result);
  });
}

/** Binds a command whose arguments a prompt chooses to the document and selection it opened for. */
function prepare<K extends DocxCommandId>(store: DocxCommandStore, id: K): DocxPendingCommand<K> {
  return (
    docxCommandController(store)?.prepare(id, 'selection') ?? {
      execute: (args) => store.execute(id, args),
    }
  );
}

function commandItem<K extends DocxCommandId>(
  store: DocxCommandStore,
  t: TFunction,
  id: K,
  args?: DocxCommandArgs[K],
  key: string = id
): OverflowMenuEntry {
  const state = store.getState(id, args) as DocxCommandState;
  const icon = iconFor(id, args);
  return {
    kind: 'item',
    id: key,
    label: t(store.getDescriptor(id).labelKey),
    icon: icon ? <MaterialSymbol name={icon} size={16} /> : undefined,
    shortcut: commandShortcut(id, args) ?? undefined,
    checked: state.active,
    disabled: !state.enabled,
    description: reasonOf(state),
    onSelect: () => run(store, id, (args ?? null) as DocxCommandArgs[K]),
  };
}

function optionIsCurrent(id: DocxCommandId, value: unknown, args: unknown): boolean {
  if (id === 'fontFamily') {
    return (
      typeof value === 'string' &&
      value.toLowerCase() === (args as DocxCommandArgs['fontFamily']).family.toLowerCase()
    );
  }
  if (id === 'fontSize') return value === (args as DocxCommandArgs['fontSize']).points;
  return false;
}

/** What overflow presentations need beyond the store: the theme and a value prompt. */
export interface OverflowTools {
  theme: Theme | null;
  prompt?: (request: OverflowPrompt) => void;
}

function selectSubmenu<K extends DocxSelectCommandId>(
  store: DocxCommandStore,
  t: TFunction,
  id: K,
  tools: OverflowTools
): OverflowMenuEntry {
  const state = store.getState(id) as DocxCommandState<K>;
  const reason = reasonOf(state as DocxCommandState);
  const entries: OverflowMenuEntry[] = (state.options ?? []).map((option, index) => {
    const optionState = store.getState(id, option.args) as DocxCommandState;
    return {
      kind: 'item',
      id: `${id}-${index}`,
      label: option.label,
      radio: true,
      checked: optionState.active ?? optionIsCurrent(id, state.value, option.args),
      disabled: !optionState.enabled,
      description: reasonOf(optionState),
      onSelect: () => run(store, id, option.args),
    };
  });
  if (id === 'fontSize' && tools.prompt) {
    const prompt = tools.prompt;
    entries.push({
      kind: 'item',
      id: 'fontSize-custom',
      label: t('commands.customFontSize'),
      disabled: !state.enabled,
      description: reason,
      onSelect: () => {
        const pending = prepare(store, 'fontSize');
        prompt({
          title: t('fontSize.label'),
          label: t('commands.fontSizePoints'),
          initialValue: typeof state.value === 'number' ? String(state.value) : '',
          valid: (value) =>
            /^\d+(\.\d+)?$/.test(value) &&
            store.getState('fontSize', { points: Number(value) }).enabled,
          submit: (value) => void pending.execute({ points: Number(value) }),
        });
      },
    });
  }
  return {
    kind: 'submenu',
    id,
    label: t(store.getDescriptor(id).labelKey),
    disabled: !state.enabled,
    description: reason,
    entries,
  };
}

type ColorChoice = { hex: string; cell?: ThemeMatrixCell } | null;

interface ColorMenu<K extends DocxCommandId> {
  id: string;
  label: string;
  clearLabel: string;
  /** Current color as RGB hex without `#`; null when cleared or unknown. */
  current: string | null;
  state: DocxCommandState;
  command: K;
  /** Arguments of `command` that apply a color, or clear it for null. */
  args(color: ColorChoice): DocxCommandArgs[K];
}

/** The choices of the in-row color picker: clear, theme colors, standard colors, custom. */
function colorSubmenu<K extends DocxCommandId>(
  store: DocxCommandStore,
  t: TFunction,
  menu: ColorMenu<K>,
  tools: OverflowTools
): OverflowMenuEntry {
  const choose = (color: ColorChoice) => run(store, menu.command, menu.args(color));
  const disabled = !menu.state.enabled;
  const description = reasonOf(menu.state);
  const current = menu.current?.replace(/^#/, '').toUpperCase() ?? null;
  const swatch = (hex: string, index: number, label: string, cell?: ThemeMatrixCell): OverflowMenuEntry => ({
    kind: 'item',
    id: `${menu.id}-${cell ? 'theme' : 'standard'}-${index}`,
    label,
    radio: true,
    checked: current === hex.toUpperCase(),
    disabled,
    description,
    onSelect: () => choose({ hex, cell }),
  });
  const matrix = generateThemeTintShadeMatrix(tools.theme?.colorScheme ?? null);
  const columns = matrix[0]?.map((_, column) => matrix.map((row) => row[column])) ?? [];
  const entries: OverflowMenuEntry[] = [
    {
      kind: 'item',
      id: `${menu.id}-clear`,
      label: menu.clearLabel,
      radio: true,
      checked: current === null,
      disabled,
      description,
      onSelect: () => choose(null),
    },
    {
      kind: 'submenu',
      id: `${menu.id}-theme`,
      label: t('colorPicker.themeColors'),
      disabled,
      description,
      entries: columns.map((cells, column) => ({
        kind: 'submenu',
        id: `${menu.id}-theme-column-${column}`,
        label: cells[0].label,
        disabled,
        description,
        entries: cells.map((cell, row) =>
          swatch(cell.hex, column * cells.length + row, cell.label, cell)
        ),
      })),
    },
    {
      kind: 'group',
      id: `${menu.id}-standard`,
      label: t('colorPicker.standardColors'),
      entries: STANDARD_COLORS.map((color, index) =>
        swatch(color.hex, index, t(color.nameKey))
      ),
    },
  ];
  if (tools.prompt) {
    const prompt = tools.prompt;
    entries.push({
      kind: 'item',
      id: `${menu.id}-custom`,
      label: t('commands.customColor'),
      disabled,
      description,
      onSelect: () => {
        const pending = prepare(store, menu.command);
        prompt({
          title: menu.label,
          label: t('commands.hexColor'),
          placeholder: 'FF0000',
          initialValue: current ?? '',
          valid: (value) => /^#?[0-9a-f]{6}$/i.test(value),
          submit: (value) =>
            void pending.execute(menu.args({ hex: value.replace(/^#/, '').toUpperCase() })),
        });
      },
    });
  }
  return {
    kind: 'submenu',
    id: menu.id,
    label: menu.label,
    disabled,
    description,
    entries,
  };
}

function textColorArgs(color: { hex: string; cell?: ThemeMatrixCell } | null): DocxCommandArgs['textColor'] {
  if (!color) return { color: 'auto' };
  const cell = color.cell;
  if (!cell) return { color: { rgb: color.hex } };
  return {
    color: {
      themeColor: cell.themeSlot,
      ...(cell.tint ? { themeTint: cell.tint } : {}),
      ...(cell.shade ? { themeShade: cell.shade } : {}),
    },
  };
}

function highlightHex(value: unknown): string | null {
  if (typeof value !== 'string' || value === 'none') return null;
  const resolved = resolveHighlightColor(value);
  return (resolved || value).replace(/^#/, '');
}

function colorCommandSubmenu(
  store: DocxCommandStore,
  t: TFunction,
  id: 'textColor' | 'highlightColor',
  tools: OverflowTools
): OverflowMenuEntry {
  const state = store.getState(id) as DocxCommandState;
  const value = typeof state.value === 'string' ? state.value : null;
  return colorSubmenu(
    store,
    t,
    {
      id,
      label: t(store.getDescriptor(id).labelKey),
      clearLabel: t(id === 'textColor' ? 'colorPicker.automatic' : 'colorPicker.noColor'),
      current:
        id === 'highlightColor'
          ? highlightHex(value)
          : value && /^[0-9a-f]{6}$/i.test(value)
            ? value
            : value
              ? (resolveColorToHex({ themeColor: value } as ColorValue, tools.theme ?? undefined) ??
                null)
              : null,
      state,
      command: id,
      args: (color) =>
        id === 'textColor' ? textColorArgs(color) : { color: color ? color.hex : 'none' },
    },
    tools
  );
}

const BORDER_ACTIONS: readonly { action: DocxTableAction; labelKey: TranslationKey }[] = [
  { action: 'borderAll', labelKey: 'table.borders.all' },
  { action: 'borderOutside', labelKey: 'table.borders.outside' },
  { action: 'borderInside', labelKey: 'table.borders.inside' },
  { action: 'borderTop', labelKey: 'table.borders.top' },
  { action: 'borderBottom', labelKey: 'table.borders.bottom' },
  { action: 'borderLeft', labelKey: 'table.borders.left' },
  { action: 'borderRight', labelKey: 'table.borders.right' },
  { action: 'borderNone', labelKey: 'table.borders.none' },
];

const BORDER_WIDTHS: readonly { size: number; label: string }[] = [
  { size: 4, label: '0.5 pt' },
  { size: 8, label: '1 pt' },
  { size: 12, label: '1.5 pt' },
  { size: 16, label: '2 pt' },
  { size: 24, label: '3 pt' },
];

const TABLE_ACTIONS: readonly ({ action: DocxTableAction; labelKey: TranslationKey } | null)[] = [
  { action: 'addRowAbove', labelKey: 'table.insertRowAbove' },
  { action: 'addRowBelow', labelKey: 'table.insertRowBelow' },
  { action: 'addColumnLeft', labelKey: 'table.insertColumnLeft' },
  { action: 'addColumnRight', labelKey: 'table.insertColumnRight' },
  null,
  { action: 'mergeCells', labelKey: 'table.mergeCells' },
  { action: 'splitCell', labelKey: 'table.splitCell' },
  null,
  { action: 'selectTable', labelKey: 'table.selectTable' },
  null,
  { action: 'deleteRow', labelKey: 'table.deleteRow' },
  { action: 'deleteColumn', labelKey: 'table.deleteColumn' },
  { action: 'deleteTable', labelKey: 'table.deleteTable' },
];

const VERTICAL_ALIGNMENTS: readonly { align: 'top' | 'center' | 'bottom'; labelKey: TranslationKey }[] = [
  { align: 'top', labelKey: 'tableAdvanced.top' },
  { align: 'center', labelKey: 'tableAdvanced.middle' },
  { align: 'bottom', labelKey: 'tableAdvanced.bottom' },
];

const TABLE_ALIGNMENTS: readonly { align: 'left' | 'center' | 'right'; labelKey: TranslationKey }[] = [
  { align: 'left', labelKey: 'tableAdvanced.alignTableLeft' },
  { align: 'center', labelKey: 'tableAdvanced.alignTableCenter' },
  { align: 'right', labelKey: 'tableAdvanced.alignTableRight' },
];

const TABLE_OPTIONS: readonly { action: DocxTableAction; labelKey: TranslationKey }[] = [
  { action: { type: 'toggleHeaderRow' }, labelKey: 'tableAdvanced.toggleHeaderRow' },
  { action: { type: 'distributeColumns' }, labelKey: 'tableAdvanced.distributeColumns' },
  { action: { type: 'autoFitContents' }, labelKey: 'tableAdvanced.autoFit' },
  { action: { type: 'toggleNoWrap' }, labelKey: 'tableAdvanced.toggleNoWrap' },
  { action: { type: 'openTableProperties' }, labelKey: 'tableAdvanced.tableProperties' },
];

function tableSubmenu(store: DocxCommandStore, t: TFunction, tools: OverflowTools): OverflowMenuEntry {
  const state = store.getState('tableAction');
  const disabled = !state.enabled;
  const description = reasonOf(state as DocxCommandState);
  const key = (action: DocxTableAction) =>
    `tableAction-${typeof action === 'string' ? action : JSON.stringify(action)}`;
  const item = (action: DocxTableAction, label: string, checked?: boolean): OverflowMenuEntry => {
    const itemState = store.getState('tableAction', action);
    return {
      kind: 'item',
      id: key(action),
      label,
      ...(checked === undefined ? {} : { radio: true, checked }),
      disabled: !itemState.enabled,
      description: reasonOf(itemState as DocxCommandState),
      onSelect: () => run(store, 'tableAction', action),
    };
  };
  const submenu = (id: string, label: string, entries: OverflowMenuEntry[]): OverflowMenuEntry => ({
    kind: 'submenu',
    id: `tableAction-${id}`,
    label,
    disabled,
    description,
    entries,
  });
  const value = state.value ?? null;
  return {
    kind: 'submenu',
    id: 'tableAction',
    label: t(store.getDescriptor('tableAction').labelKey),
    disabled,
    description,
    entries: [
      ...TABLE_ACTIONS.map((entry, index): OverflowMenuEntry =>
        entry
          ? item(entry.action, t(entry.labelKey))
          : { kind: 'separator', id: `tableAction-separator-${index}` }
      ),
      { kind: 'separator', id: 'tableAction-separator-format' },
      submenu(
        'borders',
        t('table.borders.tooltip'),
        BORDER_ACTIONS.map((entry) => item(entry.action, t(entry.labelKey)))
      ),
      colorSubmenu(
        store,
        t,
        {
          id: 'tableAction-borderColor',
          label: t('table.borderColor'),
          clearLabel: t('colorPicker.automatic'),
          current: value?.borderColor ?? null,
          state: state as DocxCommandState,
          command: 'tableAction',
          args: (color) => ({ type: 'borderColor', color: color?.hex ?? '000000' }),
        },
        tools
      ),
      submenu(
        'borderWidth',
        t('table.borderWidth'),
        BORDER_WIDTHS.map((width) => item({ type: 'borderWidth', size: width.size }, width.label))
      ),
      colorSubmenu(
        store,
        t,
        {
          id: 'tableAction-cellFill',
          label: t('table.cellFillColor'),
          clearLabel: t('colorPicker.noColor'),
          current: value?.fillColor ?? null,
          state: state as DocxCommandState,
          command: 'tableAction',
          args: (color) => ({ type: 'cellFillColor', color: color?.hex ?? null }),
        },
        tools
      ),
      submenu(
        'verticalAlignment',
        t('tableAdvanced.verticalAlignment'),
        VERTICAL_ALIGNMENTS.map((entry) =>
          item({ type: 'cellVerticalAlign', align: entry.align }, t(entry.labelKey))
        )
      ),
      submenu(
        'tableAlignment',
        t('tableAdvanced.tableAlignment'),
        TABLE_ALIGNMENTS.map((entry) =>
          item(
            { type: 'tableProperties', props: { justification: entry.align } },
            t(entry.labelKey),
            (value?.justification ?? 'left') === entry.align
          )
        )
      ),
      { kind: 'separator', id: 'tableAction-separator-options' },
      ...TABLE_OPTIONS.map((entry) => item(entry.action, t(entry.labelKey))),
    ],
  };
}

const IMAGE_TRANSFORMS: readonly { action: DocxCommandArgs['imageTransform']['action']; labelKey: TranslationKey }[] = [
  { action: 'rotateCW', labelKey: 'imageTransform.rotateClockwise' },
  { action: 'rotateCCW', labelKey: 'imageTransform.rotateCounterClockwise' },
  { action: 'flipH', labelKey: 'imageTransform.flipHorizontal' },
  { action: 'flipV', labelKey: 'imageTransform.flipVertical' },
];

function transformSubmenu(store: DocxCommandStore, t: TFunction): OverflowMenuEntry {
  const state = store.getState('imageTransform');
  return {
    kind: 'submenu',
    id: 'imageTransform',
    label: t(store.getDescriptor('imageTransform').labelKey),
    disabled: !state.enabled,
    description: reasonOf(state as DocxCommandState),
    entries: IMAGE_TRANSFORMS.map((entry) => ({
      kind: 'item',
      id: `imageTransform-${entry.action}`,
      label: t(entry.labelKey),
      disabled: !state.enabled,
      onSelect: () => run(store, 'imageTransform', { action: entry.action }),
    })),
  };
}

/** Sizes the inline table grid offers. */
const TABLE_GRID_SIZE = 6;

function insertTableSubmenu(store: DocxCommandStore, t: TFunction): OverflowMenuEntry {
  const state = store.getState('insertTable');
  const disabled = !state.enabled;
  const description = reasonOf(state as DocxCommandState);
  const sizes = Array.from({ length: TABLE_GRID_SIZE }, (_, index) => index + 1);
  return {
    kind: 'submenu',
    id: 'insertTable',
    label: t(store.getDescriptor('insertTable').labelKey),
    disabled,
    description,
    entries: sizes.map((rows) => ({
      kind: 'submenu',
      id: `insertTable-rows-${rows}`,
      label: `${t('dialogs.insertTable.rowsLabel')} ${rows}`,
      disabled,
      description,
      entries: sizes.map((columns) => ({
        kind: 'item',
        id: `insertTable-${rows}x${columns}`,
        label: t('dialogs.insertTable.tableSize', { cols: columns, rows }),
        disabled,
        description,
        onSelect: () => run(store, 'insertTable', { rows, columns }),
      })),
    })),
  };
}

/** The overflow-menu presentation of one command. */
export function commandOverflowEntry<K extends DocxCommandId>(
  store: DocxCommandStore,
  t: TFunction,
  id: K,
  args?: DocxCommandArgs[K],
  tools: OverflowTools = { theme: null }
): OverflowMenuEntry {
  if (args === undefined) {
    switch (id) {
      case 'paragraphStyle':
      case 'fontFamily':
      case 'fontSize':
      case 'alignment':
      case 'lineSpacing':
      case 'imageWrap':
      case 'editingMode':
      case 'zoom':
        return selectSubmenu(store, t, id as DocxSelectCommandId, tools);
      case 'textColor':
      case 'highlightColor':
        return colorCommandSubmenu(store, t, id, tools);
      case 'tableAction':
        return tableSubmenu(store, t, tools);
      case 'imageTransform':
        return transformSubmenu(store, t);
      case 'insertTable':
        return insertTableSubmenu(store, t);
      default:
        break;
    }
  }
  return commandItem(store, t, id, args, args === undefined ? id : `${id}:${JSON.stringify(args)}`);
}

function useCommandOverflow<K extends DocxCommandId>(
  element: React.RefObject<HTMLElement | null>,
  id: K,
  args?: DocxCommandArgs[K]
): void {
  const store = useDocxCommands();
  const { t } = useTranslation();
  const theme = useDocxChrome()?.theme ?? null;
  const prompt = useContext(ToolbarOverflowContext)?.prompt;
  useOverflowSource(element, () => [commandOverflowEntry(store, t, id, args, { theme, prompt })]);
}

/**
 * A button bound to one command, showing its pressed and disabled state.
 * @experimental
 */
export function ToolbarCommandButton<K extends DocxCommandId>(
  props: ToolbarCommandButtonProps<K>
) {
  const { id, children, label, className } = props;
  const args = (props as { args?: DocxCommandArgs[K] }).args;
  const command = useDocxCommand(id, args);
  const icon = iconFor(id, args);
  const name = label ?? command.label;
  return (
    <ToolbarButton
      active={command.state.active}
      disabled={!command.state.enabled}
      description={reasonOf(command.state as DocxCommandState)}
      title={name}
      ariaLabel={name}
      shortcut={command.shortcut ?? undefined}
      className={cn(!icon && !children && 'w-auto px-2 text-sm', className)}
      onClick={() => void command.execute()}
    >
      {children ?? (icon ? <MaterialSymbol name={icon} size={ICON_SIZE} /> : name)}
    </ToolbarButton>
  );
}

function styleOptionsFor(state: DocxCommandState<'paragraphStyle'>): StyleOption[] {
  return (state.options ?? []).map((option) => ({
    styleId: option.args.styleId,
    name: option.label,
    type: 'paragraph',
    ...(option.preview?.fontSize != null ? { fontSize: option.preview.fontSize * 2 } : {}),
    ...(option.preview?.bold != null ? { bold: option.preview.bold } : {}),
    ...(option.preview?.italic != null ? { italic: option.preview.italic } : {}),
    ...(option.preview?.color != null ? { color: option.preview.color } : {}),
  }));
}

function fontOptionsFor(state: DocxCommandState<'fontFamily'>): {
  fonts: FontOption[];
  documentFonts: FontOption[];
} {
  const fonts: FontOption[] = [];
  const documentFonts: FontOption[] = [];
  for (const option of state.options ?? []) {
    const group = option.preview?.group;
    const font: FontOption = {
      name: option.label,
      fontFamily: option.preview?.fontFamily ?? option.args.family,
      ...(group === 'sans-serif' || group === 'serif' || group === 'monospace' || group === 'other'
        ? { category: group }
        : {}),
    };
    if (group === 'document') documentFonts.push(font);
    else fonts.push(font);
  }
  return { fonts, documentFonts };
}

function wrapToolbarValue(value: string | null): { wrapType: string; displayMode: string; cssFloat: string | null } {
  if (value === 'squareLeft') return { wrapType: 'square', displayMode: 'float', cssFloat: 'left' };
  if (value === 'squareRight') return { wrapType: 'square', displayMode: 'float', cssFloat: 'right' };
  return { wrapType: value ?? 'inline', displayMode: value === 'inline' ? 'inline' : 'block', cssFloat: null };
}

/**
 * The built-in picker of a selector command.
 * @experimental
 */
export function ToolbarCommandSelect<K extends DocxSelectCommandId>({
  id,
  className,
}: ToolbarCommandSelectProps<K>) {
  const store = useDocxCommands();
  const state = useDocxCommandState(id) as DocxCommandState;
  const wrapperRef = useRef<HTMLSpanElement>(null);
  useCommandOverflow(wrapperRef, id);
  const disabled = !state.enabled;
  const description = reasonOf(state);
  const choose = <C extends DocxSelectCommandId>(
    command: C,
    args: DocxCommandArgs[C],
    focus: 'pointer' | 'always' = 'pointer'
  ) =>
    void store.execute(command, args).then((result: DocxCommandResult) => {
      if (focus === 'always') restoreEditorFocus(store);
      else restoreFocusAfterPointer(store, result);
    });

  let control: ReactNode = null;
  switch (id) {
    case 'paragraphStyle':
      control = (
        <StylePicker
          value={(state.value as string | undefined) ?? 'Normal'}
          options={styleOptionsFor(state as DocxCommandState<'paragraphStyle'>)}
          onChange={(styleId) => choose('paragraphStyle', { styleId })}
          disabled={disabled}
          description={description}
          width={120}
        />
      );
      break;
    case 'fontFamily': {
      const { fonts, documentFonts } = fontOptionsFor(state as DocxCommandState<'fontFamily'>);
      control = (
        <FontPicker
          value={(state.value as string | null) ?? 'Arial'}
          fonts={fonts}
          documentFonts={documentFonts}
          onChange={(family) => choose('fontFamily', { family })}
          disabled={disabled}
          description={description}
          width={60}
          placeholder="Arial"
        />
      );
      break;
    }
    case 'fontSize':
      control = (
        <FontSizePicker
          value={(state.value as number | null) ?? 11}
          onChange={(points) => choose('fontSize', { points }, 'always')}
          disabled={disabled}
          description={description}
          width={42}
          placeholder="11"
        />
      );
      break;
    case 'alignment':
      control = (
        <AlignmentButtons
          value={(state.value as DocxCommandArgs['alignment']['value'] | null) ?? 'left'}
          onChange={(value) => choose('alignment', { value })}
          disabled={disabled}
          description={description}
        />
      );
      break;
    case 'lineSpacing':
      control = (
        <LineSpacingPicker
          value={(state.value as number | null) ?? undefined}
          onChange={(value) => choose('lineSpacing', { value })}
          disabled={disabled}
          description={description}
        />
      );
      break;
    case 'imageWrap':
      control = (
        <ImageWrapDropdown
          imageContext={wrapToolbarValue(state.value as string | null)}
          onChange={(value) => {
            const wrap = toolbarValueToLayoutTarget(value);
            if (wrap) choose('imageWrap', { wrap });
          }}
          disabled={disabled}
          description={description}
        />
      );
      break;
    case 'editingMode':
      control = (
        <EditingModeDropdown
          mode={(state.value as DocxCommandArgs['editingMode']['mode']) ?? 'editing'}
          onModeChange={(mode) => choose('editingMode', { mode })}
          disabled={disabled}
          description={description}
          optionState={(mode) => {
            const option = store.getState('editingMode', { mode });
            return { enabled: option.enabled, description: reasonOf(option as DocxCommandState) };
          }}
        />
      );
      break;
    case 'zoom':
      control = (
        <ZoomControl
          value={(state.value as number | undefined) ?? 1}
          onChange={(scale) => choose('zoom', { scale })}
          disabled={disabled}
          compact
        />
      );
      break;
  }
  return (
    <span ref={wrapperRef} className={cn('inline-flex flex-shrink-0 items-center', className)}>
      {control}
    </span>
  );
}

function CommandColorPicker({ id }: { id: 'textColor' | 'highlightColor' }) {
  const store = useDocxCommands();
  const state = useDocxCommandState(id);
  const { t } = useTranslation();
  const wrapperRef = useRef<HTMLSpanElement>(null);
  useCommandOverflow(wrapperRef, id);
  const theme = useDocxChrome()?.theme ?? null;
  const raw = typeof state.value === 'string' ? state.value : undefined;
  const value =
    id === 'textColor' && raw && !/^[0-9a-f]{6}$/i.test(raw)
      ? (resolveColorToHex({ themeColor: raw } as ColorValue, theme ?? undefined) ?? undefined)
      : raw;
  return (
    <span ref={wrapperRef} className="inline-flex flex-shrink-0">
      <ColorPicker
        mode={id === 'textColor' ? 'text' : 'highlight'}
        value={value}
        theme={theme}
        disabled={!state.enabled}
        description={reasonOf(state as DocxCommandState)}
        title={t(store.getDescriptor(id).labelKey)}
        onChange={(color) => {
          const args: DocxCommandArgs['textColor'] | DocxCommandArgs['highlightColor'] =
            id === 'highlightColor'
              ? { color: typeof color === 'string' ? color : 'none' }
              : {
                  color:
                    typeof color === 'string'
                      ? { rgb: color.replace(/^#/, '') }
                      : color.auto
                        ? 'auto'
                        : color.themeColor
                          ? {
                              themeColor: color.themeColor,
                              ...(color.themeTint ? { themeTint: color.themeTint } : {}),
                              ...(color.themeShade ? { themeShade: color.themeShade } : {}),
                            }
                          : { rgb: (color.rgb ?? '000000').replace(/^#/, '') },
                };
          void store
            .execute(id, args as DocxCommandArgs['textColor'])
            .then((result) => restoreFocusAfterPointer(store, result));
        }}
      />
    </span>
  );
}

function CommandTableControls() {
  const store = useDocxCommands();
  const state = useDocxCommandState('tableAction');
  const { t } = useTranslation();
  const wrapperRef = useRef<HTMLSpanElement>(null);
  useCommandOverflow(wrapperRef, 'tableAction');
  const theme = useDocxChrome()?.theme ?? null;
  const value = state.value;
  const disabled = !state.enabled;
  const onAction = (action: DocxTableAction) =>
    void store
      .execute('tableAction', action)
      .then((result) => restoreFocusAfterPointer(store, result));
  return (
    <span
      ref={wrapperRef}
      className="inline-flex flex-shrink-0 items-center"
      role="group"
      aria-label={t('formattingBar.groups.table')}
    >
      <TableBorderPicker onAction={onAction} disabled={disabled} />
      <TableBorderColorPicker
        onAction={onAction}
        disabled={disabled}
        theme={theme}
        value={value?.borderColor ?? undefined}
      />
      <TableBorderWidthPicker onAction={onAction} disabled={disabled} />
      <TableCellFillPicker
        onAction={onAction}
        disabled={disabled}
        theme={theme}
        value={value?.fillColor ?? undefined}
      />
      <TableMoreDropdown
        onAction={onAction}
        disabled={disabled}
        tableContext={{
          isInTable: state.enabled,
          table: value?.justification ? { attrs: { justification: value.justification } } : undefined,
        }}
        actionState={(action) => {
          const itemState = store.getState('tableAction', action);
          return { enabled: itemState.enabled, description: reasonOf(itemState as DocxCommandState) };
        }}
      />
    </span>
  );
}

function CommandInsertTable({ className }: { className?: string }) {
  const command = useDocxCommand('insertTable');
  const store = useDocxCommands();
  const wrapperRef = useRef<HTMLSpanElement>(null);
  useCommandOverflow(wrapperRef, 'insertTable');
  const [open, setOpen] = useState(false);
  const close = useCallback(() => setOpen(false), []);
  const { containerRef, dropdownRef, dropdownStyle, handleMouseDown } = useFixedDropdown({
    isOpen: open,
    onClose: close,
  });
  const disabled = !command.state.enabled;
  const reason = reasonOf(command.state as DocxCommandState);
  const button = (
    <Button
      variant="ghost"
      size="icon-sm"
      className={cn('oox-toolbar-toggle text-muted-foreground', disabled && 'opacity-30', className)}
      onMouseDown={handleMouseDown}
      onClick={() => !disabled && setOpen((value) => !value)}
      disabled={disabled && !reason}
      aria-disabled={disabled && reason ? true : undefined}
      aria-label={command.label}
      aria-haspopup="true"
      aria-expanded={open}
      title={reason}
    >
      <MaterialSymbol name="grid_on" size={ICON_SIZE} />
    </Button>
  );
  return (
    <span ref={wrapperRef} className="inline-flex flex-shrink-0">
      <div ref={containerRef} style={{ position: 'relative', display: 'inline-block' }}>
        {open ? button : <Tooltip content={reason ? `${command.label}: ${reason}` : command.label}>{button}</Tooltip>}
        {open && (
          <div
            ref={dropdownRef}
            data-docx-escape-layer="true"
            onMouseDown={(event) => event.preventDefault()}
            style={{
              ...dropdownStyle,
              background: 'var(--doc-surface)',
              border: '1px solid var(--doc-border)',
              borderRadius: 8,
              boxShadow: '0 4px 12px var(--doc-shadow)',
              padding: 8,
            }}
          >
            <TableGridInline
              onInsert={(rows, columns) => {
                setOpen(false);
                void command
                  .execute({ rows, columns })
                  .then((result) => restoreFocusAfterPointer(store, result));
              }}
            />
          </div>
        )}
      </div>
    </span>
  );
}

function CommandImageTransform() {
  const store = useDocxCommands();
  const state = useDocxCommandState('imageTransform');
  const wrapperRef = useRef<HTMLSpanElement>(null);
  useCommandOverflow(wrapperRef, 'imageTransform');
  return (
    <span ref={wrapperRef} className="inline-flex flex-shrink-0">
      <ImageTransformDropdown
        disabled={!state.enabled}
        description={reasonOf(state as DocxCommandState)}
        onTransform={(action) =>
          void store
            .execute('imageTransform', { action })
            .then((result) => restoreFocusAfterPointer(store, result))
        }
      />
    </span>
  );
}

/**
 * The built-in control of any command, as the default toolbar presents it.
 * @experimental
 */
export function ToolbarCommand<K extends DocxCommandId>(props: ToolbarCommandProps<K>) {
  const { id, args, className } = props;
  if (args === undefined) {
    switch (id) {
      case 'paragraphStyle':
      case 'fontFamily':
      case 'fontSize':
      case 'alignment':
      case 'lineSpacing':
      case 'imageWrap':
      case 'editingMode':
      case 'zoom':
        return <ToolbarCommandSelect id={id as DocxSelectCommandId} className={className} />;
      case 'textColor':
      case 'highlightColor':
        return <CommandColorPicker id={id} />;
      case 'tableAction':
        return <CommandTableControls />;
      case 'imageTransform':
        return <CommandImageTransform />;
      case 'insertTable':
        return <CommandInsertTable className={className} />;
      default:
        break;
    }
  }
  const CommandButton = ToolbarCommandButton as unknown as ComponentType<{
    id: DocxCommandId;
    args?: unknown;
    className?: string;
  }>;
  return <CommandButton id={id} args={args} className={className} />;
}
