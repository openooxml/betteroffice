import type { TFunction } from '@betteroffice/xlsx-i18n';
import {
  isPluginCommandId,
  XLSX_COMMAND_DESCRIPTORS,
  XLSX_COMMAND_IDS,
} from '../../commands/descriptors';
import {
  commandReason,
  DEFAULT_FONT_FAMILIES,
  DEFAULT_FONT_SIZES,
  evaluateXlsxCommand,
  type XlsxCommandEnvironment,
} from '../../commands/evaluate';
import type {
  XlsxCommandArgs,
  XlsxCommandDescriptor,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandState,
  XlsxCommandStore,
} from '../../commands/types';
import type { FormattingAction, ToolbarProps } from '../Toolbar';

type Handler = 'onFormat' | 'onMerge' | 'onSearchMenus' | 'onUndo' | 'onRedo' | 'onPrint' | 'onZoomChange';

const HANDLERS: Record<XlsxCommandId, Handler | null> = {
  bold: 'onFormat',
  italic: 'onFormat',
  strikethrough: 'onFormat',
  paintFormat: 'onFormat',
  fontFamily: 'onFormat',
  fontSize: 'onFormat',
  fontSizeStep: 'onFormat',
  textColor: 'onFormat',
  fillColor: 'onFormat',
  numberFormat: 'onFormat',
  decimalPlaces: 'onFormat',
  borderPreset: 'onFormat',
  borderStyle: 'onFormat',
  borderColor: 'onFormat',
  merge: 'onMerge',
  horizontalAlignment: 'onFormat',
  verticalAlignment: 'onFormat',
  textWrapping: 'onFormat',
  searchMenus: 'onSearchMenus',
  save: null,
  exportPng: null,
  print: 'onPrint',
  undo: 'onUndo',
  redo: 'onRedo',
  zoom: 'onZoomChange',
  proposalsPanel: null,
  proposalAccept: null,
  proposalReject: null,
};

const LEGACY_FONT_SIZE = 10;

function nextFontSize(value: number, sizes: readonly number[], direction: -1 | 1): number {
  if (direction > 0) return sizes.find((size) => size > value) ?? value + 1;
  return [...sizes].reverse().find((size) => size < value) ?? Math.max(1, value - 1);
}

function formattingAction<K extends XlsxCommandId>(
  id: K,
  args: XlsxCommandArgs[K],
  env: XlsxCommandEnvironment
): FormattingAction | null {
  const bound = args as unknown as Record<string, never>;
  switch (id) {
    case 'bold':
    case 'italic':
    case 'strikethrough':
    case 'paintFormat':
      return id;
    case 'numberFormat':
      return { type: 'numberFormat', value: bound.value };
    case 'decimalPlaces':
      return bound.direction === 'increase' ? 'increaseDecimal' : 'decreaseDecimal';
    case 'fontSizeStep': {
      const selection = env.selection;
      const current =
        (typeof selection === 'object' && selection?.formatting?.fontSize) || LEGACY_FONT_SIZE;
      return {
        type: 'fontSize',
        value: nextFontSize(current, env.fontSizes, bound.direction === 'increase' ? 1 : -1),
      };
    }
    case 'fontFamily':
      return { type: 'fontFamily', value: bound.family };
    case 'fontSize':
      return { type: 'fontSize', value: bound.points };
    case 'textColor':
    case 'fillColor':
    case 'borderColor':
      return { type: id, value: bound.color };
    case 'borderPreset':
    case 'borderStyle':
    case 'horizontalAlignment':
    case 'verticalAlignment':
    case 'textWrapping':
      return { type: id, value: bound.value } as FormattingAction;
    default:
      return null;
  }
}

/** The gate inputs a prop-configured toolbar describes. */
function legacyEnvironment(props: ToolbarProps, t: TFunction): XlsxCommandEnvironment {
  const formatting = props.currentFormatting ?? {};
  const rows = Math.max(1, props.selectionShape?.rows ?? 1);
  const columns = Math.max(1, props.selectionShape?.columns ?? 1);
  const cell = { row: 0, col: 0 };
  return {
    status: 'ready',
    readOnly: false,
    collaborative: false,
    selection: {
      sheet: 0,
      top: 0,
      left: 0,
      bottom: rows - 1,
      right: columns - 1,
      merged: props.selectionShape?.canUnmerge ? [{ start: cell, end: cell }] : [],
      formatting: {
        ...formatting,
        fontFamily: formatting.fontFamily ?? 'Calibri',
        fontSize: formatting.fontSize ?? LEGACY_FONT_SIZE,
        bold: formatting.bold ?? false,
        italic: formatting.italic ?? false,
        strikethrough: formatting.strikethrough ?? false,
      },
    },
    canUndo: props.canUndo ?? false,
    canRedo: props.canRedo ?? false,
    pendingInput: false,
    zoom: props.zoom ?? 1,
    paintFormat: formatting.paintFormat ?? false,
    borderStyle: formatting.borderStyle ?? null,
    borderColor: formatting.borderColor ?? null,
    pngExport: false,
    proposals: null,
    proposalsPanelOpen: false,
    fontFamilies: props.fontFamilies ?? DEFAULT_FONT_FAMILIES,
    fontSizes: props.fontSizes ?? DEFAULT_FONT_SIZES,
    hostDisabled: new Set(
      XLSX_COMMAND_IDS.filter((id) => {
        const handler = HANDLERS[id];
        return props.disabled || !handler || !props[handler];
      })
    ),
    translate: t,
  };
}

/**
 * A command store over a prop-configured toolbar: state comes from the props
 * and execution calls their callbacks, so legacy toolbars render the same
 * controls as command-bound ones.
 */
export function createLegacyToolbarStore(
  props: ToolbarProps,
  handlers: () => ToolbarProps,
  t: TFunction
): XlsxCommandStore {
  const env = legacyEnvironment(props, t);
  const snapshots = new Map<string, XlsxCommandState>();
  const evaluate = <K extends XlsxCommandId>(id: K, args?: XlsxCommandArgs[K]) => {
    if (isPluginCommandId(id)) {
      return {
        enabled: false,
        disabledReason: commandReason('unsupported-command', env),
      } as XlsxCommandState<K>;
    }
    if (id === 'searchMenus' && !env.hostDisabled?.has(id)) {
      return { enabled: true } as XlsxCommandState<K>;
    }
    return evaluateXlsxCommand(id, args, env);
  };
  return Object.freeze({
    getDescriptor<K extends XlsxCommandId>(id: K): XlsxCommandDescriptor<K> | null {
      return isPluginCommandId(id)
        ? null
        : (XLSX_COMMAND_DESCRIPTORS[id] as XlsxCommandDescriptor<K>);
    },
    getState<K extends XlsxCommandId>(id: K, args?: XlsxCommandArgs[K]): XlsxCommandState<K> {
      const key = args === undefined ? id : `${id}\u0000${JSON.stringify(args)}`;
      let state = snapshots.get(key);
      if (!state) {
        state = evaluate(id, args) as XlsxCommandState;
        snapshots.set(key, state);
      }
      return state as XlsxCommandState<K>;
    },
    subscribe: () => () => {},
    async execute<K extends XlsxCommandId>(
      id: K,
      args: XlsxCommandArgs[K]
    ): Promise<XlsxCommandResult> {
      const state = evaluate(id, args);
      if (!state.enabled) return { ok: false, failure: state.disabledReason };
      const current = handlers();
      if (id === 'merge') current.onMerge?.((args as XlsxCommandArgs['merge']).value);
      else if (id === 'searchMenus') current.onSearchMenus?.();
      else if (id === 'undo') current.onUndo?.();
      else if (id === 'redo') current.onRedo?.();
      else if (id === 'print') current.onPrint?.();
      else if (id === 'zoom') current.onZoomChange?.((args as XlsxCommandArgs['zoom']).scale);
      else {
        const action = formattingAction(id, args, env);
        if (action) current.onFormat?.(action);
      }
      return { ok: true, status: 'requested' };
    },
  }) as XlsxCommandStore;
}

type CommandCall = { [K in XlsxCommandId]: { id: K; args: XlsxCommandArgs[K] } }[XlsxCommandId];

/** The command a legacy formatting callback stands for. */
export function commandForAction(action: FormattingAction): CommandCall {
  switch (action) {
    case 'bold':
    case 'italic':
    case 'strikethrough':
    case 'paintFormat':
      return { id: action, args: null };
    case 'currency':
    case 'percent':
      return { id: 'numberFormat', args: { value: action } };
    case 'increaseDecimal':
      return { id: 'decimalPlaces', args: { direction: 'increase' } };
    case 'decreaseDecimal':
      return { id: 'decimalPlaces', args: { direction: 'decrease' } };
  }
  switch (action.type) {
    case 'fontFamily':
      return { id: 'fontFamily', args: { family: action.value } };
    case 'fontSize':
      return { id: 'fontSize', args: { points: action.value } };
    case 'textColor':
    case 'fillColor':
    case 'borderColor':
      return { id: action.type, args: { color: action.value } } as CommandCall;
    default:
      return { id: action.type, args: { value: action.value } } as CommandCall;
  }
}
