import type { MergedRange, SelectionFormatting } from '@betteroffice/xlsx';
import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import type {
  BorderPreset,
  BorderStyle,
  HorizontalAlignment,
  MergeAction,
  NumberFormat,
  TextWrapping,
  VerticalAlignment,
} from '../components/Toolbar';
import { commandLabelKey, XLSX_COMMAND_DESCRIPTORS } from './descriptors';
import type {
  CommandReason,
  XlsxCommandArgs,
  XlsxCommandDisabledCode,
  XlsxCommandFailureCode,
  XlsxCommandId,
  XlsxCommandOption,
  XlsxCommandState,
  XlsxCommandValues,
} from './types';

/** The selected cell range, with what the engine reports about it. */
export interface XlsxSelectionEnvironment {
  sheet: number;
  /** Zero-based, inclusive. */
  top: number;
  left: number;
  bottom: number;
  right: number;
  /** Merged ranges intersecting the selection. */
  merged: readonly MergedRange[];
  /** The engine's common formatting, absent fields being mixed; `null` when unread. */
  formatting: SelectionFormatting | null;
}

export interface XlsxProposalEnvironment {
  id: string;
  label: string;
}

/** Inputs of the command gate; the editor reads them live before execution. */
export interface XlsxCommandEnvironment {
  status: 'ready' | 'loading' | 'empty';
  readOnly: boolean;
  collaborative: boolean;
  /** `'chart'` while a chart object is selected instead of cells. */
  selection: XlsxSelectionEnvironment | 'chart' | null;
  canUndo: boolean;
  canRedo: boolean;
  /** Cell or formula text accepted but not yet written. */
  pendingInput: boolean;
  zoom: number;
  paintFormat: boolean;
  borderStyle: BorderStyle | null;
  borderColor: string | null;
  pngExport: boolean;
  /** `null` when the engine was built without proposals. */
  proposals: readonly XlsxProposalEnvironment[] | null;
  proposalsPanelOpen: boolean;
  fontFamilies: readonly string[];
  fontSizes: readonly number[];
  /** Commands the host left without an implementation. */
  hostDisabled?: ReadonlySet<XlsxCommandId>;
  translate(key: TranslationKey): string;
}

const REASON_KEYS: Record<XlsxCommandFailureCode, TranslationKey> = {
  'editor-unavailable': 'commands.reasons.editorUnavailable',
  'document-loading': 'commands.reasons.documentLoading',
  'no-document': 'commands.reasons.noDocument',
  'read-only': 'commands.reasons.readOnly',
  'cell-selection-required': 'commands.reasons.cellSelectionRequired',
  'unsupported-selection': 'commands.reasons.unsupportedSelection',
  'multiple-cells-required': 'commands.reasons.multipleCellsRequired',
  'multiple-columns-required': 'commands.reasons.multipleColumnsRequired',
  'multiple-rows-required': 'commands.reasons.multipleRowsRequired',
  'no-merged-cells': 'commands.reasons.noMergedCells',
  'collaboration-unsupported': 'commands.reasons.collaborationUnsupported',
  'nothing-to-undo': 'commands.reasons.nothingToUndo',
  'nothing-to-redo': 'commands.reasons.nothingToRedo',
  'png-unavailable': 'commands.reasons.pngUnavailable',
  'no-proposals': 'commands.reasons.noProposals',
  'proposal-not-found': 'commands.reasons.proposalNotFound',
  'host-disabled': 'commands.reasons.hostDisabled',
  'unsupported-command': 'commands.reasons.unsupportedCommand',
  'invalid-arguments': 'commands.reasons.invalidArguments',
  'permission-denied': 'commands.reasons.permissionDenied',
  'unsupported-policy': 'commands.reasons.unsupportedPolicy',
  'plugin-unavailable': 'commands.reasons.pluginUnavailable',
  aborted: 'commands.reasons.aborted',
  'input-failed': 'commands.reasons.inputFailed',
  'document-replaced': 'commands.reasons.documentReplaced',
  'target-changed': 'commands.reasons.targetChanged',
  'gesture-active': 'commands.reasons.gestureActive',
  'proposal-stale': 'commands.reasons.proposalStale',
  'render-failed': 'commands.reasons.renderFailed',
  'execution-failed': 'commands.reasons.executionFailed',
};

const ENGLISH_REASONS: Partial<Record<XlsxCommandFailureCode, string>> = {
  'editor-unavailable': 'The editor is not ready.',
};

/** A localized reason; without an editor the message is English. */
export function commandReason<Code extends XlsxCommandFailureCode>(
  code: Code,
  env: Pick<XlsxCommandEnvironment, 'translate'> | null
): CommandReason<Code> {
  return { code, message: env ? env.translate(REASON_KEYS[code]) : (ENGLISH_REASONS[code] ?? code) };
}

export const NUMBER_FORMATS: readonly NumberFormat[] = [
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
export const BORDER_PRESETS: readonly BorderPreset[] = [
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
export const BORDER_STYLES: readonly BorderStyle[] = ['solid', 'dashed', 'dotted', 'double'];
export const HORIZONTAL_ALIGNMENTS: readonly HorizontalAlignment[] = ['left', 'center', 'right'];
export const VERTICAL_ALIGNMENTS: readonly VerticalAlignment[] = ['top', 'middle', 'bottom'];
export const TEXT_WRAPPINGS: readonly TextWrapping[] = ['overflow', 'wrap', 'clip'];
export const MERGE_ACTIONS: readonly MergeAction[] = ['all', 'horizontal', 'vertical', 'unmerge'];
export const ZOOM_LEVELS: readonly number[] = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2];
export const DEFAULT_FONT_FAMILIES: readonly string[] = [
  'Arial',
  'Calibri',
  'Cambria',
  'Georgia',
  'Roboto',
  'Times New Roman',
  'Verdana',
];
export const DEFAULT_FONT_SIZES: readonly number[] = [
  8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72,
];
export const MIN_FONT_POINTS = 1;
export const MAX_FONT_POINTS = 400;
export const MIN_ZOOM = 0.25;
export const MAX_ZOOM = 4;

const CELL_COMMANDS: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>([
  'bold',
  'italic',
  'strikethrough',
  'paintFormat',
  'fontFamily',
  'fontSize',
  'fontSizeStep',
  'textColor',
  'fillColor',
  'numberFormat',
  'decimalPlaces',
  'borderPreset',
  'borderStyle',
  'borderColor',
  'merge',
  'horizontalAlignment',
  'verticalAlignment',
  'textWrapping',
]);

/** Commands whose effect is a session change the read-only restriction also covers. */
const WRITE_COMMANDS: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>(['proposalReject']);

const DIRECTIONS: ReadonlySet<string> = new Set(['increase', 'decrease']);

type Gate = XlsxCommandDisabledCode | null;

type Choice<K extends XlsxCommandId> = Omit<XlsxCommandOption<K>, 'state'>;

interface Evaluation<K extends XlsxCommandId> {
  gate: Gate;
  active?: boolean | 'mixed';
  value?: XlsxCommandValues[K];
  options?: readonly Choice<K>[];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function member<T extends string>(values: readonly T[], value: unknown): value is T {
  return typeof value === 'string' && (values as readonly string[]).includes(value);
}

export function isHexColor(value: unknown): value is string {
  return typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value);
}

function finiteIn(value: unknown, min: number, max: number): boolean {
  return typeof value === 'number' && Number.isFinite(value) && value >= min && value <= max;
}

/** Whether `args` are well-formed for `id`; `undefined` evaluates the whole control. */
export function validArguments(id: XlsxCommandId, args: unknown): boolean {
  if (args === undefined) return true;
  switch (id) {
    case 'fontFamily':
      return isRecord(args) && typeof args.family === 'string' && args.family.trim() !== '';
    case 'fontSize':
      return isRecord(args) && finiteIn(args.points, MIN_FONT_POINTS, MAX_FONT_POINTS);
    case 'fontSizeStep':
    case 'decimalPlaces':
      return isRecord(args) && typeof args.direction === 'string' && DIRECTIONS.has(args.direction);
    case 'textColor':
    case 'fillColor':
    case 'borderColor':
      return isRecord(args) && isHexColor(args.color);
    case 'numberFormat':
      return isRecord(args) && member(NUMBER_FORMATS, args.value);
    case 'borderPreset':
      return isRecord(args) && member(BORDER_PRESETS, args.value);
    case 'borderStyle':
      return isRecord(args) && member(BORDER_STYLES, args.value);
    case 'merge':
      return isRecord(args) && member(MERGE_ACTIONS, args.value);
    case 'horizontalAlignment':
      return isRecord(args) && member(HORIZONTAL_ALIGNMENTS, args.value);
    case 'verticalAlignment':
      return isRecord(args) && member(VERTICAL_ALIGNMENTS, args.value);
    case 'textWrapping':
      return isRecord(args) && member(TEXT_WRAPPINGS, args.value);
    case 'zoom':
      return isRecord(args) && finiteIn(args.scale, MIN_ZOOM, MAX_ZOOM);
    case 'proposalsPanel':
      return args === null || (isRecord(args) && typeof args.open === 'boolean');
    case 'proposalAccept':
      return (
        isRecord(args) &&
        typeof args.proposalId === 'string' &&
        (args.force === undefined || typeof args.force === 'boolean')
      );
    case 'proposalReject':
      return isRecord(args) && typeof args.proposalId === 'string';
    default:
      return args === null;
  }
}

function mark(value: boolean | undefined, formatting: SelectionFormatting | null): boolean | 'mixed' | undefined {
  if (!formatting) return undefined;
  return value === undefined ? 'mixed' : value;
}

function lifecycle(env: XlsxCommandEnvironment): Gate {
  if (env.status === 'loading') return 'document-loading';
  if (env.status === 'empty') return 'no-document';
  return null;
}

function options<K extends XlsxCommandId>(
  env: XlsxCommandEnvironment,
  id: K,
  list: readonly XlsxCommandArgs[K][]
): Choice<K>[] {
  return list.map((args) => ({ args, label: env.translate(commandLabelKey(id, args)) }));
}

function mergeGate(
  env: XlsxCommandEnvironment,
  selection: XlsxSelectionEnvironment,
  action: MergeAction
): Gate {
  if (env.collaborative) return 'collaboration-unsupported';
  const rows = selection.bottom - selection.top + 1;
  const columns = selection.right - selection.left + 1;
  switch (action) {
    case 'all':
      return rows > 1 || columns > 1 ? null : 'multiple-cells-required';
    case 'horizontal':
      return columns > 1 ? null : 'multiple-columns-required';
    case 'vertical':
      return rows > 1 ? null : 'multiple-rows-required';
    case 'unmerge':
      return selection.merged.length > 0 ? null : 'no-merged-cells';
  }
}

function evaluateCells<K extends XlsxCommandId>(
  id: K,
  args: XlsxCommandArgs[K] | undefined,
  env: XlsxCommandEnvironment,
  selection: XlsxSelectionEnvironment
): Evaluation<K> {
  const formatting = selection.formatting;
  const bound = args as Record<string, unknown> | undefined;
  const choose = <V extends string>(value: V | null, list: readonly V[]): Evaluation<K> => ({
    gate: null,
    value: value as XlsxCommandValues[K],
    active: bound ? value === bound.value : undefined,
    options: options(
      env,
      id,
      list.map((entry) => ({ value: entry }) as unknown as XlsxCommandArgs[K])
    ),
  });
  switch (id) {
    case 'bold':
      return { gate: null, active: mark(formatting?.bold, formatting) };
    case 'italic':
      return { gate: null, active: mark(formatting?.italic, formatting) };
    case 'strikethrough':
      return { gate: null, active: mark(formatting?.strikethrough, formatting) };
    case 'paintFormat':
      return { gate: null, active: env.paintFormat };
    case 'fontFamily': {
      const value = formatting?.fontFamily ?? null;
      return {
        gate: null,
        value: value as XlsxCommandValues[K],
        active: bound ? value?.toLowerCase() === String(bound.family).toLowerCase() : undefined,
        options: env.fontFamilies.map((family) => ({
          args: { family } as XlsxCommandArgs[K],
          label: family,
        })),
      };
    }
    case 'fontSize': {
      const value = formatting?.fontSize ?? null;
      return {
        gate: null,
        value: value as XlsxCommandValues[K],
        active: bound ? value === bound.points : undefined,
        options: env.fontSizes.map((points) => ({
          args: { points } as XlsxCommandArgs[K],
          label: String(points),
        })),
      };
    }
    case 'textColor':
      return { gate: null, value: (formatting?.textColor ?? null) as XlsxCommandValues[K] };
    case 'fillColor':
      return { gate: null, value: (formatting?.fillColor ?? null) as XlsxCommandValues[K] };
    case 'borderColor':
      return { gate: null, value: env.borderColor as XlsxCommandValues[K] };
    case 'numberFormat': {
      const value = formatting
        ? { kind: formatting.numberFormat ?? null, pattern: formatting.numberFormatPattern ?? null }
        : null;
      return {
        gate: null,
        value: value as XlsxCommandValues[K],
        active: bound ? value?.kind === bound.value : undefined,
        options: NUMBER_FORMATS.map((entry) => ({
          args: { value: entry } as XlsxCommandArgs[K],
          label: env.translate(`toolbar.numberFormats.${entry}`),
        })),
      };
    }
    case 'borderPreset':
      return choose(formatting?.borderPreset ?? null, BORDER_PRESETS);
    case 'borderStyle':
      return choose(env.borderStyle, BORDER_STYLES);
    case 'horizontalAlignment':
      return choose(formatting?.horizontalAlignment ?? null,
        HORIZONTAL_ALIGNMENTS);
    case 'verticalAlignment':
      return choose(formatting?.verticalAlignment ?? null,
        VERTICAL_ALIGNMENTS);
    case 'textWrapping':
      return choose(formatting?.textWrapping ?? null, TEXT_WRAPPINGS);
    case 'merge': {
      const value = {
        rows: selection.bottom - selection.top + 1,
        columns: selection.right - selection.left + 1,
      };
      const list = options(
        env,
        id,
        MERGE_ACTIONS.map((entry) => ({ value: entry }) as XlsxCommandArgs[K])
      );
      if (bound) {
        return {
          gate: mergeGate(env, selection, bound.value as MergeAction),
          value: value as XlsxCommandValues[K],
        };
      }
      const available = MERGE_ACTIONS.some((action) => !mergeGate(env, selection, action));
      return {
        gate: available ? null : mergeGate(env, selection, 'all'),
        value: value as XlsxCommandValues[K],
        options: list,
      };
    }
    default:
      return { gate: null };
  }
}

function evaluate<K extends XlsxCommandId>(
  id: K,
  args: XlsxCommandArgs[K] | undefined,
  env: XlsxCommandEnvironment
): Evaluation<K> {
  if (id === 'searchMenus') return { gate: 'unsupported-command' };
  if (id === 'zoom') {
    return {
      gate: null,
      value: env.zoom as XlsxCommandValues[K],
      active: args ? env.zoom === (args as XlsxCommandArgs['zoom']).scale : undefined,
      options: ZOOM_LEVELS.map((scale) => ({
        args: { scale } as XlsxCommandArgs[K],
        label: `${Math.round(scale * 100)}%`,
      })),
    };
  }
  const stage = lifecycle(env);
  if (stage) return { gate: stage };
  const readOnly: Gate =
    env.readOnly && (XLSX_COMMAND_DESCRIPTORS[id].mutatesDocument || WRITE_COMMANDS.has(id))
      ? 'read-only'
      : null;
  let evaluation: Evaluation<K>;
  if (!CELL_COMMANDS.has(id)) evaluation = evaluateSession(id, args, env);
  else if (env.selection === null) evaluation = { gate: 'cell-selection-required' };
  else if (env.selection === 'chart') evaluation = { gate: 'unsupported-selection' };
  else evaluation = evaluateCells(id, args, env, env.selection);
  return readOnly ? { ...evaluation, gate: readOnly } : evaluation;
}

function evaluateSession<K extends XlsxCommandId>(
  id: K,
  args: XlsxCommandArgs[K] | undefined,
  env: XlsxCommandEnvironment
): Evaluation<K> {
  switch (id) {
    case 'undo':
      return { gate: env.canUndo || env.pendingInput ? null : 'nothing-to-undo' };
    case 'redo':
      return { gate: env.canRedo ? null : 'nothing-to-redo' };
    case 'exportPng':
      return { gate: env.pngExport ? null : 'png-unavailable' };
    case 'proposalsPanel':
      if (!env.proposals) return { gate: 'unsupported-command' };
      return {
        gate: null,
        active: env.proposalsPanelOpen,
        value: env.proposals.length as XlsxCommandValues[K],
      };
    case 'proposalAccept':
    case 'proposalReject': {
      if (!env.proposals) return { gate: 'unsupported-command' };
      const list = env.proposals.map((proposal) => ({
        args: { proposalId: proposal.id } as XlsxCommandArgs[K],
        label: proposal.label,
      }));
      if (!args) return { gate: list.length > 0 ? null : 'no-proposals', options: list };
      const proposalId = (args as XlsxCommandArgs['proposalReject']).proposalId;
      return {
        gate: env.proposals.some((proposal) => proposal.id === proposalId)
          ? null
          : 'proposal-not-found',
      };
    }
    default:
      return { gate: null };
  }
}

/** The editor's gate for a contributed command; null when it passes. */
export function contributedCommandGate(
  mutatesDocument: boolean,
  env: XlsxCommandEnvironment | null
): CommandReason<XlsxCommandDisabledCode> | null {
  if (!env) return commandReason('editor-unavailable', null);
  const gate = lifecycle(env) ?? (mutatesDocument && env.readOnly ? 'read-only' : null);
  return gate ? commandReason(gate, env) : null;
}

/** Whether `id` acts on the selected cells. */
export function isCellCommand(id: XlsxCommandId): boolean {
  return CELL_COMMANDS.has(id);
}

/** Availability and state of `id`, or of one option when `args` are given. */
export function evaluateXlsxCommand<K extends XlsxCommandId>(
  id: K,
  args: XlsxCommandArgs[K] | undefined,
  env: XlsxCommandEnvironment | null
): XlsxCommandState<K> {
  if (!env) {
    return { enabled: false, disabledReason: commandReason('editor-unavailable', null) };
  }
  const evaluation: Evaluation<K> = !validArguments(id, args)
    ? { gate: 'invalid-arguments' }
    : env.hostDisabled?.has(id)
      ? { gate: 'host-disabled' }
      : evaluate(id, args, env);
  const extras = {
    ...(evaluation.active !== undefined ? { active: evaluation.active } : {}),
    ...(evaluation.value !== undefined ? { value: evaluation.value } : {}),
    ...(evaluation.options !== undefined && args === undefined
      ? {
          options: evaluation.options.map((option) => ({
            ...option,
            state: evaluateXlsxCommand(id, option.args, env),
          })),
        }
      : {}),
  };
  return evaluation.gate
    ? ({
        ...extras,
        enabled: false,
        disabledReason: commandReason(evaluation.gate, env),
      } as XlsxCommandState<K>)
    : ({ ...extras, enabled: true } as XlsxCommandState<K>);
}
