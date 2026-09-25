import type { ParagraphAlignment } from '@betteroffice/pptx';
import type { TFunction, TranslationKey } from '@betteroffice/pptx-i18n';
import {
  adjustmentLimit,
  primaryAdjustment,
  shapePresetFromTool,
  SHAPE_PRESETS,
  type PptxEditorTool,
  type PptxZoom,
} from '../components/toolbarTypes';
import type {
  CommandReason,
  PptxCommandArgs,
  PptxCommandDisabledCode,
  PptxCommandFailureCode,
  PptxCommandId,
  PptxCommandOption,
  PptxCommandState,
  PptxCommandValues,
} from './types';

/** A selector choice before its own state is evaluated. */
export type PptxCommandChoice<K extends PptxCommandId> = Omit<PptxCommandOption<K>, 'state'>;

/** Formatting of the text a command applies to. */
export interface PptxTextEnvironment {
  /** A collapsed caret, a text range, or the whole story of a selected shape. */
  kind: 'caret' | 'range' | 'shape';
  bold: boolean | 'mixed';
  italic: boolean | 'mixed';
  underline: boolean | 'mixed';
  /** `null` when mixed. */
  fontFamily: string | null;
  /** Points; `null` when mixed. */
  fontSize: number | null;
  /** `#rrggbb`; `null` when mixed. */
  color: string | null;
  /** `null` when the paragraphs differ. */
  alignment: ParagraphAlignment | null;
}

/** The selected object. */
export interface PptxShapeEnvironment {
  /** Fill, outline and adjustments apply: a preset shape, not a picture or frame. */
  formattable: boolean;
  geometry: string | null;
  fill: string | null;
  stroke: string | null;
  strokeWidth: number | null;
  adjustments: Readonly<Record<string, number>>;
  /** Paint position among its siblings; `null` when unknown. */
  order: { index: number; count: number } | null;
}

export interface PptxProposalEnvironment {
  id: string;
  label: string;
  stale: boolean;
}

/** The proposal review the canvas shows. */
export interface PptxCanvasReviewEnvironment {
  /** Proposals touching the current slide, in canvas order. */
  available: readonly PptxProposalEnvironment[];
  selectedId: string | null;
  /** The canvas shows the diff rather than the editable slide. */
  diff: boolean;
}

/** Inputs of the command gate; the editor reads them live before execution. */
export interface PptxCommandEnvironment {
  status: 'ready' | 'loading' | 'empty' | 'unavailable';
  readOnly: boolean;
  /** The canvas shows a proposal diff instead of the editable slide. */
  reviewing: boolean;
  pendingInput: boolean;
  canUndo: boolean;
  canRedo: boolean;
  /** The current slide, or `null` in a deck without slides. */
  slide: { id: string; layoutPartPath: string | null } | null;
  layouts: readonly PptxCommandChoice<'insertSlide'>[];
  /** `'unsupported'` when the text target could not be read. */
  text: PptxTextEnvironment | 'unsupported' | null;
  shape: PptxShapeEnvironment | null;
  tool: PptxEditorTool;
  zoom: PptxZoom;
  fontFamilies: readonly string[];
  fontSizes: readonly number[];
  /** `null` when this editor cannot hold proposals. */
  proposals: {
    open: boolean;
    pending: readonly PptxProposalEnvironment[];
    canvas: PptxCanvasReviewEnvironment;
  } | null;
  /** A host-owned binding without a handler for `id`. */
  hostDisabled?(id: PptxCommandId): boolean;
  translate: TFunction;
}

const REASON_KEYS: Record<PptxCommandFailureCode, TranslationKey> = {
  'editor-unavailable': 'commands.reasons.editorUnavailable',
  'document-loading': 'commands.reasons.documentLoading',
  'no-document': 'commands.reasons.noDocument',
  'read-only': 'commands.reasons.readOnly',
  'nothing-to-undo': 'commands.reasons.nothingToUndo',
  'nothing-to-redo': 'commands.reasons.nothingToRedo',
  'text-selection-required': 'commands.reasons.textSelectionRequired',
  'shape-required': 'commands.reasons.shapeRequired',
  'slide-required': 'commands.reasons.slideRequired',
  'unsupported-selection': 'commands.reasons.unsupportedSelection',
  'invalid-layout': 'commands.reasons.invalidLayout',
  'invalid-adjustment': 'commands.reasons.invalidAdjustment',
  'z-order-boundary': 'commands.reasons.zOrderBoundary',
  'review-active': 'commands.reasons.reviewActive',
  'preview-pending': 'commands.reasons.previewPending',
  'preview-failed': 'commands.reasons.previewFailed',
  'proposal-stale': 'commands.reasons.proposalStale',
  'proposal-not-found': 'commands.reasons.proposalNotFound',
  'proposals-unavailable': 'commands.reasons.proposalsUnavailable',
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
  'command-failed': 'commands.reasons.commandFailed',
};

/** A localized reason; without an editor the message is English. */
export function commandReason<Code extends PptxCommandFailureCode>(
  code: Code,
  env: Pick<PptxCommandEnvironment, 'translate'> | null,
  key: TranslationKey = REASON_KEYS[code]
): CommandReason<Code> {
  return {
    code,
    message: env
      ? env.translate(key)
      : code === 'editor-unavailable'
      ? 'The editor is not ready.'
      : code,
  };
}

type Gate = PptxCommandDisabledCode | null;

interface Evaluation<K extends PptxCommandId> {
  gate: Gate;
  reasonKey?: TranslationKey;
  active?: boolean | 'mixed';
  value?: PptxCommandValues[K];
  options?: readonly PptxCommandChoice<K>[];
}

export const MIN_FONT_POINTS = 1;
export const MAX_FONT_POINTS = 400;
export const MIN_ZOOM = 0.25;
export const MAX_ZOOM = 4;
export const MAX_STROKE_POINTS = 1584;
export const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2] as const;
export const STROKE_WIDTHS = [1, 2, 3, 4, 8] as const;
export const DEFAULT_FONT_FAMILIES = [
  'Arial',
  'Calibri',
  'Cambria',
  'Georgia',
  'Roboto',
  'Times New Roman',
  'Verdana',
] as const;
export const DEFAULT_FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72] as const;
/** The size stepping starts from when the selection mixes sizes. */
export const FALLBACK_FONT_POINTS = 24;

const ALIGNMENTS: readonly { value: ParagraphAlignment; labelKey: TranslationKey }[] = [
  { value: 'l', labelKey: 'toolbar.align.left' },
  { value: 'ctr', labelKey: 'toolbar.align.center' },
  { value: 'r', labelKey: 'toolbar.align.right' },
  { value: 'just', labelKey: 'toolbar.align.justify' },
];

const Z_ORDER_MOVES = new Set(['front', 'forward', 'backward', 'back']);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function finite(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function nonEmpty(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

function hex(value: unknown): value is string {
  return typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value);
}

export function isEditorTool(value: unknown): value is PptxEditorTool {
  return (
    value === 'select' ||
    value === 'textBox' ||
    (typeof value === 'string' && shapePresetFromTool(value) !== null)
  );
}

/** The next size in `sizes` above or below `points`, stepping by one past either end. */
export function nextFontSize(points: number, sizes: readonly number[], direction: -1 | 1): number {
  if (direction > 0)
    return sizes.find((size) => size > points) ?? Math.min(MAX_FONT_POINTS, points + 1);
  return (
    [...sizes].reverse().find((size) => size < points) ?? Math.max(MIN_FONT_POINTS, points - 1)
  );
}

function validArgs<K extends PptxCommandId>(id: K, args: PptxCommandArgs[K] | undefined): boolean {
  if (args === undefined) return true;
  const value = args as unknown;
  switch (id) {
    case 'fontFamily':
      return isRecord(value) && nonEmpty(value.family);
    case 'fontSize':
      return (
        isRecord(value) &&
        finite(value.points) &&
        value.points >= MIN_FONT_POINTS &&
        value.points <= MAX_FONT_POINTS
      );
    case 'fontSizeStep':
      return isRecord(value) && (value.direction === 'increase' || value.direction === 'decrease');
    case 'textColor':
      return isRecord(value) && hex(value.color);
    case 'alignment':
      return isRecord(value) && ALIGNMENTS.some((option) => option.value === value.value);
    case 'insertSlide':
      return (
        isRecord(value) &&
        (value.layoutPartPath === undefined ||
          value.layoutPartPath === null ||
          nonEmpty(value.layoutPartPath))
      );
    case 'tool':
      return isRecord(value) && isEditorTool(value.value);
    case 'shapeFill':
    case 'shapeStrokeColor':
      return isRecord(value) && (value.color === null || hex(value.color));
    case 'shapeStrokeWidth':
      return (
        isRecord(value) &&
        (value.points === null ||
          (finite(value.points) && value.points > 0 && value.points <= MAX_STROKE_POINTS))
      );
    case 'shapeAdjustment':
      return isRecord(value) && nonEmpty(value.name) && finite(value.value);
    case 'zOrder':
      return isRecord(value) && Z_ORDER_MOVES.has(value.value as string);
    case 'zoom':
      return (
        isRecord(value) &&
        (value.scale === 'fit' ||
          (finite(value.scale) && value.scale >= MIN_ZOOM && value.scale <= MAX_ZOOM))
      );
    case 'proposalsPanel':
      return value === null || (isRecord(value) && typeof value.open === 'boolean');
    case 'proposalSelect':
    case 'proposalReject':
      return isRecord(value) && nonEmpty(value.proposalId);
    case 'proposalDiff':
      return isRecord(value) && typeof value.enabled === 'boolean';
    case 'proposalAccept':
      return (
        isRecord(value) &&
        nonEmpty(value.proposalId) &&
        (value.force === undefined || typeof value.force === 'boolean')
      );
    default:
      return value === null;
  }
}

function documentGate(env: PptxCommandEnvironment): Gate {
  switch (env.status) {
    case 'unavailable':
      return 'editor-unavailable';
    case 'loading':
      return 'document-loading';
    case 'empty':
      return 'no-document';
    default:
      return null;
  }
}

function writeGate(env: PptxCommandEnvironment): Gate {
  return documentGate(env) ?? (env.readOnly ? 'read-only' : null);
}

function slideGate(env: PptxCommandEnvironment): Gate {
  return env.slide ? null : 'slide-required';
}

function reviewGate(env: PptxCommandEnvironment): Gate {
  return env.reviewing ? 'review-active' : null;
}

function textGate(env: PptxCommandEnvironment): Gate {
  return (
    writeGate(env) ??
    reviewGate(env) ??
    (env.text === 'unsupported'
      ? 'unsupported-selection'
      : env.text
      ? null
      : 'text-selection-required')
  );
}

function shapeGate(env: PptxCommandEnvironment): Gate {
  if (writeGate(env)) return writeGate(env);
  if (env.reviewing) return 'review-active';
  if (!env.shape) return 'shape-required';
  return env.shape.formattable ? null : 'unsupported-selection';
}

function text(env: PptxCommandEnvironment): PptxTextEnvironment | null {
  return env.text && env.text !== 'unsupported' ? env.text : null;
}

function zOrderGate(
  env: PptxCommandEnvironment,
  move: PptxCommandArgs['zOrder']['value'] | undefined
): Gate {
  const base = writeGate(env) ?? reviewGate(env) ?? (env.shape ? null : 'shape-required');
  if (base || !env.shape?.order) return base;
  const { index, count } = env.shape.order;
  const top = index >= count - 1;
  const bottom = index <= 0;
  if (move === undefined) return top && bottom ? 'z-order-boundary' : null;
  return (move === 'front' || move === 'forward' ? top : bottom) ? 'z-order-boundary' : null;
}

function zoomLabel(scale: PptxZoom, env: PptxCommandEnvironment): string {
  return scale === 'fit' ? env.translate('toolbar.fit') : `${Math.round(scale * 100)}%`;
}

function sameZoom(a: PptxZoom, b: PptxZoom): boolean {
  return a === 'fit' || b === 'fit' ? a === b : Math.abs(a - b) < 0.001;
}

function toolOptions(env: PptxCommandEnvironment): PptxCommandChoice<'tool'>[] {
  return [
    { args: { value: 'select' }, label: env.translate('commands.selectTool') },
    { args: { value: 'textBox' }, label: env.translate('toolbar.textBoxTool') },
    ...SHAPE_PRESETS.map((preset) => ({
      args: { value: `shape:${preset.geometry}` as PptxEditorTool },
      label: env.translate(preset.labelKey),
    })),
  ];
}

function evaluateProposal(
  id: 'proposalAccept' | 'proposalReject',
  env: PptxCommandEnvironment,
  args: PptxCommandArgs['proposalAccept'] | undefined
): Evaluation<'proposalAccept'> {
  const base = writeGate(env) ?? (env.proposals ? null : 'proposals-unavailable');
  if (base || !args) return { gate: base };
  const proposal = env.proposals!.pending.find((candidate) => candidate.id === args.proposalId);
  if (!proposal) return { gate: 'proposal-not-found' };
  if (id === 'proposalReject') return { gate: null };
  return { gate: proposal.stale && !args.force ? 'proposal-stale' : null };
}

function evaluate<K extends PptxCommandId>(
  id: K,
  args: PptxCommandArgs[K] | undefined,
  env: PptxCommandEnvironment
): Evaluation<K> {
  const current = text(env);
  const t = env.translate;
  const result = <V extends PptxCommandId>(evaluation: Evaluation<V>) =>
    evaluation as unknown as Evaluation<K>;
  switch (id) {
    case 'bold':
    case 'italic':
    case 'underline':
      return { gate: textGate(env), active: current?.[id as 'bold' | 'italic' | 'underline'] };
    case 'fontFamily': {
      const families = env.fontFamilies;
      return result<'fontFamily'>({
        gate: textGate(env),
        value: current ? current.fontFamily : undefined,
        options: families.map((family) => ({ args: { family }, label: family })),
        active: args
          ? current?.fontFamily === (args as PptxCommandArgs['fontFamily']).family
          : undefined,
      });
    }
    case 'fontSize': {
      const sizeArgs = args as PptxCommandArgs['fontSize'] | undefined;
      return result<'fontSize'>({
        gate: textGate(env),
        value: current ? current.fontSize : undefined,
        options: env.fontSizes.map((points) => ({ args: { points }, label: String(points) })),
        active: sizeArgs ? current?.fontSize === sizeArgs.points : undefined,
      });
    }
    case 'fontSizeStep':
      return { gate: textGate(env) };
    case 'textColor':
      return result<'textColor'>({
        gate: textGate(env),
        value: current ? current.color : undefined,
      });
    case 'alignment': {
      const alignmentArgs = args as PptxCommandArgs['alignment'] | undefined;
      return result<'alignment'>({
        gate: textGate(env),
        value: current ? current.alignment : undefined,
        options: ALIGNMENTS.map((option) => ({
          args: { value: option.value },
          label: t(option.labelKey),
        })),
        active: alignmentArgs ? current?.alignment === alignmentArgs.value : undefined,
      });
    }
    case 'insertSlide': {
      const slideArgs = args as PptxCommandArgs['insertSlide'] | undefined;
      const layout = slideArgs?.layoutPartPath;
      const known =
        typeof layout !== 'string' ||
        env.layouts.some((option) => option.args.layoutPartPath === layout);
      return result<'insertSlide'>({
        gate: writeGate(env) ?? (known ? null : 'invalid-layout'),
        value: env.slide ? env.slide.layoutPartPath : undefined,
        options: env.layouts,
        active:
          slideArgs && layout !== undefined && env.slide
            ? (layout ?? null) === env.slide.layoutPartPath
            : undefined,
      });
    }
    case 'insertImage':
      return { gate: writeGate(env) ?? slideGate(env) ?? reviewGate(env) };
    case 'tool': {
      const tool = (args as PptxCommandArgs['tool'] | undefined)?.value;
      const gate = tool === 'select' ? documentGate(env) : writeGate(env) ?? slideGate(env);
      return result<'tool'>({
        gate,
        value: env.tool,
        options: toolOptions(env),
        active: tool ? env.tool === tool : undefined,
      });
    }
    case 'shapeFill':
      return result<'shapeFill'>({
        gate: shapeGate(env),
        value: env.shape?.formattable ? env.shape.fill : undefined,
      });
    case 'shapeStrokeColor':
      return result<'shapeStrokeColor'>({
        gate: shapeGate(env),
        value: env.shape?.formattable ? env.shape.stroke : undefined,
      });
    case 'shapeStrokeWidth': {
      const widthArgs = args as PptxCommandArgs['shapeStrokeWidth'] | undefined;
      const value = env.shape?.formattable ? env.shape.strokeWidth : undefined;
      return result<'shapeStrokeWidth'>({
        gate: shapeGate(env),
        value,
        options: [
          { args: { points: null }, label: t('toolbar.noBorder') },
          ...STROKE_WIDTHS.map((width) => ({
            args: { points: width },
            label: t('toolbar.borderWidthValue', { width }),
          })),
        ],
        active: widthArgs && value !== undefined ? value === widthArgs.points : undefined,
      });
    }
    case 'shapeAdjustment': {
      const shape = env.shape?.formattable ? env.shape : null;
      const primary = shape ? primaryAdjustment(shape.adjustments as Record<string, number>) : null;
      const adjustmentArgs = args as PptxCommandArgs['shapeAdjustment'] | undefined;
      const gate = shapeGate(env) ?? (primary ? null : 'invalid-adjustment');
      const limit = primary ? adjustmentLimit(shape!.geometry, primary[0]) : 1;
      const invalid =
        adjustmentArgs &&
        shape &&
        (!Object.prototype.hasOwnProperty.call(shape.adjustments, adjustmentArgs.name) ||
          adjustmentArgs.value < 0 ||
          adjustmentArgs.value > adjustmentLimit(shape.geometry, adjustmentArgs.name));
      const steps = primary
        ? Array.from({ length: Math.round(limit * 10) + 1 }, (_, index) => index / 10)
        : [];
      return result<'shapeAdjustment'>({
        gate: gate ?? (invalid ? 'invalid-adjustment' : null),
        value: shape ? (primary ? { name: primary[0], value: primary[1] } : null) : undefined,
        options: primary
          ? steps.map((value) => ({
              args: { name: primary[0], value },
              label: `${Math.round(value * 100)}%`,
            }))
          : [],
        active:
          adjustmentArgs && primary
            ? adjustmentArgs.name === primary[0] &&
              Math.abs(adjustmentArgs.value - primary[1]) < 0.0005
            : undefined,
      });
    }
    case 'zOrder':
      return {
        gate: zOrderGate(env, (args as PptxCommandArgs['zOrder'] | undefined)?.value),
        reasonKey: env.shape ? undefined : 'commands.reasons.objectRequired',
      };
    case 'save':
      return { gate: documentGate(env) };
    case 'exportPng':
    case 'slideshow':
      return { gate: documentGate(env) ?? slideGate(env) };
    case 'undo':
      return {
        gate: writeGate(env) ?? (env.canUndo || env.pendingInput ? null : 'nothing-to-undo'),
      };
    case 'redo':
      return {
        gate: writeGate(env) ?? (env.canRedo || env.pendingInput ? null : 'nothing-to-redo'),
      };
    case 'zoom': {
      const zoomArgs = args as PptxCommandArgs['zoom'] | undefined;
      const levels: PptxZoom[] = ['fit', ...ZOOM_LEVELS];
      return result<'zoom'>({
        gate: documentGate(env),
        value: env.zoom,
        options: levels.map((scale) => ({ args: { scale }, label: zoomLabel(scale, env) })),
        active: zoomArgs ? sameZoom(env.zoom, zoomArgs.scale) : undefined,
      });
    }
    case 'proposalsPanel': {
      const panelArgs = args as PptxCommandArgs['proposalsPanel'] | undefined;
      const open = env.proposals?.open ?? false;
      return result<'proposalsPanel'>({
        gate: writeGate(env) ?? (env.proposals ? null : 'proposals-unavailable'),
        value: open,
        active: panelArgs ? open === panelArgs.open : open,
      });
    }
    case 'proposalSelect': {
      const selectArgs = args as PptxCommandArgs['proposalSelect'] | undefined;
      const canvas = env.proposals?.canvas;
      const base = writeGate(env) ?? (canvas ? null : 'proposals-unavailable');
      const known =
        !selectArgs ||
        !canvas ||
        canvas.available.some((proposal) => proposal.id === selectArgs.proposalId);
      return result<'proposalSelect'>({
        gate:
          base ??
          (canvas && canvas.available.length === 0 ? 'proposal-not-found' : null) ??
          (known ? null : 'proposal-not-found'),
        value: canvas?.selectedId ?? null,
        options: (canvas?.available ?? []).map((proposal) => ({
          args: { proposalId: proposal.id },
          label: proposal.label,
        })),
        active: selectArgs ? canvas?.selectedId === selectArgs.proposalId : undefined,
      });
    }
    case 'proposalDiff': {
      const diffArgs = args as PptxCommandArgs['proposalDiff'] | undefined;
      const canvas = env.proposals?.canvas;
      return result<'proposalDiff'>({
        gate:
          writeGate(env) ??
          (canvas ? null : 'proposals-unavailable') ??
          (canvas?.selectedId ? null : 'proposal-not-found'),
        value: canvas?.diff ?? false,
        active: diffArgs ? (canvas?.diff ?? false) === diffArgs.enabled : canvas?.diff ?? false,
      });
    }
    case 'proposalAccept':
    case 'proposalReject':
      return result(
        evaluateProposal(id, env, args as PptxCommandArgs['proposalAccept'] | undefined)
      );
    default:
      return { gate: 'unsupported-command' };
  }
}

function choiceState<K extends PptxCommandId>(
  state: PptxCommandState<K>
): PptxCommandOption<K>['state'] {
  const active = state.active === undefined ? {} : { active: state.active };
  return state.enabled
    ? { ...active, enabled: true }
    : { ...active, enabled: false, disabledReason: state.disabledReason };
}

/** The editor's gate for a contributed command; null when it passes. */
export function contributedCommandGate(
  mutatesDocument: boolean,
  env: PptxCommandEnvironment | null
): CommandReason<PptxCommandDisabledCode> | null {
  if (!env) return commandReason('editor-unavailable', null);
  const gate = mutatesDocument ? writeGate(env) : documentGate(env);
  return gate ? commandReason(gate, env) : null;
}

/** The single availability gate for snapshots and execution. */
export function evaluatePptxCommand<K extends PptxCommandId>(
  id: K,
  args: PptxCommandArgs[K] | undefined,
  env: PptxCommandEnvironment | null
): PptxCommandState<K> {
  if (!env) {
    return { enabled: false, disabledReason: commandReason('editor-unavailable', null) };
  }
  const evaluation: Evaluation<K> = validArgs(id, args)
    ? evaluate(id, args, env)
    : { gate: documentGate(env) ?? 'invalid-arguments' };
  const gate = env.hostDisabled?.(id) ? 'host-disabled' : evaluation.gate;
  const presentation: {
    active?: boolean | 'mixed';
    value?: PptxCommandValues[K];
    options?: readonly PptxCommandOption<K>[];
  } = {};
  if (evaluation.active !== undefined) presentation.active = evaluation.active;
  if (evaluation.value !== undefined) presentation.value = evaluation.value;
  if (evaluation.options !== undefined && args === undefined) {
    presentation.options = evaluation.options.map((option) => ({
      ...option,
      state: choiceState(evaluatePptxCommand(id, option.args, env)),
    }));
  }
  return gate
    ? {
        ...presentation,
        enabled: false,
        disabledReason: commandReason(
          gate,
          env,
          gate === evaluation.gate ? evaluation.reasonKey : undefined
        ),
      }
    : { ...presentation, enabled: true };
}
