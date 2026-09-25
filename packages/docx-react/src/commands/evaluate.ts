import type { TranslationKey } from '@betteroffice/docx-i18n';
import type { ImageLayoutTarget } from '@betteroffice/docx/docx';
import type { ParagraphAlignment } from '@betteroffice/docx/types/document';
import type { YrsSelectionContext, YrsTriState } from '@betteroffice/docx/yrs';
import type { TableContextInfo } from '../components/DocxEditor/types';
import { EDITING_MODES, type EditorMode } from '../components/DocxEditor/internals/editing-modes';
import type {
  CommandReason,
  DocxCommandArgs,
  DocxCommandDisabledCode,
  DocxCommandFailureCode,
  DocxCommandId,
  DocxCommandOption,
  DocxCommandState,
  DocxCommandValues,
  DocxTableAction,
} from './types';

/** Who owns a piece of editor state: the editor, the host with a setter, or the host alone. */
export type DocxCommandControl = 'internal' | 'host' | 'fixed';

export interface DocxSelectionEnvironment {
  /** Authoritative selection state, including stored caret formatting. */
  context: YrsSelectionContext;
  /** Display font family, falling back to the paragraph style. */
  fontFamily: string | null;
  /** Display size in points, falling back to the paragraph style. */
  fontSize: number | null;
}

export interface DocxRevisionEnvironment {
  revisionId: string;
  /** Display position of the revision start, for navigation order. */
  position: number;
}

/** Inputs of the command gate; the editor reads them live before execution. */
export interface DocxCommandEnvironment {
  status: 'ready' | 'loading' | 'empty' | 'unavailable';
  readOnly: boolean;
  mode: EditorMode;
  modeControl: DocxCommandControl;
  sidebarOpen: boolean;
  sidebarControl: DocxCommandControl;
  zoom: number;
  canOpen: boolean;
  canReportIssue: boolean;
  /** The body, not a header, footer or note, has the input. */
  bodyStory: boolean;
  pendingInput: boolean;
  canUndo: boolean;
  canRedo: boolean;
  selection: DocxSelectionEnvironment | 'unsupported' | null;
  table: TableContextInfo | null;
  image: { wrap: ImageLayoutTarget | null } | null;
  /** Tracked changes the review commands navigate, in display order. */
  revisions: readonly DocxRevisionEnvironment[];
  /** Every tracked change in the document, including headers, footers and notes. */
  revisionIds: ReadonlySet<string>;
  currentRevisionId: string | null;
  styles: readonly DocxCommandOption<'paragraphStyle'>[];
  fonts: readonly DocxCommandOption<'fontFamily'>[];
  translate(key: TranslationKey): string;
}

const REASON_KEYS: Record<DocxCommandFailureCode, TranslationKey> = {
  'editor-unavailable': 'commands.reasons.editorUnavailable',
  'document-loading': 'commands.reasons.documentLoading',
  'no-document': 'commands.reasons.noDocument',
  'read-only': 'commands.reasons.readOnly',
  'viewing-mode': 'commands.reasons.viewingMode',
  'suggesting-unsupported': 'commands.reasons.suggestingUnsupported',
  'selection-required': 'commands.reasons.selectionRequired',
  'unsupported-selection': 'commands.reasons.unsupportedSelection',
  'unsupported-story': 'commands.reasons.unsupportedStory',
  'cannot-outdent': 'commands.reasons.cannotOutdent',
  'nothing-to-undo': 'commands.reasons.nothingToUndo',
  'nothing-to-redo': 'commands.reasons.nothingToRedo',
  'image-required': 'commands.reasons.imageRequired',
  'table-required': 'commands.reasons.tableRequired',
  'multiple-cells-required': 'commands.reasons.multipleCellsRequired',
  'cannot-split-cell': 'commands.reasons.cannotSplitCell',
  'last-row': 'commands.reasons.lastRow',
  'last-column': 'commands.reasons.lastColumn',
  'revision-required': 'commands.reasons.revisionRequired',
  'revision-not-found': 'commands.reasons.revisionNotFound',
  'no-revisions': 'commands.reasons.noRevisions',
  'controlled-mode': 'commands.reasons.controlledMode',
  'controlled-sidebar': 'commands.reasons.controlledSidebar',
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
  'command-failed': 'commands.reasons.commandFailed',
};

const ENGLISH_REASONS: Partial<Record<DocxCommandFailureCode, string>> = {
  'editor-unavailable': 'The editor is not ready.',
};

/** A localized reason; without an editor the message is English. */
export function commandReason<Code extends DocxCommandFailureCode>(
  code: Code,
  env: Pick<DocxCommandEnvironment, 'translate'> | null,
  key: TranslationKey = REASON_KEYS[code]
): CommandReason<Code> {
  return {
    code,
    message: env ? env.translate(key) : (ENGLISH_REASONS[code] ?? code),
  };
}

type Gate = DocxCommandDisabledCode | null;

interface Evaluation<K extends DocxCommandId> {
  gate: Gate;
  reasonKey?: TranslationKey;
  active?: boolean | 'mixed';
  value?: DocxCommandValues[K];
  options?: readonly DocxCommandOption<K>[];
}

const PARAGRAPH_ALIGNMENTS: ReadonlySet<string> = new Set<ParagraphAlignment>([
  'left',
  'center',
  'right',
  'both',
  'distribute',
  'mediumKashida',
  'highKashida',
  'lowKashida',
  'thaiDistribute',
]);

const IMAGE_WRAPS: ReadonlySet<string> = new Set<ImageLayoutTarget>([
  'inline',
  'squareLeft',
  'squareRight',
  'topAndBottom',
  'behind',
  'inFront',
  'square',
  'tight',
  'through',
]);

const IMAGE_TRANSFORMS: ReadonlySet<string> = new Set(['rotateCW', 'rotateCCW', 'flipH', 'flipV']);

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2];
const MAX_FONT_POINTS = 1638;
const MAX_TABLE_ROWS = 1000;
const MAX_TABLE_COLUMNS = 63;

const LINE_SPACINGS: readonly { twips: number; labelKey?: TranslationKey; label: string }[] = [
  { twips: 240, labelKey: 'lineSpacing.single', label: 'Single' },
  { twips: 276, label: '1.15' },
  { twips: 360, label: '1.5' },
  { twips: 480, labelKey: 'lineSpacing.double', label: 'Double' },
];

const ALIGNMENT_OPTIONS: readonly { value: ParagraphAlignment; labelKey: TranslationKey }[] = [
  { value: 'left', labelKey: 'alignment.alignLeft' },
  { value: 'center', labelKey: 'alignment.center' },
  { value: 'right', labelKey: 'alignment.alignRight' },
  { value: 'both', labelKey: 'alignment.justify' },
];

const WRAP_OPTIONS: readonly { wrap: ImageLayoutTarget; labelKey: TranslationKey }[] = [
  { wrap: 'inline', labelKey: 'imageWrap.inline' },
  { wrap: 'squareLeft', labelKey: 'imageWrap.floatLeft' },
  { wrap: 'squareRight', labelKey: 'imageWrap.floatRight' },
  { wrap: 'topAndBottom', labelKey: 'imageWrap.topAndBottom' },
  { wrap: 'behind', labelKey: 'imageWrap.behindText' },
  { wrap: 'inFront', labelKey: 'imageWrap.inFrontOfText' },
];

const SUPPORTED_TABLE_OBJECT_ACTIONS: ReadonlySet<string> = new Set([
  'cellFillColor',
  'borderColor',
  'borderWidth',
  'cellBorder',
  'tableProperties',
  'openTableProperties',
  'applyTableStyle',
]);

const SUGGESTING_TABLE_ACTIONS: ReadonlySet<string> = new Set([
  'addRowAbove',
  'addRowBelow',
  'deleteRow',
  'selectTable',
  'selectRow',
  'selectColumn',
]);

const TABLE_SELECTIONS: ReadonlySet<string> = new Set(['selectTable', 'selectRow', 'selectColumn']);

const TABLE_STRING_ACTIONS: ReadonlySet<string> = new Set([
  'addRowAbove',
  'addRowBelow',
  'addColumnLeft',
  'addColumnRight',
  'deleteRow',
  'deleteColumn',
  'mergeCells',
  'splitCell',
  'deleteTable',
  'selectTable',
  'selectRow',
  'selectColumn',
  'borderAll',
  'borderOutside',
  'borderInside',
  'borderNone',
  'borderTop',
  'borderBottom',
  'borderLeft',
  'borderRight',
]);

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
  return typeof value === 'string' && /^#?[0-9a-f]{6}$/i.test(value);
}

function byteOrAbsent(value: unknown): boolean {
  return value === undefined || (typeof value === 'string' && /^[0-9a-f]{2}$/i.test(value));
}

function tri(value: YrsTriState | undefined): boolean | 'mixed' {
  return value === 'mixed' ? 'mixed' : value === true;
}

function numberingOf(env: DocxCommandEnvironment): { numId?: number; ilvl?: number } | null {
  if (!env.selection || env.selection === 'unsupported') return null;
  const numPr = env.selection.context.paragraphProperties.numPr;
  return isRecord(numPr) && (finite(numPr.numId) || finite(numPr.ilvl))
    ? { numId: finite(numPr.numId) ? numPr.numId : undefined, ilvl: finite(numPr.ilvl) ? numPr.ilvl : undefined }
    : null;
}

function documentGate(env: DocxCommandEnvironment): Gate {
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

function writeGate(env: DocxCommandEnvironment): Gate {
  return (
    documentGate(env) ?? (env.readOnly ? 'read-only' : env.mode === 'viewing' ? 'viewing-mode' : null)
  );
}

function selectionGate(env: DocxCommandEnvironment): Gate {
  if (env.selection === null) return 'selection-required';
  return env.selection === 'unsupported' ? 'unsupported-selection' : null;
}

function textGate(env: DocxCommandEnvironment): Gate {
  return writeGate(env) ?? selectionGate(env);
}

function suggestingGate(env: DocxCommandEnvironment): Gate {
  return env.mode === 'suggesting' ? 'suggesting-unsupported' : null;
}

function bodyGate(env: DocxCommandEnvironment): Gate {
  return writeGate(env) ?? (env.bodyStory ? null : 'unsupported-story') ?? selectionGate(env);
}

function validArgs<K extends DocxCommandId>(id: K, args: DocxCommandArgs[K] | undefined): boolean {
  if (args === undefined) return true;
  const value = args as unknown;
  switch (id) {
    case 'paragraphStyle':
      return isRecord(value) && nonEmpty(value.styleId);
    case 'fontFamily':
      return isRecord(value) && nonEmpty(value.family);
    case 'fontSize':
      return (
        isRecord(value) && finite(value.points) && value.points >= 1 && value.points <= MAX_FONT_POINTS
      );
    case 'textColor':
      if (value === null || !isRecord(value)) return false;
      return (
        value.color === 'auto' ||
        (isRecord(value.color) &&
          (hex(value.color.rgb) ||
            (value.color.rgb === undefined &&
              nonEmpty(value.color.themeColor) &&
              byteOrAbsent(value.color.themeTint) &&
              byteOrAbsent(value.color.themeShade))))
      );
    case 'highlightColor':
      return isRecord(value) && nonEmpty(value.color);
    case 'alignment':
      return isRecord(value) && typeof value.value === 'string' && PARAGRAPH_ALIGNMENTS.has(value.value);
    case 'lineSpacing':
      return isRecord(value) && finite(value.value) && value.value > 0 && value.value <= 31680;
    case 'insertTable':
      return (
        isRecord(value) &&
        Number.isInteger(value.rows) &&
        Number.isInteger(value.columns) &&
        (value.rows as number) >= 1 &&
        (value.rows as number) <= MAX_TABLE_ROWS &&
        (value.columns as number) >= 1 &&
        (value.columns as number) <= MAX_TABLE_COLUMNS
      );
    case 'imageWrap':
      return isRecord(value) && typeof value.wrap === 'string' && IMAGE_WRAPS.has(value.wrap);
    case 'imageTransform':
      return isRecord(value) && typeof value.action === 'string' && IMAGE_TRANSFORMS.has(value.action);
    case 'tableAction':
      return validTableAction(value);
    case 'editingMode':
      return isRecord(value) && EDITING_MODES.some((mode) => mode.value === value.mode);
    case 'reviewAccept':
    case 'reviewReject':
      return value === null || (isRecord(value) && nonEmpty(value.revisionId));
    case 'zoom':
      return isRecord(value) && finite(value.scale) && value.scale >= 0.1 && value.scale <= 5;
    default:
      return value === null;
  }
}

function validTableAction(value: unknown): boolean {
  if (typeof value === 'string') return TABLE_STRING_ACTIONS.has(value);
  if (!isRecord(value) || typeof value.type !== 'string') return false;
  switch (value.type) {
    case 'cellFillColor':
      return value.color === null || hex(value.color);
    case 'borderColor':
      return hex(value.color);
    case 'borderWidth':
      return finite(value.size) && value.size >= 0;
    case 'cellBorder':
      return (
        ['top', 'bottom', 'left', 'right', 'all'].includes(String(value.side)) &&
        nonEmpty(value.style) &&
        finite(value.size) &&
        hex(value.color)
      );
    case 'tableProperties':
      return isRecord(value.props);
    case 'applyTableStyle':
      return nonEmpty(value.styleId);
    default:
      return true;
  }
}

function tableActionKey(action: DocxTableAction): string {
  return typeof action === 'string' ? action : action.type;
}

function evaluateTableAction(env: DocxCommandEnvironment, action: DocxTableAction | undefined): Gate {
  const key = action === undefined ? null : tableActionKey(action);
  const base = key && TABLE_SELECTIONS.has(key) ? documentGate(env) : writeGate(env);
  if (base) return base;
  if (!env.bodyStory) return 'unsupported-story';
  const table = env.table;
  if (!table?.isInTable) return 'table-required';
  if (key === null) return null;
  if (typeof action === 'object' && !SUPPORTED_TABLE_OBJECT_ACTIONS.has(key)) {
    return 'unsupported-command';
  }
  if (!SUGGESTING_TABLE_ACTIONS.has(key) && env.mode === 'suggesting') {
    return 'suggesting-unsupported';
  }
  switch (key) {
    case 'deleteRow':
      return (table.rowCount ?? 0) <= 1 ? 'last-row' : null;
    case 'deleteColumn':
      return (table.columnCount ?? 0) <= 1 ? 'last-column' : null;
    case 'mergeCells':
      return table.hasMultiCellSelection ? null : 'multiple-cells-required';
    case 'splitCell':
      return table.canSplitCell ? null : 'cannot-split-cell';
    default:
      return null;
  }
}

function selection(env: DocxCommandEnvironment): DocxSelectionEnvironment | null {
  return env.selection && env.selection !== 'unsupported' ? env.selection : null;
}

function evaluateEditingMode(
  env: DocxCommandEnvironment,
  args: DocxCommandArgs['editingMode'] | undefined
): Evaluation<'editingMode'> {
  const current: EditorMode = env.readOnly ? 'viewing' : env.mode;
  const options = EDITING_MODES.map((mode) => ({
    args: { mode: mode.value },
    label: env.translate(mode.labelKey),
  }));
  const unavailable = env.status === 'unavailable' ? 'editor-unavailable' : null;
  if (!args) {
    return {
      gate: unavailable ?? (env.readOnly ? 'read-only' : env.modeControl === 'fixed' ? 'controlled-mode' : null),
      value: current,
      options,
    };
  }
  const gate =
    unavailable ??
    (args.mode === current
      ? null
      : env.readOnly
        ? 'read-only'
        : env.modeControl === 'fixed'
          ? 'controlled-mode'
          : null);
  return { gate, value: current, options, active: args.mode === current };
}

function evaluateRevision(
  env: DocxCommandEnvironment,
  args: DocxCommandArgs['reviewAccept'] | undefined
): Gate {
  const gate = writeGate(env);
  if (gate) return gate;
  if (args) return env.revisionIds.has(args.revisionId) ? null : 'revision-not-found';
  return env.currentRevisionId ? null : 'revision-required';
}

function evaluate<K extends DocxCommandId>(
  id: K,
  args: DocxCommandArgs[K] | undefined,
  env: DocxCommandEnvironment
): Evaluation<K> {
  const current = selection(env);
  const context = current?.context;
  const t = env.translate;
  const result = <V extends DocxCommandId>(evaluation: Evaluation<V>) =>
    evaluation as unknown as Evaluation<K>;
  switch (id) {
    case 'undo':
      return result<'undo'>({
        gate: writeGate(env) ?? (env.canUndo || env.pendingInput ? null : 'nothing-to-undo'),
      });
    case 'redo':
      return result<'redo'>({
        gate: writeGate(env) ?? (env.canRedo || env.pendingInput ? null : 'nothing-to-redo'),
      });
    case 'bold':
    case 'italic':
    case 'underline':
      return { gate: textGate(env), active: tri(context?.[id as 'bold' | 'italic' | 'underline']) };
    case 'strikethrough':
      return { gate: textGate(env), active: tri(context?.strike) };
    case 'superscript':
    case 'subscript':
      return { gate: textGate(env), active: tri(context?.[id as 'superscript' | 'subscript']) };
    case 'clearFormatting':
      return { gate: textGate(env) };
    case 'paragraphStyle': {
      const styleArgs = args as DocxCommandArgs['paragraphStyle'] | undefined;
      const known =
        !styleArgs || env.styles.some((option) => option.args.styleId === styleArgs.styleId);
      return result<'paragraphStyle'>({
        gate: textGate(env) ?? (known ? null : 'invalid-arguments'),
        value: context?.styleId ?? 'Normal',
        options: env.styles,
        active: styleArgs ? (context?.styleId ?? 'Normal') === styleArgs.styleId : undefined,
      });
    }
    case 'fontFamily':
      return result<'fontFamily'>({
        gate: textGate(env),
        value: current?.fontFamily ?? null,
        options: env.fonts,
      });
    case 'fontSize':
      return result<'fontSize'>({
        gate: textGate(env),
        value: current?.fontSize ?? null,
        options: FONT_SIZES.map((points) => ({ args: { points }, label: String(points) })),
      });
    case 'textColor':
      return result<'textColor'>({ gate: textGate(env), value: context?.color ?? null });
    case 'highlightColor':
      return result<'highlightColor'>({ gate: textGate(env), value: context?.highlight ?? null });
    case 'alignment': {
      const value = context?.alignment && PARAGRAPH_ALIGNMENTS.has(context.alignment)
        ? (context.alignment as ParagraphAlignment)
        : 'left';
      const alignmentArgs = args as DocxCommandArgs['alignment'] | undefined;
      return result<'alignment'>({
        gate: textGate(env),
        value: current ? value : null,
        options: ALIGNMENT_OPTIONS.map((option) => ({
          args: { value: option.value },
          label: t(option.labelKey),
        })),
        active: alignmentArgs ? current !== null && value === alignmentArgs.value : undefined,
      });
    }
    case 'lineSpacing': {
      const spacing = context?.paragraphProperties.lineSpacing;
      const spacingArgs = args as DocxCommandArgs['lineSpacing'] | undefined;
      return result<'lineSpacing'>({
        gate: textGate(env),
        value: finite(spacing) ? spacing : null,
        options: LINE_SPACINGS.map((option) => ({
          args: { value: option.twips },
          label: option.labelKey ? t(option.labelKey) : option.label,
        })),
        active: spacingArgs ? spacing === spacingArgs.value : undefined,
      });
    }
    case 'bulletList':
      return { gate: textGate(env), active: numberingOf(env)?.numId === 1 };
    case 'numberedList': {
      const numbering = numberingOf(env);
      return { gate: textGate(env), active: numbering?.numId != null && numbering.numId !== 1 };
    }
    case 'indent':
      return { gate: textGate(env) };
    case 'outdent': {
      const indented =
        numberingOf(env) !== null ||
        (finite(context?.paragraphProperties.indentLeft) &&
          (context.paragraphProperties.indentLeft as number) > 0);
      return { gate: textGate(env) ?? (indented ? null : 'cannot-outdent') };
    }
    case 'setLtr':
      return { gate: textGate(env), active: current ? context?.paragraphProperties.bidi !== true : false };
    case 'setRtl':
      return { gate: textGate(env), active: context?.paragraphProperties.bidi === true };
    case 'insertLink':
    case 'insertImage':
      return { gate: bodyGate(env) };
    case 'insertTable':
    case 'insertPageBreak':
    case 'insertSectionBreakNextPage':
    case 'insertSectionBreakContinuous':
      return { gate: bodyGate(env) ?? suggestingGate(env) };
    case 'insertTOC': {
      const gate = writeGate(env);
      return gate
        ? { gate }
        : { gate: 'unsupported-command', reasonKey: 'commands.reasons.tableOfContentsUnsupported' };
    }
    case 'imageWrap':
    case 'imageTransform':
    case 'imageProperties': {
      const gate =
        writeGate(env) ??
        (env.bodyStory ? null : 'unsupported-story') ??
        (env.image ? null : 'image-required') ??
        suggestingGate(env);
      if (id !== 'imageWrap') return { gate };
      const wrapArgs = args as DocxCommandArgs['imageWrap'] | undefined;
      return result<'imageWrap'>({
        gate,
        value: env.image?.wrap ?? null,
        options: WRAP_OPTIONS.map((option) => ({
          args: { wrap: option.wrap },
          label: t(option.labelKey),
        })),
        active: wrapArgs ? env.image?.wrap === wrapArgs.wrap : undefined,
      });
    }
    case 'tableAction': {
      const table = env.table;
      const border = table?.cellBorderColor?.rgb;
      return result<'tableAction'>({
        gate: evaluateTableAction(env, args as DocxTableAction | undefined),
        value:
          args === undefined
            ? table?.isInTable
              ? {
                  borderColor: typeof border === 'string' ? border.replace(/^#/, '') : null,
                  fillColor: table.cellBackgroundColor ?? null,
                  justification: table.table?.attrs?.justification ?? null,
                }
              : null
            : undefined,
      });
    }
    case 'pageSetup':
    case 'watermark':
      return { gate: writeGate(env) ?? suggestingGate(env) };
    case 'editingMode':
      return result(evaluateEditingMode(env, args as DocxCommandArgs['editingMode'] | undefined));
    case 'reviewAccept':
    case 'reviewReject':
      return { gate: evaluateRevision(env, args as DocxCommandArgs['reviewAccept'] | undefined) };
    case 'reviewPrevious':
    case 'reviewNext':
      return { gate: documentGate(env) ?? (env.revisions.length > 0 ? null : 'no-revisions') };
    case 'commentsSidebar':
      return {
        gate: documentGate(env) ?? (env.sidebarControl === 'fixed' ? 'controlled-sidebar' : null),
        active: env.sidebarOpen,
      };
    case 'open':
      return {
        gate: env.status === 'unavailable' ? 'editor-unavailable' : env.canOpen ? null : 'host-disabled',
      };
    case 'reportIssue':
      return {
        gate:
          env.status === 'unavailable'
            ? 'editor-unavailable'
            : env.canReportIssue
              ? null
              : 'host-disabled',
      };
    case 'save':
    case 'print':
    case 'find':
      return { gate: documentGate(env) };
    case 'replace':
      return { gate: writeGate(env) };
    case 'zoom': {
      const zoomArgs = args as DocxCommandArgs['zoom'] | undefined;
      return result<'zoom'>({
        gate: documentGate(env),
        value: env.zoom,
        options: ZOOM_LEVELS.map((scale) => ({
          args: { scale },
          label: `${Math.round(scale * 100)}%`,
        })),
        active: zoomArgs ? Math.abs(env.zoom - zoomArgs.scale) < 0.001 : undefined,
      });
    }
    default:
      return { gate: 'unsupported-command' };
  }
}

/** The editor's gate for a contributed command; null when it passes. */
export function contributedCommandGate(
  mutatesDocument: boolean,
  env: DocxCommandEnvironment | null
): CommandReason<DocxCommandDisabledCode> | null {
  if (!env) return commandReason('editor-unavailable', null);
  const gate = mutatesDocument ? writeGate(env) : documentGate(env);
  return gate ? commandReason(gate, env) : null;
}

/** The single availability gate for snapshots and execution. */
export function evaluateDocxCommand<K extends DocxCommandId>(
  id: K,
  args: DocxCommandArgs[K] | undefined,
  env: DocxCommandEnvironment | null
): DocxCommandState<K> {
  if (!env) {
    return { enabled: false, disabledReason: commandReason('editor-unavailable', null) };
  }
  const evaluation: Evaluation<K> = validArgs(id, args)
    ? evaluate(id, args, env)
    : { gate: documentGate(env) ?? 'invalid-arguments' };
  const presentation: {
    active?: boolean | 'mixed';
    value?: DocxCommandValues[K];
    options?: readonly DocxCommandOption<K>[];
  } = {};
  if (evaluation.active !== undefined) presentation.active = evaluation.active;
  if (evaluation.value !== undefined) presentation.value = evaluation.value;
  if (evaluation.options !== undefined) presentation.options = evaluation.options;
  return evaluation.gate
    ? {
        ...presentation,
        enabled: false,
        disabledReason: commandReason(evaluation.gate, env, evaluation.reasonKey),
      }
    : { ...presentation, enabled: true };
}
