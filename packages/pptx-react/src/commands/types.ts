import type { ParagraphAlignment, PptxEditRefusal } from '@betteroffice/pptx';
import type { TranslationKey } from '@betteroffice/pptx-i18n';
import type { CommandReason, CommandState } from '../../../../shared/host-contracts/commands';
import type { PluginCommandResult } from '../../../../shared/host-contracts/plugins';
import type { PptxEditorTool, PptxZoom } from '../components/toolbarTypes';

export type {
  CommandReason,
  CommandState,
  JsonValue,
} from '../../../../shared/host-contracts/commands';

/** A paint-order move of the selected object among its siblings. */
export type PptxZOrderMove = 'front' | 'forward' | 'backward' | 'back';

/** Arguments of every PPTX editor command, keyed by command id. */
export interface PptxCommandArgs {
  bold: null;
  italic: null;
  underline: null;
  fontFamily: { family: string };
  /** Size in points, 1 to 400. */
  fontSize: { points: number };
  fontSizeStep: { direction: 'increase' | 'decrease' };
  /** `#rrggbb`. */
  textColor: { color: string };
  alignment: { value: ParagraphAlignment };
  /** Omit the layout to reuse the current slide's; `null` inserts without one. */
  insertSlide: { layoutPartPath?: string | null };
  insertImage: null;
  /** Selects, or arms text-box or `shape:<preset>` placement on the canvas. */
  tool: { value: PptxEditorTool };
  /** `#rrggbb`; `null` removes the fill. */
  shapeFill: { color: string | null };
  /** `#rrggbb`; `null` removes the outline. */
  shapeStrokeColor: { color: string | null };
  /** Outline width in points; `null` removes the outline. */
  shapeStrokeWidth: { points: number | null };
  /** An adjustment of the selected shape, as a fraction of its range. */
  shapeAdjustment: { name: string; value: number };
  zOrder: { value: PptxZOrderMove };
  save: null;
  exportPng: null;
  slideshow: null;
  undo: null;
  redo: null;
  /** 1 is 100%, from 0.25 to 4; `'fit'` fits the slide to the window. */
  zoom: { scale: number | 'fit' };
  /** `null` toggles the panel. */
  proposalsPanel: { open: boolean } | null;
  proposalSelect: { proposalId: string };
  /** Shows the selected proposal's changes on the canvas, or the editable slide. */
  proposalDiff: { enabled: boolean };
  /** `force` applies a proposal whose targets changed after it was made. */
  proposalAccept: { proposalId: string; force?: boolean };
  proposalReject: { proposalId: string };
}

export type PptxCommandId = keyof PptxCommandArgs;

/** Commands presented as a choice between options. */
export type PptxSelectCommandId =
  | 'fontFamily'
  | 'fontSize'
  | 'alignment'
  | 'insertSlide'
  | 'tool'
  | 'shapeStrokeWidth'
  | 'shapeAdjustment'
  | 'zoom'
  | 'proposalSelect';

/** Current values reported in {@link PptxCommandState.value}, keyed by command id. */
export interface PptxCommandValues {
  bold: null;
  italic: null;
  underline: null;
  /** `null` when the selection mixes families. */
  fontFamily: string | null;
  /** Points; `null` when the selection mixes sizes. */
  fontSize: number | null;
  fontSizeStep: null;
  /** `#rrggbb`; `null` when the selection mixes colors. */
  textColor: string | null;
  /** `null` when the selected paragraphs differ. */
  alignment: ParagraphAlignment | null;
  /** The current slide's layout part. */
  insertSlide: string | null;
  insertImage: null;
  tool: PptxEditorTool;
  /** `#rrggbb`; `null` without a fill. */
  shapeFill: string | null;
  /** `#rrggbb`; `null` without an outline. */
  shapeStrokeColor: string | null;
  /** Points; `null` without an outline. */
  shapeStrokeWidth: number | null;
  /** The shape's primary adjustment; `null` when it has none. */
  shapeAdjustment: { name: string; value: number } | null;
  zOrder: null;
  save: null;
  exportPng: null;
  slideshow: null;
  undo: null;
  redo: null;
  zoom: PptxZoom;
  /** Whether the proposals panel is open. */
  proposalsPanel: boolean;
  /** The proposal shown on the canvas. */
  proposalSelect: string | null;
  /** Whether the canvas shows the selected proposal's changes. */
  proposalDiff: boolean;
  proposalAccept: null;
  proposalReject: null;
}

/** Why a command is unavailable. */
export type PptxCommandDisabledCode =
  | 'editor-unavailable'
  | 'document-loading'
  | 'no-document'
  | 'read-only'
  | 'nothing-to-undo'
  | 'nothing-to-redo'
  | 'text-selection-required'
  | 'shape-required'
  | 'slide-required'
  | 'unsupported-selection'
  | 'invalid-layout'
  | 'invalid-adjustment'
  | 'z-order-boundary'
  | 'review-active'
  /** The canvas diff of a proposal has not painted yet; reported by the canvas review control. */
  | 'preview-pending'
  /** The canvas diff of a proposal could not be laid out; reported by the canvas review control. */
  | 'preview-failed'
  | 'proposal-stale'
  | 'proposal-not-found'
  | 'proposals-unavailable'
  | 'host-disabled'
  | 'unsupported-command'
  | 'invalid-arguments'
  | 'permission-denied'
  | 'unsupported-policy'
  | 'plugin-unavailable';

/** Why an executed command did not complete. */
export type PptxCommandFailureCode =
  | PptxCommandDisabledCode
  | 'input-failed'
  | 'document-replaced'
  | 'target-changed'
  | 'gesture-active'
  | 'command-failed'
  | 'aborted';

/** One choice of a selector command. */
export interface PptxCommandOption<K extends PptxCommandId = PptxCommandId> {
  args: PptxCommandArgs[K];
  label: string;
  /** Whether this choice is available and current, as `getState(id, args)` reports it. */
  state: CommandState<never, PptxCommandDisabledCode>;
}

/**
 * Serializable state of one PPTX command. Without arguments it describes the
 * control, including its choices; with arguments, that one choice.
 */
export type PptxCommandState<K extends PptxCommandId = PptxCommandId> = CommandState<
  PptxCommandValues[K],
  PptxCommandDisabledCode
> & {
  /** Choices a selector presents, in display order. */
  options?: readonly PptxCommandOption<K>[];
};

/** A keyboard binding; `Mod` is Cmd on macOS and Ctrl elsewhere. */
export interface PptxCommandShortcut<K extends PptxCommandId = PptxCommandId> {
  chord: string;
  args: PptxCommandArgs[K];
}

/** Static, serializable description of a command. */
export interface PptxCommandDescriptor<K extends PptxCommandId = PptxCommandId> {
  id: K;
  labelKey: TranslationKey;
  mutatesDocument: boolean;
  shortcuts: readonly PptxCommandShortcut<K>[];
}

export type PptxCommandStatus = 'executed' | 'noop' | 'opened' | 'requested';

/**
 * Outcome of {@link PptxCommandStore.execute}: `executed` changed something,
 * `noop` was valid but changed nothing, `opened` showed a picker or panel,
 * and `requested` handed the change to the host.
 */
export type PptxCommandResult =
  | { ok: true; status: PptxCommandStatus }
  | { ok: false; failure: CommandReason<PptxCommandFailureCode> };

/** A command a plugin contributes, registered as `plugin:<pluginId>/<localId>`. */
export type PptxPluginCommandId = `plugin:${string}/${string}`;

/** Static description of a contributed command. */
export interface PptxPluginCommandDescriptor {
  id: PptxPluginCommandId;
  label: string;
  mutatesDocument: boolean;
  shortcuts: readonly { chord: string; args: null }[];
}

/** State of a contributed command; the plugin chooses its own disabled codes. */
export type PptxPluginCommandState = CommandState;

/**
 * Outcome of a contributed command: a plugin-defined failure, or a refused edit batch as-is.
 *
 * @experimental The plugin API may change in minor releases.
 */
export type PptxPluginCommandResult = PluginCommandResult<PptxCommandStatus, PptxEditRefusal>;

/** The command authority of one editor, shared by built-in and host chrome. */
export interface PptxCommandStore {
  getDescriptor<K extends PptxCommandId>(id: K): PptxCommandDescriptor<K>;
  /** Null while no active plugin contributes `id`. */
  getDescriptor(id: PptxPluginCommandId): PptxPluginCommandDescriptor | null;
  /** Snapshots are stable until the state changes; pass `args` to evaluate one option. */
  getState<K extends PptxCommandId>(id: K, args?: PptxCommandArgs[K]): PptxCommandState<K>;
  getState(id: PptxPluginCommandId, args?: null): PptxPluginCommandState;
  subscribe(listener: () => void): () => void;
  /** Runs after input accepted before the call; availability is checked again first. */
  execute<K extends PptxCommandId>(id: K, args: PptxCommandArgs[K]): Promise<PptxCommandResult>;
  /** Runs a contributed command with its plugin's own clients, outside the input queue. */
  execute(id: PptxPluginCommandId, args: null): Promise<PptxPluginCommandResult>;
}
