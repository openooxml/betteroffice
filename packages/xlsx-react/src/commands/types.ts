import type { XlsxEditRefusal } from '@betteroffice/xlsx';
import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import type {
  CommandReason,
  CommandState,
} from '../../../../shared/host-contracts/commands';
import type { PluginCommandResult } from '../../../../shared/host-contracts/plugins';
import type {
  BorderPreset,
  BorderStyle,
  HorizontalAlignment,
  MergeAction,
  NumberFormat,
  TextWrapping,
  VerticalAlignment,
} from '../components/Toolbar';

export type {
  CommandReason,
  CommandState,
  JsonValue,
} from '../../../../shared/host-contracts/commands';

/**
 * Arguments of every XLSX editor command, keyed by command id.
 * @experimental
 */
export interface XlsxCommandArgs {
  bold: null;
  italic: null;
  strikethrough: null;
  /** Captures the selection's formatting; the next selection receives it. */
  paintFormat: null;
  fontFamily: { family: string };
  /** Size in points, 1 to 400. */
  fontSize: { points: number };
  fontSizeStep: { direction: 'increase' | 'decrease' };
  /** `#rrggbb`. */
  textColor: { color: string };
  /** `#rrggbb`. */
  fillColor: { color: string };
  /** `'custom'` applies the selection's custom pattern, or `0.00`. */
  numberFormat: { value: NumberFormat };
  decimalPlaces: { direction: 'increase' | 'decrease' };
  /** Applies the preset with the current border style and color. */
  borderPreset: { value: BorderPreset };
  borderStyle: { value: BorderStyle };
  /** `#rrggbb`. */
  borderColor: { color: string };
  merge: { value: MergeAction };
  horizontalAlignment: { value: HorizontalAlignment };
  verticalAlignment: { value: VerticalAlignment };
  textWrapping: { value: TextWrapping };
  searchMenus: null;
  save: null;
  exportPng: null;
  print: null;
  undo: null;
  redo: null;
  /** 1 is 100%; 0.25 to 4. */
  zoom: { scale: number };
  /** `null` toggles the panel. */
  proposalsPanel: { open: boolean } | null;
  /** `force` applies a proposal whose cells changed since it was staged. */
  proposalAccept: { proposalId: string; force?: boolean };
  proposalReject: { proposalId: string };
}

/** @experimental */
export type XlsxCommandId = keyof XlsxCommandArgs;

/**
 * Commands presented as a choice between options.
 * @experimental
 */
export type XlsxSelectCommandId =
  | 'fontFamily'
  | 'fontSize'
  | 'numberFormat'
  | 'borderPreset'
  | 'borderStyle'
  | 'merge'
  | 'horizontalAlignment'
  | 'verticalAlignment'
  | 'textWrapping'
  | 'zoom';

/**
 * Number format of the selection; `kind` is `null` when the cells disagree.
 * @experimental
 */
export type XlsxNumberFormatValue = {
  kind: NumberFormat | null;
  pattern: string | null;
};

/**
 * Current values reported in {@link XlsxCommandState.value}, keyed by command id.
 * @experimental
 */
export interface XlsxCommandValues {
  bold: null;
  italic: null;
  strikethrough: null;
  paintFormat: null;
  fontFamily: string | null;
  /** Points; `null` when mixed or unknown. */
  fontSize: number | null;
  fontSizeStep: null;
  /** `#rrggbb`; `null` when mixed or unknown. */
  textColor: string | null;
  fillColor: string | null;
  numberFormat: XlsxNumberFormatValue | null;
  decimalPlaces: null;
  borderPreset: BorderPreset | null;
  /** The style the next border preset applies. */
  borderStyle: BorderStyle | null;
  /** The color the next border preset applies. */
  borderColor: string | null;
  /** Size of the selected range. */
  merge: { rows: number; columns: number } | null;
  horizontalAlignment: HorizontalAlignment | null;
  verticalAlignment: VerticalAlignment | null;
  textWrapping: TextWrapping | null;
  searchMenus: null;
  save: null;
  exportPng: null;
  print: null;
  undo: null;
  redo: null;
  zoom: number;
  /** Number of pending proposals. */
  proposalsPanel: number;
  proposalAccept: null;
  proposalReject: null;
}

/**
 * Why a command is unavailable.
 * @experimental
 */
export type XlsxCommandDisabledCode =
  | 'editor-unavailable'
  | 'document-loading'
  | 'no-document'
  | 'read-only'
  | 'cell-selection-required'
  | 'unsupported-selection'
  | 'multiple-cells-required'
  | 'multiple-columns-required'
  | 'multiple-rows-required'
  | 'no-merged-cells'
  | 'collaboration-unsupported'
  | 'nothing-to-undo'
  | 'nothing-to-redo'
  | 'png-unavailable'
  | 'no-proposals'
  | 'proposal-not-found'
  | 'host-disabled'
  | 'unsupported-command'
  | 'invalid-arguments'
  | 'permission-denied'
  | 'unsupported-policy'
  | 'plugin-unavailable';

/**
 * Why an executed command did not complete.
 * @experimental
 */
export type XlsxCommandFailureCode =
  | XlsxCommandDisabledCode
  | 'input-failed'
  | 'document-replaced'
  | 'target-changed'
  | 'gesture-active'
  | 'proposal-stale'
  | 'render-failed'
  | 'command-failed'
  | 'aborted';

/**
 * One choice of a selector command.
 * @experimental
 */
export interface XlsxCommandOption<K extends XlsxCommandId = XlsxCommandId> {
  args: XlsxCommandArgs[K];
  label: string;
  /** The state of the command for these arguments, as `getState(id, args)` reports it. */
  state: CommandState<XlsxCommandValues[K], XlsxCommandDisabledCode>;
}

/**
 * Serializable state of one XLSX command, optionally for specific arguments.
 * @experimental
 */
export type XlsxCommandState<K extends XlsxCommandId = XlsxCommandId> = CommandState<
  XlsxCommandValues[K],
  XlsxCommandDisabledCode
> & {
  /** Choices a selector presents, in display order. */
  options?: readonly XlsxCommandOption<K>[];
};

/**
 * A keyboard binding; `Mod` is Cmd on macOS and Ctrl elsewhere.
 * @experimental
 */
export interface XlsxCommandShortcut<K extends XlsxCommandId = XlsxCommandId> {
  chord: string;
  args: XlsxCommandArgs[K];
}

/**
 * Static, serializable description of a command.
 * @experimental
 */
export interface XlsxCommandDescriptor<K extends XlsxCommandId = XlsxCommandId> {
  id: K;
  labelKey: TranslationKey;
  mutatesDocument: boolean;
  shortcuts: readonly XlsxCommandShortcut<K>[];
}

/** @experimental */
export type XlsxCommandStatus = 'executed' | 'noop' | 'opened' | 'requested';

/**
 * Outcome of {@link XlsxCommandStore.execute}: `executed` changed something,
 * `noop` was valid but changed nothing, `opened` showed a picker, and
 * `requested` handed the change to the host.
 * @experimental
 */
export type XlsxCommandResult =
  | { ok: true; status: XlsxCommandStatus }
  | { ok: false; failure: CommandReason<XlsxCommandFailureCode> };

/** A command a plugin contributes, registered as `plugin:<pluginId>/<localId>`. */
export type XlsxPluginCommandId = `plugin:${string}/${string}`;

/** Static description of a contributed command. */
export interface XlsxPluginCommandDescriptor {
  id: XlsxPluginCommandId;
  label: string;
  mutatesDocument: boolean;
  shortcuts: readonly { chord: string; args: null }[];
}

/** State of a contributed command; the plugin chooses its own disabled codes. */
export type XlsxPluginCommandState = CommandState;

/**
 * Outcome of a contributed command: a plugin-defined failure, or a refused edit batch as-is.
 *
 * @experimental The plugin API may change in minor releases.
 */
export type XlsxPluginCommandResult = PluginCommandResult<XlsxCommandStatus, XlsxEditRefusal>;

/**
 * The command authority of one editor, shared by built-in and host chrome.
 * @experimental
 */
export interface XlsxCommandStore {
  getDescriptor<K extends XlsxCommandId>(id: K): XlsxCommandDescriptor<K>;
  /** Null while no active plugin contributes `id`. */
  getDescriptor(id: XlsxPluginCommandId): XlsxPluginCommandDescriptor | null;
  /** Snapshots are stable until the state changes; pass `args` to evaluate one option. */
  getState<K extends XlsxCommandId>(id: K, args?: XlsxCommandArgs[K]): XlsxCommandState<K>;
  getState(id: XlsxPluginCommandId, args?: null): XlsxPluginCommandState;
  subscribe(listener: () => void): () => void;
  /** Runs after input accepted before the call; availability is checked again first. */
  execute<K extends XlsxCommandId>(id: K, args: XlsxCommandArgs[K]): Promise<XlsxCommandResult>;
  /** Runs a contributed command with its plugin's own clients, outside the input queue. */
  execute(id: XlsxPluginCommandId, args: null): Promise<XlsxPluginCommandResult>;
}
