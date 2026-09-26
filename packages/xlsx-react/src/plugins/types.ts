import type { ComponentType } from 'react';
import type {
  CellRange,
  Selection,
  XlsxEditRequest,
  XlsxEditResult,
  XlsxFindRequest,
  XlsxFindResult,
  XlsxReadRequest,
  XlsxReadResult,
  XlsxValidationResult,
} from '@betteroffice/xlsx';
import type { CommandState } from '../../../../shared/host-contracts/commands';
import type {
  MaybePromise,
  PluginCleanupReason,
  PluginError,
  PluginErrorPhase,
  PluginFailureCode,
  PluginGrant,
  PluginLoadReason,
  PluginRefusal,
} from '../../../../shared/host-contracts/plugins';
import type {
  XlsxCommandArgs,
  XlsxCommandDescriptor,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandState,
  XlsxPluginCommandDescriptor,
  XlsxPluginCommandId,
  XlsxPluginCommandResult,
  XlsxPluginCommandState,
} from '../commands/types';

export type {
  MaybePromise,
  PluginCleanupReason,
  PluginGrant,
  PluginLoadReason,
} from '../../../../shared/host-contracts/plugins';

/**
 * What the host lets one plugin do. Without a grant a plugin reads, validates and navigates only.
 * Built-in commands need their id listed and, when they mutate, `document: 'write'`; edit
 * batches need `document: 'write'` and `editBatches`, and `history: 'none'` also needs
 * `untrackedHistory`. `readOnly` refuses every write regardless.
 */
export type XlsxPluginGrant = PluginGrant<XlsxCommandId>;

export type XlsxPluginFailureCode = PluginFailureCode;

/** A client call refused before it reached the workbook. */
export type XlsxPluginRefusal = PluginRefusal<XlsxPluginFailureCode>;

/** Versioned workbook reads. Each runs after pending input. */
export interface XlsxPluginReadClient {
  version(): Promise<{ ok: true; version: string } | XlsxPluginRefusal>;
  readCells(request: XlsxReadRequest): Promise<XlsxReadResult | XlsxPluginRefusal>;
  findText(request: XlsxFindRequest): Promise<XlsxFindResult | XlsxPluginRefusal>;
  /** Checks a batch without granting anything; a later apply is checked again. */
  validateEdits(request: XlsxEditRequest): Promise<XlsxValidationResult | XlsxPluginRefusal>;
}

/** Granted edit batches; the grant and `readOnly` are checked when the batch applies. */
export interface XlsxPluginEditClient {
  applyEdits(request: XlsxEditRequest): Promise<XlsxEditResult | XlsxPluginRefusal>;
}

/** The editor's commands as this plugin may use them. */
export interface XlsxPluginCommandClient {
  getDescriptor<K extends XlsxCommandId>(id: K): XlsxCommandDescriptor<K>;
  getDescriptor(id: XlsxPluginCommandId): XlsxPluginCommandDescriptor | null;
  getState<K extends XlsxCommandId>(
    id: K,
    args?: XlsxCommandArgs[K]
  ): XlsxCommandState<K> | { enabled: false; disabledReason: XlsxPluginRefusal['failure'] };
  getState(id: XlsxPluginCommandId, args?: null): XlsxPluginCommandState;
  subscribe(listener: () => void): () => void;
  execute<K extends XlsxCommandId>(
    id: K,
    args: XlsxCommandArgs[K]
  ): Promise<XlsxCommandResult | XlsxPluginRefusal>;
  /** Only this plugin's own contributed commands. */
  execute(
    id: XlsxPluginCommandId,
    args: null
  ): Promise<XlsxPluginCommandResult | XlsxPluginRefusal>;
}

/** A rectangle in overlay pixels. */
export interface XlsxPluginRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** A painted grid frame and the workbook version its pixels show. */
export interface XlsxPluginLayout {
  /** Changes with every painted frame, sheet, zoom and scroll; unrelated to document versions. */
  id: string;
  version: string;
  sheetId: string;
  zoom: number;
  /** The painted part of the sheet, in unzoomed sheet pixels; `x`/`y` scroll the body. */
  viewport: { x: number; y: number; width: number; height: number };
}

/** A cell under the pointer, with the layout it was resolved against. */
export interface XlsxPluginCellPosition {
  sheetId: string;
  row: number;
  col: number;
  version: string;
  layoutId: string;
}

/**
 * Geometry of the painted grid. Rectangles are in pixels of the overlay layer, already zoomed
 * and clipped to the visible grid; every method returns null once its layout is no longer
 * painted, and for cells scrolled out of view or on another sheet.
 */
export interface XlsxPluginGeometry {
  layout: XlsxPluginLayout;
  /** One grid cell, without merge expansion; zero-based `row` and `col`. */
  getCellRect(target: { sheetId: string; row: number; col: number }): XlsxPluginRect | null;
  /** The visible part of an inclusive range, such as a merged area. */
  getRangeRect(target: { sheetId: string; range: CellRange }): XlsxPluginRect | null;
  /** Client-coordinate hit testing; null until the editor exposes it. */
  getPositionAtPoint: ((clientX: number, clientY: number) => XlsxPluginCellPosition | null) | null;
}

/** The active sheet and what is selected on it; ids are session-scoped. */
export type XlsxPluginSelection = {
  sheetId: string;
  /** Zero-based position in the workbook. */
  sheetIndex: number;
  /** Zero-based cells; `focus` may precede `anchor`. Null while nothing is selected. */
  cells: Selection | null;
  /** The selected chart, while one is selected instead of cells. */
  chartId: string | null;
} | null;

export interface XlsxPluginSnapshot {
  /** The open workbook; changes when a workbook replaces it. */
  generation: string;
  version: string;
  readOnly: boolean;
  grant: XlsxPluginGrant;
  selection: XlsxPluginSelection;
  /** Null while no painted frame shows `version`. */
  layout: XlsxPluginLayout | null;
}

export type XlsxPluginNavigationFailureCode =
  | 'stale-version'
  | 'missing-target'
  | 'ambiguous-target'
  | 'layout-unavailable'
  | 'unsupported';

export type XlsxPluginNavigationResult =
  | { ok: true }
  | XlsxPluginRefusal
  | { ok: false; failure: { code: XlsxPluginNavigationFailureCode; message: string } };

export interface XlsxPluginNavigationOptions {
  /** The version the target was read at; a newer workbook refuses with `stale-version`. */
  expectVersion: string;
}

/** Navigation after pending input, against the current workbook. It never writes. */
export interface XlsxPluginNavigation {
  /**
   * Activates the sheet, selects `selection` in its direction and reveals the focus cell.
   * Keyboard focus stays where it is unless `focus` is true.
   */
  selectCells(
    target: { sheetId: string; selection: Selection },
    options: XlsxPluginNavigationOptions & { focus?: boolean }
  ): Promise<XlsxPluginNavigationResult>;
  /**
   * Reveals a cell, `nearest` by default. The selection stays on the current sheet; on another
   * sheet, which becomes active, nothing is selected.
   */
  scrollToCell(
    target: { sheetId: string; row: number; col: number },
    options: XlsxPluginNavigationOptions & { align?: 'nearest' | 'start' | 'center' }
  ): Promise<XlsxPluginNavigationResult>;
}

/** Lifecycle notifications describe current state; several changes may arrive as one. */
export type XlsxPluginEvent =
  | { type: 'load'; generation: string; version: string; reason: PluginLoadReason }
  | { type: 'document-change'; generation: string; version: string }
  | {
      type: 'selection-change';
      generation: string;
      version: string;
      selection: XlsxPluginSelection;
    }
  | { type: 'mode-change'; generation: string; readOnly: boolean }
  | { type: 'layout-change'; generation: string; layout: XlsxPluginLayout | null }
  | { type: 'grants-change'; generation: string; grant: XlsxPluginGrant };

/** What a hook, renderer or action receives. Clients refuse once it is superseded or ended. */
export interface XlsxPluginContext<S> {
  pluginId: string;
  snapshot: XlsxPluginSnapshot;
  readonly state: Readonly<S>;
  /** Aborted when this hook is superseded or the plugin ends. */
  signal: AbortSignal;
  /** Aborted when this activation ends. */
  lifetimeSignal: AbortSignal;
  read: XlsxPluginReadClient;
  commands: XlsxPluginCommandClient;
  /** Null without an edit batch grant. */
  edits: XlsxPluginEditClient | null;
  /** Null while no painted frame shows the current version. */
  geometry: XlsxPluginGeometry | null;
  navigation: XlsxPluginNavigation;
  /**
   * Replaces the plugin's state. Returns false for a superseded context, an ended plugin, or when
   * the workbook is no longer at `atVersion`, which defaults to `snapshot.version`.
   */
  setState(next: S | ((previous: S) => S), atVersion?: string): boolean;
  /** Registers a disposer; after the plugin ended it runs at once. */
  onCleanup(cleanup: (reason: PluginCleanupReason) => MaybePromise<void>): void;
  /** Runs an event-handler action with a fresh context, isolating its failure. */
  run(action: (context: XlsxPluginContext<S>) => MaybePromise<void>): Promise<void>;
}

export interface XlsxPluginPanel<S> {
  title: string;
  placement: 'left' | 'right' | 'bottom';
  /** Width for side panels, height for the bottom panel, in CSS pixels. */
  preferredSize?: number;
  defaultCollapsed?: boolean;
  render: ComponentType<{ context: XlsxPluginContext<S>; width: number; height: number }>;
}

/** A command registered as `plugin:<pluginId>/<id>`. */
export interface XlsxPluginCommand<S> {
  id: string;
  label: string;
  /** Disables the command while read-only; grants nothing. */
  mutatesDocument: boolean;
  /**
   * Chords such as `Mod+Shift+R` or `F7`: each uses Mod or Alt, or is F1 to F12. Built-in shortcuts
   * and keys the grid handles take precedence; a clashing chord is reported and not bound.
   */
  shortcuts?: readonly string[];
  getState?(context: XlsxPluginContext<S>): CommandState;
  execute(context: XlsxPluginContext<S>): MaybePromise<XlsxPluginCommandResult>;
}

/**
 * A plugin. `createState`, renderers and command `getState` must be pure; `initialize`,
 * `onEvent`, command `execute` and `context.run` are managed effects.
 */
export interface XlsxPluginDefinition<S> {
  id: string;
  /** Changing it restarts the plugin; other definition changes keep its state. */
  revision?: string | number;
  createState(): S;
  initialize?(context: XlsxPluginContext<S>): MaybePromise<void>;
  onEvent?(context: XlsxPluginContext<S>, event: XlsxPluginEvent): MaybePromise<void>;
  panel?: XlsxPluginPanel<S>;
  overlay?: ComponentType<{ context: XlsxPluginContext<S>; geometry: XlsxPluginGeometry }>;
  commands?: readonly XlsxPluginCommand<S>[];
  /** Local command ids, in display order. */
  toolbar?: readonly string[];
}

declare const xlsxPluginBrand: unique symbol;

/** An installable plugin, created by `defineXlsxPlugin`. */
export interface XlsxPlugin {
  readonly id: string;
  readonly revision?: string | number;
  readonly [xlsxPluginBrand]: true;
}

export type XlsxPluginErrorPhase = PluginErrorPhase;

/** A plugin failure; the failing activation has already been stopped and cleaned up. */
export type XlsxPluginError = PluginError<XlsxPluginErrorPhase>;

export interface XlsxEditorPluginProps {
  plugins?: readonly XlsxPlugin[];
  /** Keyed by plugin id; omitted plugins read, validate and navigate only. */
  pluginGrants?: Readonly<Record<string, XlsxPluginGrant>>;
  onPluginError?(error: XlsxPluginError): void;
}
