import type { ComponentType } from 'react';
import type { RenderedDomContext } from '@betteroffice/docx/plugin-api';
import type {
  DocxEditRequest,
  DocxEditResult,
  DocxFindTextRequest,
  DocxFindTextResult,
  DocxReadParagraphsRequest,
  DocxReadParagraphsResult,
  DocxValidationResult,
} from '@betteroffice/docx/yrs';
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
  DocxCommandArgs,
  DocxCommandDescriptor,
  DocxCommandId,
  DocxCommandResult,
  DocxCommandState,
  DocxPluginCommandDescriptor,
  DocxPluginCommandId,
  DocxPluginCommandResult,
  DocxPluginCommandState,
} from '../commands/types';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';
import type { SelectionState } from '../components/DocxEditor/types';

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
 * `untrackedHistory`. Viewing mode and `readOnly` refuse every write regardless.
 */
export type DocxPluginGrant = PluginGrant<DocxCommandId>;

export type DocxPluginFailureCode = PluginFailureCode;

/** A client call refused before it reached the document. */
export type DocxPluginRefusal = PluginRefusal<DocxPluginFailureCode>;

/** Versioned document reads. Each flushes pending input first. */
export interface DocxPluginReadClient {
  version(): Promise<{ ok: true; version: string } | DocxPluginRefusal>;
  readParagraphs(
    request: DocxReadParagraphsRequest
  ): Promise<DocxReadParagraphsResult | DocxPluginRefusal>;
  findText(request: DocxFindTextRequest): Promise<DocxFindTextResult | DocxPluginRefusal>;
  /** Checks a batch without granting anything; a later apply is checked again. */
  validateEdits(request: DocxEditRequest): Promise<DocxValidationResult | DocxPluginRefusal>;
}

/** Granted edit batches; grant, mode and document policy are checked when the batch applies. */
export interface DocxPluginEditClient {
  applyEdits(request: DocxEditRequest): Promise<DocxEditResult | DocxPluginRefusal>;
}

/** The editor's commands as this plugin may use them. */
export interface DocxPluginCommandClient {
  getDescriptor<K extends DocxCommandId>(id: K): DocxCommandDescriptor<K>;
  getDescriptor(id: DocxPluginCommandId): DocxPluginCommandDescriptor | null;
  getState<K extends DocxCommandId>(
    id: K,
    args?: DocxCommandArgs[K]
  ): DocxCommandState<K> | { enabled: false; disabledReason: DocxPluginRefusal['failure'] };
  getState(id: DocxPluginCommandId, args?: null): DocxPluginCommandState;
  subscribe(listener: () => void): () => void;
  execute<K extends DocxCommandId>(
    id: K,
    args: DocxCommandArgs[K]
  ): Promise<DocxCommandResult | DocxPluginRefusal>;
  /** Only this plugin's own contributed commands. */
  execute(
    id: DocxPluginCommandId,
    args: null
  ): Promise<DocxPluginCommandResult | DocxPluginRefusal>;
}

/** A rectangle in the units of the space it is used in. */
export interface DocxPluginRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** A rendered layout and the document version its pixels show. */
export interface DocxPluginLayout {
  /** Changes with every rendered frame; unrelated to document versions. */
  id: string;
  version: string;
  zoom: number;
  pageCount: number;
}

export interface DocxPluginSelection {
  /** Formatting at the selection; its paragraph indices are not edit targets. */
  formatting: SelectionState | null;
  /** Display positions in one layout; never an edit range. */
  displayRange: { story: string; from: number; to: number; layoutId: string } | null;
}

export interface DocxPluginSnapshot {
  /** The open document; changes when a document replaces it. */
  generation: string;
  version: string;
  mode: EditorMode;
  /** Whether writes are refused: host `readOnly` or viewing mode. */
  readOnly: boolean;
  grant: DocxPluginGrant;
  selection: DocxPluginSelection;
  /** Null while no rendered layout shows `version`. */
  layout: DocxPluginLayout | null;
}

/**
 * Local geometry of the rendered layout. `dom` answers in pages-container units divided by zoom;
 * `toOverlayRect` converts one of those rectangles into overlay-layer pixels, and returns null
 * once its layout is no longer rendered.
 */
export interface DocxPluginGeometry {
  layout: DocxPluginLayout;
  /** @experimental DOM access that a data-only geometry facade may replace in a minor release. */
  dom: RenderedDomContext;
  toOverlayRect(rect: DocxPluginRect): DocxPluginRect | null;
}

export type DocxPluginNavigationFailureCode =
  | 'stale-version'
  | 'missing-target'
  | 'ambiguous-target'
  | 'layout-unavailable'
  | 'unsupported';

export interface DocxPluginNavigation {
  /**
   * Scrolls a body paragraph into view once pending input and the layout are current. Focus and
   * selection stay where they are unless `focus` is set.
   */
  scrollToParagraph(
    target: { story: string; paraId: string },
    options: { expectVersion: string; focus?: boolean }
  ): Promise<
    | { ok: true }
    | DocxPluginRefusal
    | { ok: false; failure: { code: DocxPluginNavigationFailureCode; message: string } }
  >;
}

/** Lifecycle notifications describe current state; several changes may arrive as one. */
export type DocxPluginEvent =
  | { type: 'load'; generation: string; version: string; reason: PluginLoadReason }
  | { type: 'document-change'; generation: string; version: string }
  | {
      type: 'selection-change';
      generation: string;
      version: string;
      selection: DocxPluginSelection;
    }
  | { type: 'mode-change'; generation: string; mode: EditorMode; readOnly: boolean }
  | { type: 'layout-change'; generation: string; layout: DocxPluginLayout | null }
  | { type: 'grants-change'; generation: string; grant: DocxPluginGrant };

/** What a hook, renderer or action receives. Clients refuse once it is superseded or ended. */
export interface DocxPluginContext<S> {
  pluginId: string;
  snapshot: DocxPluginSnapshot;
  readonly state: Readonly<S>;
  /** Aborted when this hook is superseded or the plugin ends. */
  signal: AbortSignal;
  /** Aborted when this activation ends. */
  lifetimeSignal: AbortSignal;
  read: DocxPluginReadClient;
  commands: DocxPluginCommandClient;
  /** Null without an edit batch grant. */
  edits: DocxPluginEditClient | null;
  /** Null while the layout is unavailable or behind the document. */
  geometry: DocxPluginGeometry | null;
  navigation: DocxPluginNavigation;
  /**
   * Replaces the plugin's state. Returns false for a superseded context, an ended plugin, or when
   * the document is no longer at `atVersion`, which defaults to `snapshot.version`.
   */
  setState(next: S | ((previous: S) => S), atVersion?: string): boolean;
  /** Registers a disposer; after the plugin ended it runs at once. */
  onCleanup(cleanup: (reason: PluginCleanupReason) => MaybePromise<void>): void;
  /** Runs an event-handler action with a fresh context, isolating its failure. */
  run(action: (context: DocxPluginContext<S>) => MaybePromise<void>): Promise<void>;
}

export interface DocxPluginPanel<S> {
  title: string;
  placement: 'left' | 'right' | 'bottom';
  /** Width for side panels, height for the bottom panel, in CSS pixels. */
  preferredSize?: number;
  defaultCollapsed?: boolean;
  render: ComponentType<{ context: DocxPluginContext<S>; width: number; height: number }>;
}

export interface DocxPluginSidebarItem<S> {
  /** Unique within the plugin. */
  id: string;
  /** Shown only while the document is at `version` and the paragraph resolves uniquely. */
  anchor: { version: string; story: string; paraId: string };
  priority?: number;
  estimatedHeight?: number;
  /**
   * Draws the card; one component can draw them all from `item`. A card keeps its React state
   * while the plugin returns its `id`, and resets for a new id or document.
   */
  render: ComponentType<{
    context: DocxPluginContext<S>;
    /** This item as `getSidebarItems` returned it. */
    item: DocxPluginSidebarItem<S>;
    isExpanded: boolean;
    onToggleExpand(): void;
    measureRef(element: HTMLDivElement | null): void;
  }>;
}

/** A command registered as `plugin:<pluginId>/<id>`. */
export interface DocxPluginCommand<S> {
  id: string;
  label: string;
  /** Disables the command in viewing and read-only modes; grants nothing. */
  mutatesDocument: boolean;
  /** Chords such as `Mod+Shift+R`; built-in shortcuts take precedence. */
  shortcuts?: readonly string[];
  getState?(context: DocxPluginContext<S>): CommandState;
  execute(context: DocxPluginContext<S>): MaybePromise<DocxPluginCommandResult>;
}

/**
 * A plugin. `createState`, renderers, `getSidebarItems` and command `getState` must be pure;
 * `initialize`, `onEvent`, command `execute` and `context.run` are managed effects.
 */
export interface DocxPluginDefinition<S> {
  id: string;
  /** Changing it restarts the plugin; other definition changes keep its state. */
  revision?: string | number;
  createState(): S;
  initialize?(context: DocxPluginContext<S>): MaybePromise<void>;
  onEvent?(context: DocxPluginContext<S>, event: DocxPluginEvent): MaybePromise<void>;
  panel?: DocxPluginPanel<S>;
  overlay?: ComponentType<{ context: DocxPluginContext<S>; geometry: DocxPluginGeometry }>;
  getSidebarItems?(context: DocxPluginContext<S>): readonly DocxPluginSidebarItem<S>[];
  commands?: readonly DocxPluginCommand<S>[];
  /** Local command ids, in display order. */
  toolbar?: readonly string[];
}

declare const docxPluginBrand: unique symbol;

/** An installable plugin, created by `defineDocxPlugin`. */
export interface DocxPlugin {
  readonly id: string;
  readonly revision?: string | number;
  readonly [docxPluginBrand]: true;
}

export type DocxPluginErrorPhase = PluginErrorPhase | 'sidebar';

/** A plugin failure; the failing activation has already been stopped and cleaned up. */
export type DocxPluginError = PluginError<DocxPluginErrorPhase>;

export interface DocxEditorPluginProps {
  plugins?: readonly DocxPlugin[];
  /** Keyed by plugin id; omitted plugins read, validate and navigate only. */
  pluginGrants?: Readonly<Record<string, DocxPluginGrant>>;
  onPluginError?(error: DocxPluginError): void;
}
