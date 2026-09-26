import type { ComponentType } from 'react';
import type {
  PptxEditRequest,
  PptxEditResult,
  PptxFindRequest,
  PptxFindResult,
  PptxReadRequest,
  PptxReadResult,
  PptxValidationResult,
} from '@betteroffice/pptx';
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
  PptxCommandArgs,
  PptxCommandDescriptor,
  PptxCommandId,
  PptxCommandResult,
  PptxCommandState,
  PptxPluginCommandDescriptor,
  PptxPluginCommandId,
  PptxPluginCommandResult,
  PptxPluginCommandState,
} from '../commands/types';
import type { PptxPointPosition } from '../PptxEditor';

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
export type PptxPluginGrant = PluginGrant<PptxCommandId>;

export type PptxPluginFailureCode = PluginFailureCode;

/** A client call refused before it reached the presentation. */
export type PptxPluginRefusal = PluginRefusal<PptxPluginFailureCode>;

/** Versioned presentation reads. Each runs after pending input. */
export interface PptxPluginReadClient {
  version(): Promise<{ ok: true; version: string } | PptxPluginRefusal>;
  readContent(request?: PptxReadRequest): Promise<PptxReadResult | PptxPluginRefusal>;
  findText(request: PptxFindRequest): Promise<PptxFindResult | PptxPluginRefusal>;
  /** Checks a batch without granting anything; a later apply is checked again. */
  validateEdits(request: PptxEditRequest): Promise<PptxValidationResult | PptxPluginRefusal>;
}

/** Granted edit batches; the grant and `readOnly` are checked when the batch applies. */
export interface PptxPluginEditClient {
  applyEdits(request: PptxEditRequest): Promise<PptxEditResult | PptxPluginRefusal>;
}

/** The editor's commands as this plugin may use them. */
export interface PptxPluginCommandClient {
  getDescriptor<K extends PptxCommandId>(id: K): PptxCommandDescriptor<K>;
  getDescriptor(id: PptxPluginCommandId): PptxPluginCommandDescriptor | null;
  getState<K extends PptxCommandId>(
    id: K,
    args?: PptxCommandArgs[K]
  ): PptxCommandState<K> | { enabled: false; disabledReason: PptxPluginRefusal['failure'] };
  getState(id: PptxPluginCommandId, args?: null): PptxPluginCommandState;
  subscribe(listener: () => void): () => void;
  execute<K extends PptxCommandId>(
    id: K,
    args: PptxCommandArgs[K]
  ): Promise<PptxCommandResult | PptxPluginRefusal>;
  /** Only this plugin's own contributed commands. */
  execute(
    id: PptxPluginCommandId,
    args: null
  ): Promise<PptxPluginCommandResult | PptxPluginRefusal>;
}

/** A rectangle in the units of the space it is used in. */
export interface PptxPluginRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** A presented slide frame and the presentation version its pixels show. */
export interface PptxPluginLayout {
  /** Changes with every presented frame, slide or zoom; unrelated to document versions. */
  id: string;
  version: string;
  slideId: string;
  /** One-based display number. */
  slide: number;
  /** The resolved scale, including fit. */
  zoom: number;
  /** Unzoomed slide display-list width, in slide pixels. */
  width: number;
  /** Unzoomed slide display-list height, in slide pixels. */
  height: number;
}

/** The current slide and what is selected on it; ids are session-scoped. */
export type PptxPluginSelection = {
  slideId: string;
  /** One-based display number. */
  slide: number;
  target:
    | { kind: 'slide' }
    | { kind: 'shape'; shapeId: string }
    | {
        kind: 'text';
        shapeId: string;
        storyId: string;
        /** UTF-16 story offsets; `focus` may precede `anchor`. */
        anchor: number;
        focus: number;
      };
} | null;

export interface PptxPluginSnapshot {
  /** The open presentation; changes when a presentation replaces it. */
  generation: string;
  version: string;
  readOnly: boolean;
  grant: PptxPluginGrant;
  selection: PptxPluginSelection;
  /** Null while no presented frame shows `version`. */
  layout: PptxPluginLayout | null;
}

/** A hit under the pointer, with the layout it was resolved against. */
export type PptxPluginPointPosition = PptxPointPosition & { version: string; layoutId: string };

/**
 * Geometry of the presented slide. Rectangles it returns are in pixels of the unscaled overlay
 * layer; every method returns null once its layout is no longer presented.
 */
export interface PptxPluginGeometry {
  layout: PptxPluginLayout;
  /**
   * Converts a slide rectangle into overlay pixels: `slide-emu` in EMU from the slide's top left
   * corner, `slide-px` in unzoomed display-list pixels.
   */
  toOverlayRect(input: {
    space: 'slide-emu' | 'slide-px';
    rect: PptxPluginRect;
  }): PptxPluginRect | null;
  /** The rendered bounds of a shape on the presented slide, group descendants included. */
  getShapeRect(shapeId: string): PptxPluginRect | null;
  /** Client coordinates; null outside slide content or during proposal review. */
  getPositionAtPoint(clientX: number, clientY: number): PptxPluginPointPosition | null;
}

export type PptxPluginNavigationFailureCode =
  | 'stale-version'
  | 'missing-target'
  | 'ambiguous-target'
  | 'layout-unavailable'
  | 'unsupported';

export type PptxPluginNavigationResult =
  | { ok: true }
  | PptxPluginRefusal
  | { ok: false; failure: { code: PptxPluginNavigationFailureCode; message: string } };

export interface PptxPluginNavigationOptions {
  /** The version the target was read at; a newer presentation refuses with `stale-version`. */
  expectVersion: string;
  /** Moves keyboard focus to the slide; focus stays where it is by default. */
  focus?: boolean;
}

/** Navigation after pending input, against the current presentation. */
export interface PptxPluginNavigation {
  goToSlide(
    target: { slideId: string },
    options: PptxPluginNavigationOptions
  ): Promise<PptxPluginNavigationResult>;
  selectShape(
    target: { slideId: string; shapeId: string },
    options: PptxPluginNavigationOptions
  ): Promise<PptxPluginNavigationResult>;
  /** Selects `anchor` to `focus` in UTF-16 story offsets, keeping their direction. */
  selectText(
    target: { slideId: string; shapeId: string; storyId: string; anchor: number; focus: number },
    options: PptxPluginNavigationOptions
  ): Promise<PptxPluginNavigationResult>;
}

/** Lifecycle notifications describe current state; several changes may arrive as one. */
export type PptxPluginEvent =
  | { type: 'load'; generation: string; version: string; reason: PluginLoadReason }
  | { type: 'document-change'; generation: string; version: string }
  | {
      type: 'selection-change';
      generation: string;
      version: string;
      selection: PptxPluginSelection;
    }
  | { type: 'mode-change'; generation: string; readOnly: boolean }
  | { type: 'layout-change'; generation: string; layout: PptxPluginLayout | null }
  | { type: 'grants-change'; generation: string; grant: PptxPluginGrant };

/** What a hook, renderer or action receives. Clients refuse once it is superseded or ended. */
export interface PptxPluginContext<S> {
  pluginId: string;
  snapshot: PptxPluginSnapshot;
  readonly state: Readonly<S>;
  /** Aborted when this hook is superseded or the plugin ends. */
  signal: AbortSignal;
  /** Aborted when this activation ends. */
  lifetimeSignal: AbortSignal;
  read: PptxPluginReadClient;
  commands: PptxPluginCommandClient;
  /** Null without an edit batch grant. */
  edits: PptxPluginEditClient | null;
  /** Null while no presented frame shows the current version. */
  geometry: PptxPluginGeometry | null;
  navigation: PptxPluginNavigation;
  /**
   * Replaces the plugin's state. Returns false for a superseded context, an ended plugin, or when
   * the presentation is no longer at `atVersion`, which defaults to `snapshot.version`.
   */
  setState(next: S | ((previous: S) => S), atVersion?: string): boolean;
  /** Registers a disposer; after the plugin ended it runs at once. */
  onCleanup(cleanup: (reason: PluginCleanupReason) => MaybePromise<void>): void;
  /** Runs an event-handler action with a fresh context, isolating its failure. */
  run(action: (context: PptxPluginContext<S>) => MaybePromise<void>): Promise<void>;
}

export interface PptxPluginPanel<S> {
  title: string;
  placement: 'left' | 'right' | 'bottom';
  /** Width for side panels, height for the bottom panel, in CSS pixels. */
  preferredSize?: number;
  defaultCollapsed?: boolean;
  render: ComponentType<{ context: PptxPluginContext<S>; width: number; height: number }>;
}

/** A command registered as `plugin:<pluginId>/<id>`. */
export interface PptxPluginCommand<S> {
  id: string;
  label: string;
  /** Disables the command while read-only; grants nothing. */
  mutatesDocument: boolean;
  /** Chords such as `Mod+Shift+R`; built-in shortcuts take precedence. */
  shortcuts?: readonly string[];
  getState?(context: PptxPluginContext<S>): CommandState;
  execute(context: PptxPluginContext<S>): MaybePromise<PptxPluginCommandResult>;
}

/**
 * A plugin. `createState`, renderers and command `getState` must be pure; `initialize`,
 * `onEvent`, command `execute` and `context.run` are managed effects.
 */
export interface PptxPluginDefinition<S> {
  id: string;
  /** Changing it restarts the plugin; other definition changes keep its state. */
  revision?: string | number;
  createState(): S;
  initialize?(context: PptxPluginContext<S>): MaybePromise<void>;
  onEvent?(context: PptxPluginContext<S>, event: PptxPluginEvent): MaybePromise<void>;
  panel?: PptxPluginPanel<S>;
  overlay?: ComponentType<{ context: PptxPluginContext<S>; geometry: PptxPluginGeometry }>;
  commands?: readonly PptxPluginCommand<S>[];
  /** Local command ids, in display order. */
  toolbar?: readonly string[];
}

declare const pptxPluginBrand: unique symbol;

/** An installable plugin, created by `definePptxPlugin`. */
export interface PptxPlugin {
  readonly id: string;
  readonly revision?: string | number;
  readonly [pptxPluginBrand]: true;
}

export type PptxPluginErrorPhase = PluginErrorPhase;

/** A plugin failure; the failing activation has already been stopped and cleaned up. */
export type PptxPluginError = PluginError<PptxPluginErrorPhase>;

export interface PptxEditorPluginProps {
  plugins?: readonly PptxPlugin[];
  /** Keyed by plugin id; omitted plugins read, validate and navigate only. */
  pluginGrants?: Readonly<Record<string, PptxPluginGrant>>;
  onPluginError?(error: PptxPluginError): void;
}
