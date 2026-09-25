/**
 * Plugin DTOs shared by the formats' editor plugin hosts. Clients, selections, anchors, layouts
 * and command maps stay format-owned; these shapes fix grants, lifecycle metadata and refusals.
 */

import type { CommandReason } from './commands';

export type MaybePromise<T> = T | Promise<T>;

/**
 * What a host lets one plugin do. Without a grant a plugin reads, validates and navigates only.
 * Built-in commands need their id listed; mutating ones also need `document: 'write'`.
 */
export type PluginGrant<CommandId extends string> = {
  commands?: readonly CommandId[];
} & (
  | {
      document?: 'read';
      editBatches?: never;
      untrackedHistory?: never;
    }
  | {
      document: 'write';
      /** Allows applying edit batches. */
      editBatches?: true;
      /** Allows batches with `history: 'none'`. */
      untrackedHistory?: true;
    }
);

/** Why a plugin's cleanups run. */
export type PluginCleanupReason =
  | 'removed'
  | 'document-replaced'
  | 'definition-replaced'
  | 'unmounted'
  | 'failed';

/** `attached`: the plugin was added while a document was open. */
export type PluginLoadReason = 'loaded' | 'replaced' | 'attached';

export type PluginFailureCode =
  | 'plugin-unavailable'
  | 'document-replaced'
  | 'aborted'
  | 'permission-denied'
  | 'read-only'
  | 'input-failed'
  | 'unsupported-policy'
  | 'plugin-failed';

/** A plugin client call refused before it reached the document. */
export interface PluginRefusal<Code extends string = PluginFailureCode> {
  ok: false;
  failure: CommandReason<Code>;
}

export type PluginErrorPhase =
  | 'definition'
  | 'initialize'
  | 'event'
  | 'render'
  | 'command-state'
  | 'command'
  | 'cleanup'
  | 'action';

/** A plugin failure reported to the host; `generation` is null outside a document. */
export interface PluginError<Phase extends string = PluginErrorPhase> {
  pluginId: string;
  generation: string | null;
  phase: Phase;
  error: unknown;
}

/** The id a contributed command is registered under. */
export type PluginCommandId = `plugin:${string}/${string}`;

export interface PluginLoadEvent {
  type: 'load';
  generation: string;
  version: string;
  reason: PluginLoadReason;
}

/** A committed document change; carries only the version it produced. */
export interface PluginDocumentChangeEvent {
  type: 'document-change';
  generation: string;
  version: string;
}

export interface PluginGrantsChangeEvent<Grant> {
  type: 'grants-change';
  generation: string;
  grant: Grant;
}
