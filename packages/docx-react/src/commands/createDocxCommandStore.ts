import type { Translations } from '@betteroffice/docx-i18n';
import type { Theme } from '@betteroffice/docx/types/document';
import { DOCX_COMMAND_DESCRIPTORS, isDocxCommandId, isPluginCommandId } from './descriptors';
import {
  commandReason,
  contributedCommandGate,
  evaluateDocxCommand,
  type DocxCommandEnvironment,
} from './evaluate';
import type {
  DocxCommandArgs,
  DocxCommandDisabledCode,
  DocxCommandFailureCode,
  DocxCommandId,
  DocxCommandResult,
  DocxCommandState,
  DocxCommandStore,
  DocxPluginCommandDescriptor,
  DocxPluginCommandId,
  DocxPluginCommandState,
} from './types';

/** Raised when an admitted command cannot run against its originating input or target. */
export class DocxCommandAdmissionError extends Error {
  constructor(
    readonly code: Extract<
      DocxCommandFailureCode,
      'input-failed' | 'document-replaced' | 'target-changed' | 'editor-unavailable'
    >
  ) {
    super(code);
  }
}

/** What a deferred action applies to: the whole document, or the selection it opened with. */
export type DocxDeferredTarget = 'document' | 'selection';

/** A binding's record of the document and target a deferred action opened for. */
export interface DocxCommandOrigin {
  readonly document: object;
}

/** An action a command opened, such as a dialog, finished by a later submission. */
export interface DocxDeferredCommand {
  /**
   * Runs `write` after input accepted before this call, once the command's gate
   * passes again and the document and target the action opened for are current.
   */
  complete(write: () => DocxCommandResult | Promise<DocxCommandResult>): Promise<DocxCommandResult>;
}

/** A command whose arguments are chosen later, such as in a prompt, bound to what it opened for. */
export interface DocxPendingCommand<K extends DocxCommandId> {
  /** Runs like `execute` once the document and target the command opened for are current. */
  execute(args: DocxCommandArgs[K]): Promise<DocxCommandResult>;
}

/** Editor-owned implementation behind a command store. */
export interface DocxCommandBinding {
  /** Gate inputs, read live; `executing` means preceding input has been applied. */
  environment(executing: boolean): DocxCommandEnvironment;
  /** Whether `id` must wait behind accepted input. */
  ordered<K extends DocxCommandId>(id: K, args: DocxCommandArgs[K]): boolean;
  /** Runs `operation` after input accepted before this call. */
  admit<T>(operation: () => T | Promise<T>): Promise<T>;
  /** Performs a command whose gate passed against `env`. */
  perform<K extends DocxCommandId>(
    id: K,
    args: DocxCommandArgs[K],
    env: DocxCommandEnvironment
  ): DocxCommandResult | Promise<DocxCommandResult>;
  /** The current document and, for `selection`, the selection in it; null without a document. */
  capture(target: DocxDeferredTarget): DocxCommandOrigin | null;
  /** Why `origin` no longer names the live document and target, or null when it does. */
  resume(origin: DocxCommandOrigin): DocxCommandFailureCode | null;
  /** Presentation context for chrome rendered outside the editor. */
  chrome(): DocxChromeContext;
  /** Returns focus to the document after pointer activation. */
  focusEditor(): void;
}

export interface DocxChromeContext {
  i18n: Translations | undefined;
  isDark: boolean;
  theme: Theme | null;
}

/** A contributed command as the store runs it. */
export interface DocxPluginCommandBinding {
  readonly descriptor: DocxPluginCommandDescriptor;
  /** The plugin's own state; the editor's gate applies on top. */
  state(): DocxPluginCommandState;
  /** Runs the plugin's handler with that plugin's clients. */
  execute(): Promise<DocxCommandResult>;
}

/** Who a scoped store acts for; asked again inside every operation boundary. */
export interface DocxCommandScope {
  /** Why the caller may not run `id` now, or null. */
  deny(id: DocxCommandId | DocxPluginCommandId): DocxCommandFailureCode | null;
  /** Notified whenever `deny` may answer differently. */
  subscribe(listener: () => void): () => void;
}

/** Private controls of a store: binding, refresh, deferred actions, chrome. */
export interface DocxCommandController {
  readonly store: DocxCommandStore;
  attach(binding: DocxCommandBinding): void;
  detach(binding: DocxCommandBinding): void;
  /** Re-evaluates cached snapshots and chrome, notifying the listeners of what changed. */
  refresh(): void;
  /** Keeps one snapshot cached, and its identity stable, while a subscriber shows it. */
  hold<K extends DocxCommandId>(id: K, args?: DocxCommandArgs[K]): () => void;
  hold(id: DocxPluginCommandId): () => void;
  /** Captures the document and target of an action `id` opens now and completes later. */
  defer<K extends DocxCommandId>(
    id: K,
    args: DocxCommandArgs[K],
    target: DocxDeferredTarget
  ): DocxDeferredCommand;
  /** Captures the document and target now for `id`, run later with arguments chosen then. */
  prepare<K extends DocxCommandId>(id: K, target: DocxDeferredTarget): DocxPendingCommand<K>;
  /** Marks host chrome whose keyboard shortcuts belong to this editor. */
  registerChrome(element: HTMLElement): () => void;
  ownsChrome(target: EventTarget | null): boolean;
  /** Stable until the editor's locale, color mode or theme changes. */
  chrome(): DocxChromeContext | null;
  subscribeChrome(listener: () => void): () => void;
  focusEditor(): void;
  /** Replaces the contributed commands and the toolbar order of their ids. */
  setPluginCommands(
    commands: readonly DocxPluginCommandBinding[],
    toolbar: readonly DocxPluginCommandId[]
  ): void;
  /** Contributed toolbar commands, in display order. */
  pluginToolbar(): readonly DocxPluginCommandId[];
  /** Keyboard bindings of the contributed commands. */
  pluginShortcuts(): readonly { id: DocxPluginCommandId; chord: string }[];
  /**
   * A store acting for `scope`: its answer is checked inside each operation, and mutating
   * built-in commands, which have no authoritative policy path yet, refuse with
   * `unsupported-policy`.
   */
  scoped(scope: DocxCommandScope): DocxCommandStore;
}

const MAX_CACHED_SNAPSHOTS = 512;

type AnyCommandId = DocxCommandId | DocxPluginCommandId;

interface CachedSnapshot {
  id: AnyCommandId;
  args: unknown;
  json: string;
  state: DocxCommandState | DocxPluginCommandState;
  /** Subscriptions that hold this snapshot. */
  holds: number;
  /** Notification round in which it was last read. */
  read: number;
}

const controllers = new WeakMap<DocxCommandStore, DocxCommandController>();

function failure(code: DocxCommandFailureCode, env: DocxCommandEnvironment | null): DocxCommandResult {
  return { ok: false, failure: commandReason(code, env) };
}

function snapshotKey(id: AnyCommandId, args: unknown): string {
  return args === undefined ? id : `${id}\u0000${JSON.stringify(args)}`;
}

function sameChrome(a: DocxChromeContext | null, b: DocxChromeContext | null): boolean {
  if (!a || !b) return a === b;
  return a.i18n === b.i18n && a.isDark === b.isDark && a.theme === b.theme;
}

/** The disabled code a refusal shows as while the command cannot run at all. */
function disabledCode(code: DocxCommandFailureCode): DocxCommandDisabledCode {
  switch (code) {
    case 'input-failed':
    case 'target-changed':
    case 'command-failed':
      return 'editor-unavailable';
    case 'document-replaced':
    case 'aborted':
      return 'plugin-unavailable';
    default:
      return code;
  }
}

function mutatingBuiltIn(id: AnyCommandId): boolean {
  return !isPluginCommandId(id) && DOCX_COMMAND_DESCRIPTORS[id].mutatesDocument;
}

/** Creates the command store of one editor mount. */
export function createDocxCommandController(): DocxCommandController {
  let binding: DocxCommandBinding | null = null;
  let pluginCommands = new Map<DocxPluginCommandId, DocxPluginCommandBinding>();
  let pluginToolbarIds: readonly DocxPluginCommandId[] = Object.freeze([]);
  let pluginShortcutList: readonly { id: DocxPluginCommandId; chord: string }[] = Object.freeze([]);
  const scopedStores = new WeakMap<DocxCommandScope, DocxCommandStore>();
  const listeners = new Set<() => void>();
  const snapshots = new Map<string, CachedSnapshot>();
  const chromeRoots = new Set<HTMLElement>();
  const chromeListeners = new Set<() => void>();
  let chromeSnapshot: DocxChromeContext | null = null;

  const environment = (executing: boolean): DocxCommandEnvironment | null => {
    if (!binding) return null;
    try {
      return binding.environment(executing);
    } catch (error) {
      console.error('[docx commands] reading editor state failed', error);
      return null;
    }
  };

  const pluginState = (
    id: DocxPluginCommandId,
    env: DocxCommandEnvironment | null
  ): DocxPluginCommandState => {
    const command = pluginCommands.get(id);
    if (!command) {
      return { enabled: false, disabledReason: commandReason('unsupported-command', env) };
    }
    const gate = contributedCommandGate(command.descriptor.mutatesDocument, env);
    return gate ? { enabled: false, disabledReason: gate } : command.state();
  };

  const compute = (id: AnyCommandId, args: unknown, env: DocxCommandEnvironment | null) =>
    isPluginCommandId(id)
      ? pluginState(id, env)
      : (evaluateDocxCommand(id, args as never, env) as DocxCommandState);

  let round = 0;

  const notify = () => {
    round += 1;
    for (const listener of [...listeners]) listener();
  };

  /** Drops snapshots nobody holds or read since the last notification. */
  const trim = () => {
    if (snapshots.size <= MAX_CACHED_SNAPSHOTS) return;
    for (const [key, entry] of snapshots) {
      if (snapshots.size <= MAX_CACHED_SNAPSHOTS) break;
      if (entry.holds === 0 && entry.read < round) snapshots.delete(key);
    }
  };

  const snapshot = (id: AnyCommandId, args: unknown): CachedSnapshot => {
    const key = snapshotKey(id, args);
    let entry = snapshots.get(key);
    if (!entry) {
      const state =
        isDocxCommandId(id) || isPluginCommandId(id)
          ? compute(id, args, environment(false))
          : ({
              enabled: false,
              disabledReason: commandReason('unsupported-command', environment(false)),
            } as DocxCommandState);
      entry = { id, args, json: JSON.stringify(state), state, holds: 0, read: round };
      snapshots.set(key, entry);
      trim();
    }
    entry.read = round;
    return entry;
  };

  const capture = (target: DocxDeferredTarget): DocxCommandOrigin | null => {
    try {
      return binding?.capture(target) ?? null;
    } catch (error) {
      console.error('[docx commands] capturing an action target failed', error);
      return null;
    }
  };

  const refreshChrome = () => {
    let next: DocxChromeContext | null = null;
    try {
      next = binding?.chrome() ?? null;
    } catch (error) {
      console.error('[docx commands] reading editor chrome failed', error);
    }
    if (sameChrome(next, chromeSnapshot)) return;
    chromeSnapshot = next;
    for (const listener of [...chromeListeners]) listener();
  };

  const refresh = () => {
    refreshChrome();
    trim();
    if (snapshots.size === 0) return;
    const env = environment(false);
    let changed = false;
    for (const entry of snapshots.values()) {
      const next = compute(entry.id, entry.args, env);
      const json = JSON.stringify(next);
      if (json === entry.json) continue;
      entry.state = next;
      entry.json = json;
      changed = true;
    }
    if (changed) notify();
  };

  const run = async (
    id: AnyCommandId,
    args: unknown,
    perform: (env: DocxCommandEnvironment) => DocxCommandResult | Promise<DocxCommandResult>,
    origin?: DocxCommandOrigin | null,
    scope?: DocxCommandScope
  ): Promise<DocxCommandResult> => {
    const current = binding;
    if (!current) return failure('editor-unavailable', null);
    const deferred = origin !== undefined;
    const attempt = () => {
      if (deferred) {
        const stale = origin ? current.resume(origin) : 'editor-unavailable';
        if (stale) return failure(stale, environment(true));
      }
      const denied = scope?.deny(id) ?? null;
      if (denied) return failure(denied, environment(true));
      const env = environment(true);
      const state = compute(id, args, env);
      if (!state.enabled) return { ok: false, failure: state.disabledReason } as DocxCommandResult;
      if (scope && mutatingBuiltIn(id)) return failure('unsupported-policy', env);
      return perform(env!);
    };
    const ordered = !isPluginCommandId(id) && current.ordered(id, args as never);
    try {
      return await (deferred || ordered ? current.admit(attempt) : attempt());
    } catch (error) {
      if (error instanceof DocxCommandAdmissionError) return failure(error.code, environment(false));
      console.error(`[docx commands] ${id} failed`, error);
      return failure('command-failed', environment(false));
    } finally {
      refresh();
    }
  };

  const perform = (id: AnyCommandId, args: unknown, env: DocxCommandEnvironment) =>
    isPluginCommandId(id)
      ? pluginCommands.get(id)!.execute()
      : binding!.perform(id, args as never, env);

  const dispatch = (
    id: unknown,
    args: unknown,
    scope?: DocxCommandScope
  ): Promise<DocxCommandResult> => {
    if (!isDocxCommandId(id) && !isPluginCommandId(id)) {
      return Promise.resolve(failure('unsupported-command', environment(false)));
    }
    return run(id, args, (env) => perform(id, args, env), undefined, scope);
  };

  const getDescriptor = (id: AnyCommandId) =>
    isPluginCommandId(id)
      ? (pluginCommands.get(id)?.descriptor ?? null)
      : DOCX_COMMAND_DESCRIPTORS[id];

  const store = Object.freeze({
    getDescriptor,
    getState: (id: AnyCommandId, args?: unknown) => snapshot(id, args).state,
    subscribe(listener: () => void): () => void {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    execute: (id: AnyCommandId, args: unknown) => dispatch(id, args),
  }) as DocxCommandStore;

  const createScoped = (scope: DocxCommandScope): DocxCommandStore => {
    const cache = new Map<
      string,
      { base: unknown; denial: DocxCommandFailureCode | null; policy: boolean; state: unknown }
    >();
    const getState = (id: AnyCommandId, args?: unknown) => {
      const base = snapshot(id, args).state;
      const denial = scope.deny(id);
      const policy = !denial && base.enabled && mutatingBuiltIn(id);
      const key = snapshotKey(id, args);
      const cached = cache.get(key);
      if (cached && cached.base === base && cached.denial === denial && cached.policy === policy) {
        return cached.state;
      }
      const code = denial ? disabledCode(denial) : policy ? 'unsupported-policy' : null;
      const { enabled: _enabled, disabledReason: _reason, ...presentation } = base;
      const state = code
        ? {
            ...presentation,
            enabled: false,
            disabledReason: commandReason(code, environment(false)),
          }
        : base;
      if (cache.size >= MAX_CACHED_SNAPSHOTS) cache.delete(cache.keys().next().value!);
      cache.set(key, { base, denial, policy, state });
      return state;
    };
    const scopedStore = Object.freeze({
      getDescriptor,
      getState,
      subscribe(listener: () => void): () => void {
        const release = [store.subscribe(listener), scope.subscribe(listener)];
        return () => {
          for (const unsubscribe of release) unsubscribe();
        };
      },
      execute: (id: AnyCommandId, args: unknown) => dispatch(id, args, scope),
    }) as DocxCommandStore;
    controllers.set(scopedStore, {
      ...controller,
      store: scopedStore,
      defer(id, args, target) {
        const origin = capture(target);
        return { complete: (write) => run(id, args, () => write(), origin, scope) };
      },
      prepare(id, target) {
        const origin = capture(target);
        return {
          execute: (args) => run(id, args, (env) => perform(id, args, env), origin, scope),
        };
      },
    });
    return scopedStore;
  };

  const controller: DocxCommandController = {
    store,
    attach(next) {
      binding = next;
      refresh();
    },
    detach(previous) {
      if (binding !== previous) return;
      binding = null;
      refresh();
    },
    refresh,
    hold(id: AnyCommandId, args?: unknown) {
      const entry = snapshot(id, args);
      entry.holds += 1;
      let held = true;
      return () => {
        if (!held) return;
        held = false;
        entry.holds -= 1;
      };
    },
    defer(id, args, target) {
      const origin = capture(target);
      return {
        complete: (write) => run(id, args, () => write(), origin),
      };
    },
    prepare(id, target) {
      const origin = capture(target);
      return {
        execute: (args) => {
          const current = binding;
          return run(id, args, (env) => current!.perform(id, args, env), origin);
        },
      };
    },
    registerChrome(element) {
      chromeRoots.add(element);
      return () => {
        chromeRoots.delete(element);
      };
    },
    ownsChrome(target) {
      if (!(target instanceof Node)) return false;
      for (const root of chromeRoots) if (root.contains(target)) return true;
      return false;
    },
    chrome: () => chromeSnapshot,
    subscribeChrome(listener) {
      chromeListeners.add(listener);
      return () => {
        chromeListeners.delete(listener);
      };
    },
    focusEditor() {
      binding?.focusEditor();
    },
    setPluginCommands(commands, toolbar) {
      pluginCommands = new Map(commands.map((command) => [command.descriptor.id, command]));
      pluginToolbarIds = Object.freeze(toolbar.filter((id) => pluginCommands.has(id)));
      pluginShortcutList = Object.freeze(
        commands.flatMap((command) =>
          command.descriptor.shortcuts.map(({ chord }) => ({ id: command.descriptor.id, chord }))
        )
      );
      refresh();
      notify();
    },
    pluginToolbar: () => pluginToolbarIds,
    pluginShortcuts: () => pluginShortcutList,
    scoped(scope) {
      let scoped = scopedStores.get(scope);
      if (!scoped) {
        scoped = createScoped(scope);
        scopedStores.set(scope, scoped);
      }
      return scoped;
    },
  };
  controllers.set(store, controller);
  return controller;
}

/** Private controls of a store created by this package, if any. */
export function docxCommandController(store: DocxCommandStore): DocxCommandController | null {
  return controllers.get(store) ?? null;
}

/** A store without an editor: every command reports `editor-unavailable`. */
export const UNAVAILABLE_DOCX_COMMANDS: DocxCommandStore = createDocxCommandController().store;
