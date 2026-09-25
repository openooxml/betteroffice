import type { Translations } from '@betteroffice/pptx-i18n';
import { isPluginCommandId, isPptxCommandId, PPTX_COMMAND_DESCRIPTORS } from './descriptors';
import {
  commandReason,
  contributedCommandGate,
  evaluatePptxCommand,
  type PptxCommandEnvironment,
} from './evaluate';
import type {
  PptxCommandArgs,
  PptxCommandDisabledCode,
  PptxCommandFailureCode,
  PptxCommandId,
  PptxCommandResult,
  PptxCommandState,
  PptxCommandStore,
  PptxPluginCommandDescriptor,
  PptxPluginCommandId,
  PptxPluginCommandState,
} from './types';

/** Raised when an admitted command cannot run against its originating input or target. */
export class PptxCommandAdmissionError extends Error {
  constructor(
    readonly code: Extract<
      PptxCommandFailureCode,
      | 'input-failed'
      | 'document-replaced'
      | 'target-changed'
      | 'gesture-active'
      | 'editor-unavailable'
    >
  ) {
    super(code);
  }
}

/** A binding's record of the document and target a command was admitted for. */
export interface PptxCommandOrigin {
  readonly generation: number;
}

/** An action a command opened, such as a picker, finished by a later submission. */
export interface PptxDeferredCommand {
  /**
   * Runs `write` after input accepted before this call, once the command's gate
   * passes again and the document the action opened for is current.
   */
  complete(write: () => PptxCommandResult | Promise<PptxCommandResult>): Promise<PptxCommandResult>;
}

/** Editor-owned implementation behind a command store. */
export interface PptxCommandBinding {
  /** Gate inputs, read live; `executing` means preceding input has been applied. */
  environment(executing: boolean): PptxCommandEnvironment;
  /** Whether `id` must wait behind accepted input. */
  ordered<K extends PptxCommandId>(id: K, args: PptxCommandArgs[K]): boolean;
  /** Runs `operation` after input accepted before this call. */
  admit<T>(operation: () => T | Promise<T>): Promise<T>;
  /** Performs a command whose gate passed against `env`. */
  perform<K extends PptxCommandId>(
    id: K,
    args: PptxCommandArgs[K],
    env: PptxCommandEnvironment
  ): PptxCommandResult | Promise<PptxCommandResult>;
  /** The current document and the target `id` applies to; null without a document. */
  capture(id: PptxCommandId): PptxCommandOrigin | null;
  /** Why `origin` no longer names the live document and target, or null when it does. */
  resume(origin: PptxCommandOrigin): PptxCommandFailureCode | null;
  /** Presentation context for chrome rendered outside the editor. */
  chrome(): PptxChromeContext;
  /** Returns focus to the slide after pointer activation. */
  focusEditor(): void;
}

export interface PptxChromeContext {
  i18n: Translations | undefined;
}

/** A command whose arguments are chosen later, such as in a prompt, bound to what it opened for. */
export interface PptxPendingCommand<K extends PptxCommandId> {
  /** Runs like `execute` once the document and target the command opened for are current. */
  execute(args: PptxCommandArgs[K]): Promise<PptxCommandResult>;
}

/** A contributed command as the store runs it. */
export interface PptxPluginCommandBinding {
  readonly descriptor: PptxPluginCommandDescriptor;
  /** The plugin's own state; the editor's gate applies on top. */
  state(): PptxPluginCommandState;
  /** Runs the plugin's handler with that plugin's clients. */
  execute(): Promise<PptxCommandResult>;
}

/** Who a scoped store acts for; asked again inside every operation boundary. */
export interface PptxCommandScope {
  /** Why the caller may not run `id` now, or null. */
  deny(id: PptxCommandId | PptxPluginCommandId): PptxCommandFailureCode | null;
  /** Notified whenever `deny` may answer differently. */
  subscribe(listener: () => void): () => void;
}

/** Private controls of a store: binding, refresh, deferred actions, chrome. */
export interface PptxCommandController {
  readonly store: PptxCommandStore;
  attach(binding: PptxCommandBinding): void;
  detach(binding: PptxCommandBinding): void;
  /** Re-evaluates cached snapshots and chrome, notifying the listeners of what changed. */
  refresh(): void;
  /** Keeps one snapshot cached, and its identity stable, while a subscriber shows it. */
  hold<K extends PptxCommandId>(id: K, args?: PptxCommandArgs[K]): () => void;
  hold(id: PptxPluginCommandId): () => void;
  /** Captures the document and target of an action `id` opens now and completes later. */
  defer<K extends PptxCommandId>(id: K, args: PptxCommandArgs[K]): PptxDeferredCommand;
  /** Captures the document and target now for `id`, run later with arguments chosen then. */
  prepare<K extends PptxCommandId>(id: K): PptxPendingCommand<K>;
  /** Marks host chrome whose keyboard shortcuts belong to this editor. */
  registerChrome(element: HTMLElement): () => void;
  ownsChrome(target: EventTarget | null): boolean;
  /** Stable until the editor's locale changes. */
  chrome(): PptxChromeContext | null;
  subscribeChrome(listener: () => void): () => void;
  focusEditor(): void;
  /** Replaces the contributed commands and the toolbar order of their ids. */
  setPluginCommands(
    commands: readonly PptxPluginCommandBinding[],
    toolbar: readonly PptxPluginCommandId[]
  ): void;
  /** Contributed toolbar commands, in display order. */
  pluginToolbar(): readonly PptxPluginCommandId[];
  /** Keyboard bindings of the contributed commands. */
  pluginShortcuts(): readonly { id: PptxPluginCommandId; chord: string }[];
  /**
   * A store acting for `scope`: its answer is checked inside each operation, and mutating
   * built-in commands, which have no authoritative policy path yet, refuse with
   * `unsupported-policy`.
   */
  scoped(scope: PptxCommandScope): PptxCommandStore;
  /** Whether `store` is this editor's store or one scoped from it. */
  ownsStore(store: PptxCommandStore): boolean;
}

const MAX_CACHED_SNAPSHOTS = 512;

type AnyCommandId = PptxCommandId | PptxPluginCommandId;

interface CachedSnapshot {
  id: AnyCommandId;
  args: unknown;
  json: string;
  state: PptxCommandState | PptxPluginCommandState;
  holds: number;
  read: number;
}

const controllers = new WeakMap<PptxCommandStore, PptxCommandController>();

function failure(
  code: PptxCommandFailureCode,
  env: PptxCommandEnvironment | null
): PptxCommandResult {
  return { ok: false, failure: commandReason(code, env) };
}

function snapshotKey(id: AnyCommandId, args: unknown): string {
  return args === undefined ? id : `${id}\u0000${JSON.stringify(args)}`;
}

/** The disabled code a refusal shows as while the command cannot run at all. */
function disabledCode(code: PptxCommandFailureCode): PptxCommandDisabledCode {
  switch (code) {
    case 'input-failed':
    case 'target-changed':
    case 'gesture-active':
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
  return !isPluginCommandId(id) && PPTX_COMMAND_DESCRIPTORS[id].mutatesDocument;
}

function knownCommand(id: unknown): id is AnyCommandId {
  return isPptxCommandId(id) || isPluginCommandId(id);
}

/** Creates the command store of one editor mount. */
export function createPptxCommandController(): PptxCommandController {
  let binding: PptxCommandBinding | null = null;
  let pluginCommands = new Map<PptxPluginCommandId, PptxPluginCommandBinding>();
  let pluginToolbarIds: readonly PptxPluginCommandId[] = Object.freeze([]);
  let pluginShortcutList: readonly { id: PptxPluginCommandId; chord: string }[] = Object.freeze([]);
  const scopedStores = new WeakMap<PptxCommandScope, PptxCommandStore>();
  const ownStores = new WeakSet<PptxCommandStore>();
  const listeners = new Set<() => void>();
  const snapshots = new Map<string, CachedSnapshot>();
  const chromeRoots = new Set<HTMLElement>();
  const chromeListeners = new Set<() => void>();
  let chromeSnapshot: PptxChromeContext | null = null;
  let round = 0;

  const environment = (executing: boolean): PptxCommandEnvironment | null => {
    if (!binding) return null;
    try {
      return binding.environment(executing);
    } catch (error) {
      console.error('[pptx commands] reading editor state failed', error);
      return null;
    }
  };

  const pluginState = (
    id: PptxPluginCommandId,
    env: PptxCommandEnvironment | null
  ): PptxPluginCommandState => {
    const command = pluginCommands.get(id);
    if (!command) {
      return { enabled: false, disabledReason: commandReason('unsupported-command', env) };
    }
    const gate = contributedCommandGate(command.descriptor.mutatesDocument, env);
    return gate ? { enabled: false, disabledReason: gate } : command.state();
  };

  const compute = (id: AnyCommandId, args: unknown, env: PptxCommandEnvironment | null) =>
    isPluginCommandId(id)
      ? pluginState(id, env)
      : (evaluatePptxCommand(id, args as never, env) as PptxCommandState);

  const notify = () => {
    round += 1;
    for (const listener of [...listeners]) listener();
  };

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
      const state = knownCommand(id)
        ? compute(id, args, environment(false))
        : ({
            enabled: false,
            disabledReason: commandReason('unsupported-command', environment(false)),
          } as PptxCommandState);
      entry = { id, args, json: JSON.stringify(state), state, holds: 0, read: round };
      snapshots.set(key, entry);
      trim();
    }
    entry.read = round;
    return entry;
  };

  const refreshChrome = () => {
    let next: PptxChromeContext | null = null;
    try {
      next = binding?.chrome() ?? null;
    } catch (error) {
      console.error('[pptx commands] reading editor chrome failed', error);
    }
    const same =
      next === chromeSnapshot ||
      (next !== null && chromeSnapshot !== null && next.i18n === chromeSnapshot.i18n);
    if (same) return;
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
      if (!knownCommand(entry.id)) continue;
      const next = compute(entry.id, entry.args, env);
      const json = JSON.stringify(next);
      if (json === entry.json) continue;
      entry.state = next;
      entry.json = json;
      changed = true;
    }
    if (changed) notify();
  };

  const capture = (id: PptxCommandId): PptxCommandOrigin | null => {
    try {
      return binding?.capture(id) ?? null;
    } catch (error) {
      console.error('[pptx commands] capturing a command target failed', error);
      return null;
    }
  };

  const run = async (
    id: AnyCommandId,
    args: unknown,
    perform: (env: PptxCommandEnvironment) => PptxCommandResult | Promise<PptxCommandResult>,
    deferred?: { origin: PptxCommandOrigin | null },
    scope?: PptxCommandScope
  ): Promise<PptxCommandResult> => {
    const current = binding;
    if (!current) return failure('editor-unavailable', null);
    const builtIn = isPluginCommandId(id) ? null : id;
    const ordered =
      deferred !== undefined || (builtIn !== null && current.ordered(builtIn, args as never));
    const origin = deferred ? deferred.origin : ordered && builtIn ? capture(builtIn) : undefined;
    const attempt = () => {
      if (origin !== undefined) {
        const stale = origin ? current.resume(origin) : 'editor-unavailable';
        if (stale) return failure(stale, environment(true));
      }
      const denied = scope?.deny(id) ?? null;
      if (denied) return failure(denied, environment(true));
      const env = environment(true);
      const state = compute(id, args, env);
      if (!state.enabled) return { ok: false, failure: state.disabledReason } as PptxCommandResult;
      if (scope && mutatingBuiltIn(id)) return failure('unsupported-policy', env);
      return perform(env!);
    };
    try {
      return await (ordered ? current.admit(attempt) : attempt());
    } catch (error) {
      if (error instanceof PptxCommandAdmissionError)
        return failure(error.code, environment(false));
      console.error(`[pptx commands] ${id} failed`, error);
      return failure('command-failed', environment(false));
    } finally {
      refresh();
    }
  };

  const perform = (id: AnyCommandId, args: unknown, env: PptxCommandEnvironment) =>
    isPluginCommandId(id)
      ? pluginCommands.get(id)!.execute()
      : binding!.perform(id, args as never, env);

  const dispatch = (
    id: unknown,
    args: unknown,
    scope?: PptxCommandScope
  ): Promise<PptxCommandResult> => {
    if (!knownCommand(id)) {
      return Promise.resolve(failure('unsupported-command', environment(false)));
    }
    return run(id, args, (env) => perform(id, args, env), undefined, scope);
  };

  const getDescriptor = (id: AnyCommandId) =>
    isPluginCommandId(id)
      ? pluginCommands.get(id)?.descriptor ?? null
      : PPTX_COMMAND_DESCRIPTORS[id];

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
  }) as PptxCommandStore;

  const createScoped = (scope: PptxCommandScope): PptxCommandStore => {
    const cache = new Map<
      string,
      { base: unknown; denial: PptxCommandFailureCode | null; policy: boolean; state: unknown }
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
    }) as PptxCommandStore;
    controllers.set(scopedStore, {
      ...controller,
      store: scopedStore,
      defer(id, args) {
        const origin = capture(id);
        return { complete: (write) => run(id, args, () => write(), { origin }, scope) };
      },
      prepare(id) {
        const origin = capture(id);
        return {
          execute: (args) => run(id, args, (env) => perform(id, args, env), { origin }, scope),
        };
      },
    });
    return scopedStore;
  };

  const controller: PptxCommandController = {
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
    defer(id, args) {
      const origin = capture(id);
      return { complete: (write) => run(id, args, () => write(), { origin }) };
    },
    prepare(id) {
      const origin = capture(id);
      return {
        execute: (args) => run(id, args, (env) => perform(id, args, env), { origin }),
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
        ownStores.add(scoped);
      }
      return scoped;
    },
    ownsStore: (candidate) => ownStores.has(candidate),
  };
  ownStores.add(store);
  controllers.set(store, controller);
  return controller;
}

/** Private controls of a store created by this package, if any. */
export function pptxCommandController(store: PptxCommandStore): PptxCommandController | null {
  return controllers.get(store) ?? null;
}

/** A store without an editor: every command reports `editor-unavailable`. */
export const UNAVAILABLE_PPTX_COMMANDS: PptxCommandStore = createPptxCommandController().store;
