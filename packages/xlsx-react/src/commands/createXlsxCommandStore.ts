import type { Translations } from '@betteroffice/xlsx-i18n';
import { isPluginCommandId, isXlsxCommandId, XLSX_COMMAND_DESCRIPTORS } from './descriptors';
import {
  commandReason,
  contributedCommandGate,
  evaluateXlsxCommand,
  type XlsxCommandEnvironment,
} from './evaluate';
import type {
  XlsxCommandArgs,
  XlsxCommandDisabledCode,
  XlsxCommandFailureCode,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandState,
  XlsxCommandStore,
  XlsxPluginCommandDescriptor,
  XlsxPluginCommandId,
  XlsxPluginCommandState,
} from './types';

/** Why work queued behind accepted input could not run. */
export type XlsxCommandAdmissionCode = Extract<
  XlsxCommandFailureCode,
  'input-failed' | 'document-replaced' | 'target-changed' | 'gesture-active' | 'editor-unavailable'
>;

const ADMISSION_MESSAGES: Record<XlsxCommandAdmissionCode, string> = {
  'input-failed': 'Input accepted earlier could not be written',
  'document-replaced': 'The workbook was replaced',
  'target-changed': 'The selection changed',
  'gesture-active': 'Finish the chart drag first',
  'editor-unavailable': 'The editor is not ready',
};

/**
 * Raised when work queued behind accepted input cannot run: a command, or a version, read or
 * edit-batch call of the editor API.
 */
export class XlsxCommandAdmissionError extends Error {
  constructor(readonly code: XlsxCommandAdmissionCode) {
    super(ADMISSION_MESSAGES[code]);
    this.name = 'XlsxCommandAdmissionError';
  }
}

/** The document and target a command was admitted or a choice opened for. */
export interface XlsxCommandOrigin {
  readonly generation: number;
  readonly target: string;
}

/** A command whose arguments are chosen later, bound to what it opened for. */
export interface XlsxPendingCommand<K extends XlsxCommandId> {
  /** Runs like `execute` once the document and selection the choice opened for are current. */
  execute(args: XlsxCommandArgs[K]): Promise<XlsxCommandResult>;
}

/** Editor-owned implementation behind a command store. */
export interface XlsxCommandBinding {
  /** Gate inputs; `executing` reads them live, after preceding input was written. */
  environment(executing: boolean): XlsxCommandEnvironment;
  /** Whether `id` must wait behind accepted input. */
  ordered(id: XlsxCommandId): boolean;
  /** Runs `operation` after input accepted before this call. */
  admit<T>(operation: () => T | Promise<T>): Promise<T>;
  /** Performs a command whose gate passed against `env`. */
  perform<K extends XlsxCommandId>(
    id: K,
    args: XlsxCommandArgs[K],
    env: XlsxCommandEnvironment
  ): XlsxCommandResult | Promise<XlsxCommandResult>;
  /** The current document and what `id` acts on in it; null without a document. */
  capture(id: XlsxCommandId): XlsxCommandOrigin | null;
  /** Why `origin` no longer names the live document and target, or null when it does. */
  resume(origin: XlsxCommandOrigin, id: XlsxCommandId): XlsxCommandFailureCode | null;
  /** Presentation context for chrome rendered outside the editor. */
  chrome(): XlsxChromeContext;
  /** Returns focus to the grid after pointer activation. */
  focusEditor(): void;
}

export interface XlsxChromeContext {
  i18n: Translations | undefined;
}

/** A contributed command as the store runs it. */
export interface XlsxPluginCommandBinding {
  readonly descriptor: XlsxPluginCommandDescriptor;
  /** The plugin's own state; the editor's gate applies on top. */
  state(): XlsxPluginCommandState;
  /** Runs the plugin's handler with that plugin's clients. */
  execute(): Promise<XlsxCommandResult>;
}

/** Who a scoped store acts for; asked again inside every operation boundary. */
export interface XlsxCommandScope {
  /** Why the caller may not run `id` now, or null. */
  deny(id: XlsxCommandId | XlsxPluginCommandId): XlsxCommandFailureCode | null;
  /** Notified whenever `deny` may answer differently. */
  subscribe(listener: () => void): () => void;
}

/** Private controls of a store: binding, refresh, deferred choices, chrome. */
export interface XlsxCommandController {
  readonly store: XlsxCommandStore;
  attach(binding: XlsxCommandBinding): void;
  detach(binding: XlsxCommandBinding): void;
  /** Re-evaluates cached snapshots and chrome, notifying the listeners of what changed. */
  refresh(): void;
  /** Keeps one snapshot cached, and its identity stable, while a subscriber shows it. */
  hold<K extends XlsxCommandId>(id: K, args?: XlsxCommandArgs[K]): () => void;
  hold(id: XlsxPluginCommandId): () => void;
  /** Captures the document and target now for `id`, run later with arguments chosen then. */
  prepare<K extends XlsxCommandId>(id: K): XlsxPendingCommand<K>;
  /** Marks host chrome whose keyboard shortcuts belong to this editor. */
  registerChrome(element: HTMLElement): () => void;
  ownsChrome(target: EventTarget | null): boolean;
  /** Stable until the editor's locale changes. */
  chrome(): XlsxChromeContext | null;
  subscribeChrome(listener: () => void): () => void;
  focusEditor(): void;
  /** Replaces the contributed commands and the toolbar order of their ids. */
  setPluginCommands(
    commands: readonly XlsxPluginCommandBinding[],
    toolbar: readonly XlsxPluginCommandId[]
  ): void;
  /** Contributed toolbar commands, in display order. */
  pluginToolbar(): readonly XlsxPluginCommandId[];
  /** Keyboard bindings of the contributed commands. */
  pluginShortcuts(): readonly { id: XlsxPluginCommandId; chord: string }[];
  /**
   * A store acting for `scope`: its answer is checked inside each operation, and mutating
   * built-in commands, which have no authoritative policy path yet, refuse with
   * `unsupported-policy`.
   */
  scoped(scope: XlsxCommandScope): XlsxCommandStore;
  /** Whether `store` is this editor's store or one scoped from it. */
  ownsStore(store: XlsxCommandStore): boolean;
}

const MAX_CACHED_SNAPSHOTS = 512;

type AnyCommandId = XlsxCommandId | XlsxPluginCommandId;

interface CachedSnapshot {
  id: AnyCommandId;
  args: unknown;
  json: string;
  state: XlsxCommandState | XlsxPluginCommandState;
  holds: number;
  /** Notification round in which it was last read. */
  read: number;
}

const controllers = new WeakMap<XlsxCommandStore, XlsxCommandController>();

function failure(code: XlsxCommandFailureCode, env: XlsxCommandEnvironment | null): XlsxCommandResult {
  return { ok: false, failure: commandReason(code, env) };
}

function snapshotKey(id: AnyCommandId, args: unknown): string {
  return args === undefined ? id : `${id}\u0000${JSON.stringify(args)}`;
}

/** The disabled code a refusal shows as while the command cannot run at all. */
function disabledCode(code: XlsxCommandFailureCode): XlsxCommandDisabledCode {
  switch (code) {
    case 'input-failed':
    case 'target-changed':
    case 'gesture-active':
    case 'proposal-stale':
    case 'render-failed':
    case 'execution-failed':
      return 'editor-unavailable';
    case 'document-replaced':
    case 'aborted':
      return 'plugin-unavailable';
    default:
      return code;
  }
}

function mutatingBuiltIn(id: AnyCommandId): boolean {
  return !isPluginCommandId(id) && XLSX_COMMAND_DESCRIPTORS[id].mutatesDocument;
}

function knownCommand(id: unknown): id is AnyCommandId {
  return isXlsxCommandId(id) || isPluginCommandId(id);
}

/** Creates the command store of one editor mount. */
export function createXlsxCommandController(): XlsxCommandController {
  let binding: XlsxCommandBinding | null = null;
  let pluginCommands = new Map<XlsxPluginCommandId, XlsxPluginCommandBinding>();
  let pluginToolbarIds: readonly XlsxPluginCommandId[] = Object.freeze([]);
  let pluginShortcutList: readonly { id: XlsxPluginCommandId; chord: string }[] = Object.freeze([]);
  const scopedStores = new WeakMap<XlsxCommandScope, XlsxCommandStore>();
  const ownStores = new WeakSet<XlsxCommandStore>();
  const listeners = new Set<() => void>();
  const snapshots = new Map<string, CachedSnapshot>();
  const chromeRoots = new Set<HTMLElement>();
  const chromeListeners = new Set<() => void>();
  let chromeSnapshot: XlsxChromeContext | null = null;
  let round = 0;

  const environment = (executing: boolean): XlsxCommandEnvironment | null => {
    if (!binding) return null;
    try {
      return binding.environment(executing);
    } catch (error) {
      console.error('[xlsx commands] reading editor state failed', error);
      return null;
    }
  };

  const pluginState = (
    id: XlsxPluginCommandId,
    env: XlsxCommandEnvironment | null
  ): XlsxPluginCommandState => {
    const command = pluginCommands.get(id);
    if (!command) {
      return { enabled: false, disabledReason: commandReason('unsupported-command', env) };
    }
    const gate = contributedCommandGate(command.descriptor.mutatesDocument, env);
    return gate ? { enabled: false, disabledReason: gate } : command.state();
  };

  const compute = (id: AnyCommandId, args: unknown, env: XlsxCommandEnvironment | null) =>
    isPluginCommandId(id)
      ? pluginState(id, env)
      : (evaluateXlsxCommand(id, args as never, env) as XlsxCommandState);

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
          } as XlsxCommandState);
      entry = { id, args, json: JSON.stringify(state), state, holds: 0, read: round };
      snapshots.set(key, entry);
      trim();
    }
    entry.read = round;
    return entry;
  };

  const refreshChrome = () => {
    let next: XlsxChromeContext | null = null;
    try {
      next = binding?.chrome() ?? null;
    } catch (error) {
      console.error('[xlsx commands] reading editor chrome failed', error);
    }
    if (next?.i18n === chromeSnapshot?.i18n && (next === null) === (chromeSnapshot === null)) return;
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

  const capture = (current: XlsxCommandBinding, id: XlsxCommandId) => {
    try {
      return current.capture(id);
    } catch (error) {
      console.error('[xlsx commands] capturing a command target failed', error);
      return null;
    }
  };

  const perform = (id: AnyCommandId, args: unknown, env: XlsxCommandEnvironment) =>
    isPluginCommandId(id)
      ? pluginCommands.get(id)!.execute()
      : binding!.perform(id, args as never, env);

  const run = async (
    id: AnyCommandId,
    args: unknown,
    prepared?: XlsxCommandOrigin | null,
    scope?: XlsxCommandScope
  ): Promise<XlsxCommandResult> => {
    const current = binding;
    if (!current) return failure('editor-unavailable', null);
    const builtIn = isPluginCommandId(id) ? null : id;
    const ordered = builtIn !== null && current.ordered(builtIn);
    const origin =
      prepared !== undefined ? prepared : ordered && builtIn ? capture(current, builtIn) : null;
    const attempt = () => {
      if (binding !== current) return failure('editor-unavailable', null);
      if (origin && builtIn) {
        const stale = current.resume(origin, builtIn);
        if (stale) return failure(stale, environment(true));
      } else if (prepared !== undefined) {
        return failure('editor-unavailable', environment(true));
      }
      const denied = scope?.deny(id) ?? null;
      if (denied) return failure(denied, environment(true));
      const env = environment(true);
      const state = compute(id, args, env);
      if (!state.enabled) return { ok: false, failure: state.disabledReason } as XlsxCommandResult;
      if (scope && mutatingBuiltIn(id)) return failure('unsupported-policy', env);
      return perform(id, args, env!);
    };
    try {
      return await (ordered || prepared !== undefined ? current.admit(attempt) : attempt());
    } catch (error) {
      if (error instanceof XlsxCommandAdmissionError) {
        return failure(error.code, environment(false));
      }
      console.error(`[xlsx commands] ${id} failed`, error);
      return failure('execution-failed', environment(false));
    } finally {
      refresh();
    }
  };

  const dispatch = (
    id: unknown,
    args: unknown,
    scope?: XlsxCommandScope
  ): Promise<XlsxCommandResult> => {
    if (!knownCommand(id)) {
      return Promise.resolve(failure('unsupported-command', environment(false)));
    }
    return run(id, args, undefined, scope);
  };

  const getDescriptor = (id: AnyCommandId) =>
    isPluginCommandId(id)
      ? (pluginCommands.get(id)?.descriptor ?? null)
      : XLSX_COMMAND_DESCRIPTORS[id];

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
  }) as XlsxCommandStore;

  const prepareFor = (id: XlsxCommandId, scope?: XlsxCommandScope) => {
    const origin = binding ? capture(binding, id) : null;
    return { execute: (args: unknown) => run(id, args, origin, scope) };
  };

  const createScoped = (scope: XlsxCommandScope): XlsxCommandStore => {
    const cache = new Map<
      string,
      { base: unknown; denial: XlsxCommandFailureCode | null; policy: boolean; state: unknown }
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
    }) as XlsxCommandStore;
    controllers.set(scopedStore, {
      ...controller,
      store: scopedStore,
      prepare: (id) => prepareFor(id, scope),
    });
    return scopedStore;
  };

  const controller: XlsxCommandController = {
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
    prepare: (id) => prepareFor(id),
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
export function xlsxCommandController(store: XlsxCommandStore): XlsxCommandController | null {
  return controllers.get(store) ?? null;
}

/** A store without an editor: every command reports `editor-unavailable`. */
export const UNAVAILABLE_XLSX_COMMANDS: XlsxCommandStore = createXlsxCommandController().store;
