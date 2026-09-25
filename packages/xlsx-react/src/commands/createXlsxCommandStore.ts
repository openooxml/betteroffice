import type { Translations } from '@betteroffice/xlsx-i18n';
import { isXlsxCommandId, XLSX_COMMAND_DESCRIPTORS } from './descriptors';
import {
  commandReason,
  evaluateXlsxCommand,
  type XlsxCommandEnvironment,
} from './evaluate';
import type {
  XlsxCommandArgs,
  XlsxCommandDescriptor,
  XlsxCommandFailureCode,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandState,
  XlsxCommandStore,
} from './types';

/** Raised when an admitted command cannot run after the input before it. */
export class XlsxCommandAdmissionError extends Error {
  constructor(
    readonly code: Extract<
      XlsxCommandFailureCode,
      'input-failed' | 'document-replaced' | 'target-changed' | 'gesture-active' | 'editor-unavailable'
    >
  ) {
    super(code);
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

/** Private controls of a store: binding, refresh, deferred choices, chrome. */
export interface XlsxCommandController {
  readonly store: XlsxCommandStore;
  attach(binding: XlsxCommandBinding): void;
  detach(binding: XlsxCommandBinding): void;
  /** Re-evaluates cached snapshots and chrome, notifying the listeners of what changed. */
  refresh(): void;
  /** Keeps one snapshot cached, and its identity stable, while a subscriber shows it. */
  hold<K extends XlsxCommandId>(id: K, args?: XlsxCommandArgs[K]): () => void;
  /** Captures the document and target now for `id`, run later with arguments chosen then. */
  prepare<K extends XlsxCommandId>(id: K): XlsxPendingCommand<K>;
  /** Marks host chrome whose keyboard shortcuts belong to this editor. */
  registerChrome(element: HTMLElement): () => void;
  ownsChrome(target: EventTarget | null): boolean;
  /** Stable until the editor's locale changes. */
  chrome(): XlsxChromeContext | null;
  subscribeChrome(listener: () => void): () => void;
  focusEditor(): void;
}

const MAX_CACHED_SNAPSHOTS = 512;

interface CachedSnapshot {
  id: XlsxCommandId;
  args: unknown;
  json: string;
  state: XlsxCommandState;
  holds: number;
  /** Notification round in which it was last read. */
  read: number;
}

const controllers = new WeakMap<XlsxCommandStore, XlsxCommandController>();

function failure(code: XlsxCommandFailureCode, env: XlsxCommandEnvironment | null): XlsxCommandResult {
  return { ok: false, failure: commandReason(code, env) };
}

function snapshotKey(id: XlsxCommandId, args: unknown): string {
  return args === undefined ? id : `${id}\u0000${JSON.stringify(args)}`;
}

/** Creates the command store of one editor mount. */
export function createXlsxCommandController(): XlsxCommandController {
  let binding: XlsxCommandBinding | null = null;
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

  const compute = (id: XlsxCommandId, args: unknown, env: XlsxCommandEnvironment | null) =>
    evaluateXlsxCommand(id, args as never, env) as XlsxCommandState;

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

  const snapshot = (id: XlsxCommandId, args: unknown): CachedSnapshot => {
    const key = snapshotKey(id, args);
    let entry = snapshots.get(key);
    if (!entry) {
      const state = isXlsxCommandId(id)
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

  const run = async <K extends XlsxCommandId>(
    id: K,
    args: XlsxCommandArgs[K],
    prepared?: XlsxCommandOrigin | null
  ): Promise<XlsxCommandResult> => {
    const current = binding;
    if (!current) return failure('editor-unavailable', null);
    const ordered = current.ordered(id);
    const origin = prepared !== undefined ? prepared : ordered ? capture(current, id) : null;
    const attempt = () => {
      if (binding !== current) return failure('editor-unavailable', null);
      if (origin) {
        const stale = current.resume(origin, id);
        if (stale) return failure(stale, environment(true));
      } else if (prepared !== undefined) {
        return failure('editor-unavailable', environment(true));
      }
      const env = environment(true);
      const state = evaluateXlsxCommand(id, args, env);
      if (!state.enabled) return { ok: false, failure: state.disabledReason } as XlsxCommandResult;
      return current.perform(id, args, env!);
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

  const store: XlsxCommandStore = Object.freeze({
    getDescriptor<K extends XlsxCommandId>(id: K): XlsxCommandDescriptor<K> {
      return XLSX_COMMAND_DESCRIPTORS[id] as XlsxCommandDescriptor<K>;
    },
    getState<K extends XlsxCommandId>(id: K, args?: XlsxCommandArgs[K]): XlsxCommandState<K> {
      return snapshot(id, args).state as XlsxCommandState<K>;
    },
    subscribe(listener: () => void): () => void {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    execute<K extends XlsxCommandId>(id: K, args: XlsxCommandArgs[K]): Promise<XlsxCommandResult> {
      if (!isXlsxCommandId(id)) {
        return Promise.resolve(failure('unsupported-command', environment(false)));
      }
      return run(id, args);
    },
  });

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
    hold(id, args) {
      const entry = snapshot(id, args);
      entry.holds += 1;
      let held = true;
      return () => {
        if (!held) return;
        held = false;
        entry.holds -= 1;
      };
    },
    prepare(id) {
      const origin = binding ? capture(binding, id) : null;
      return { execute: (args) => run(id, args, origin) };
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
  };
  controllers.set(store, controller);
  return controller;
}

/** Private controls of a store created by this package, if any. */
export function xlsxCommandController(store: XlsxCommandStore): XlsxCommandController | null {
  return controllers.get(store) ?? null;
}

/** A store without an editor: every command reports `editor-unavailable`. */
export const UNAVAILABLE_XLSX_COMMANDS: XlsxCommandStore = createXlsxCommandController().store;
