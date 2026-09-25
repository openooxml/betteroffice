import type { Translations } from '@betteroffice/docx-i18n';
import type { Theme } from '@betteroffice/docx/types/document';
import { DOCX_COMMAND_DESCRIPTORS, isDocxCommandId } from './descriptors';
import {
  commandReason,
  evaluateDocxCommand,
  type DocxCommandEnvironment,
} from './evaluate';
import type {
  DocxCommandArgs,
  DocxCommandDescriptor,
  DocxCommandFailureCode,
  DocxCommandId,
  DocxCommandResult,
  DocxCommandState,
  DocxCommandStore,
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

/** Private controls of a store: binding, refresh, deferred actions, chrome. */
export interface DocxCommandController {
  readonly store: DocxCommandStore;
  attach(binding: DocxCommandBinding): void;
  detach(binding: DocxCommandBinding): void;
  /** Re-evaluates cached snapshots and chrome, notifying the listeners of what changed. */
  refresh(): void;
  /** Keeps one snapshot cached, and its identity stable, while a subscriber shows it. */
  hold<K extends DocxCommandId>(id: K, args?: DocxCommandArgs[K]): () => void;
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
}

const MAX_CACHED_SNAPSHOTS = 512;

interface CachedSnapshot {
  id: DocxCommandId;
  args: unknown;
  json: string;
  state: DocxCommandState;
  /** Subscriptions that hold this snapshot. */
  holds: number;
  /** Notification round in which it was last read. */
  read: number;
}

const controllers = new WeakMap<DocxCommandStore, DocxCommandController>();

function failure(code: DocxCommandFailureCode, env: DocxCommandEnvironment | null): DocxCommandResult {
  return { ok: false, failure: commandReason(code, env) };
}

function snapshotKey(id: DocxCommandId, args: unknown): string {
  return args === undefined ? id : `${id}\u0000${JSON.stringify(args)}`;
}

function sameChrome(a: DocxChromeContext | null, b: DocxChromeContext | null): boolean {
  if (!a || !b) return a === b;
  return a.i18n === b.i18n && a.isDark === b.isDark && a.theme === b.theme;
}

/** Creates the command store of one editor mount. */
export function createDocxCommandController(): DocxCommandController {
  let binding: DocxCommandBinding | null = null;
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

  const compute = (id: DocxCommandId, args: unknown, env: DocxCommandEnvironment | null) =>
    evaluateDocxCommand(id, args as never, env) as DocxCommandState;

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

  const snapshot = (id: DocxCommandId, args: unknown): CachedSnapshot => {
    const key = snapshotKey(id, args);
    let entry = snapshots.get(key);
    if (!entry) {
      const state = isDocxCommandId(id)
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

  const run = async <K extends DocxCommandId>(
    id: K,
    args: DocxCommandArgs[K],
    perform: (env: DocxCommandEnvironment) => DocxCommandResult | Promise<DocxCommandResult>,
    origin?: DocxCommandOrigin | null
  ): Promise<DocxCommandResult> => {
    const current = binding;
    if (!current) return failure('editor-unavailable', null);
    const deferred = origin !== undefined;
    const attempt = () => {
      if (deferred) {
        const stale = origin ? current.resume(origin) : 'editor-unavailable';
        if (stale) return failure(stale, environment(true));
      }
      const env = environment(true);
      const state = evaluateDocxCommand(id, args, env);
      if (!state.enabled) return { ok: false, failure: state.disabledReason } as DocxCommandResult;
      return perform(env!);
    };
    try {
      return await (deferred || current.ordered(id, args) ? current.admit(attempt) : attempt());
    } catch (error) {
      if (error instanceof DocxCommandAdmissionError) return failure(error.code, environment(false));
      console.error(`[docx commands] ${id} failed`, error);
      return failure('command-failed', environment(false));
    } finally {
      refresh();
    }
  };

  const store: DocxCommandStore = Object.freeze({
    getDescriptor<K extends DocxCommandId>(id: K): DocxCommandDescriptor<K> {
      return DOCX_COMMAND_DESCRIPTORS[id] as DocxCommandDescriptor<K>;
    },
    getState<K extends DocxCommandId>(id: K, args?: DocxCommandArgs[K]): DocxCommandState<K> {
      return snapshot(id, args).state as DocxCommandState<K>;
    },
    subscribe(listener: () => void): () => void {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    execute<K extends DocxCommandId>(id: K, args: DocxCommandArgs[K]): Promise<DocxCommandResult> {
      if (!isDocxCommandId(id)) {
        return Promise.resolve(failure('unsupported-command', environment(false)));
      }
      const current = binding;
      return run(id, args, (env) => current!.perform(id, args, env));
    },
  });

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
