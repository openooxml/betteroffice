import type { Translations } from '@betteroffice/pptx-i18n';
import { isPptxCommandId, PPTX_COMMAND_DESCRIPTORS } from './descriptors';
import { commandReason, evaluatePptxCommand, type PptxCommandEnvironment } from './evaluate';
import type {
  PptxCommandArgs,
  PptxCommandDescriptor,
  PptxCommandFailureCode,
  PptxCommandId,
  PptxCommandResult,
  PptxCommandState,
  PptxCommandStore,
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

/** Private controls of a store: binding, refresh, deferred actions, chrome. */
export interface PptxCommandController {
  readonly store: PptxCommandStore;
  attach(binding: PptxCommandBinding): void;
  detach(binding: PptxCommandBinding): void;
  /** Re-evaluates cached snapshots and chrome, notifying the listeners of what changed. */
  refresh(): void;
  /** Keeps one snapshot cached, and its identity stable, while a subscriber shows it. */
  hold<K extends PptxCommandId>(id: K, args?: PptxCommandArgs[K]): () => void;
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
}

const MAX_CACHED_SNAPSHOTS = 512;

interface CachedSnapshot {
  id: PptxCommandId;
  args: unknown;
  json: string;
  state: PptxCommandState;
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

function snapshotKey(id: PptxCommandId, args: unknown): string {
  return args === undefined ? id : `${id}\u0000${JSON.stringify(args)}`;
}

/** Creates the command store of one editor mount. */
export function createPptxCommandController(): PptxCommandController {
  let binding: PptxCommandBinding | null = null;
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

  const compute = (id: PptxCommandId, args: unknown, env: PptxCommandEnvironment | null) =>
    evaluatePptxCommand(id, args as never, env) as PptxCommandState;

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

  const snapshot = (id: PptxCommandId, args: unknown): CachedSnapshot => {
    const key = snapshotKey(id, args);
    let entry = snapshots.get(key);
    if (!entry) {
      const state = isPptxCommandId(id)
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
      if (!isPptxCommandId(entry.id)) continue;
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

  const run = async <K extends PptxCommandId>(
    id: K,
    args: PptxCommandArgs[K],
    perform: (env: PptxCommandEnvironment) => PptxCommandResult | Promise<PptxCommandResult>,
    deferred?: { origin: PptxCommandOrigin | null }
  ): Promise<PptxCommandResult> => {
    const current = binding;
    if (!current) return failure('editor-unavailable', null);
    const ordered = deferred !== undefined || current.ordered(id, args);
    const origin = deferred ? deferred.origin : ordered ? capture(id) : undefined;
    const attempt = () => {
      if (origin !== undefined) {
        const stale = origin ? current.resume(origin) : 'editor-unavailable';
        if (stale) return failure(stale, environment(true));
      }
      const env = environment(true);
      const state = evaluatePptxCommand(id, args, env);
      if (!state.enabled) return { ok: false, failure: state.disabledReason } as PptxCommandResult;
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

  const store: PptxCommandStore = Object.freeze({
    getDescriptor<K extends PptxCommandId>(id: K): PptxCommandDescriptor<K> {
      return PPTX_COMMAND_DESCRIPTORS[id] as PptxCommandDescriptor<K>;
    },
    getState<K extends PptxCommandId>(id: K, args?: PptxCommandArgs[K]): PptxCommandState<K> {
      return snapshot(id, args).state as PptxCommandState<K>;
    },
    subscribe(listener: () => void): () => void {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    execute<K extends PptxCommandId>(id: K, args: PptxCommandArgs[K]): Promise<PptxCommandResult> {
      if (!isPptxCommandId(id)) {
        return Promise.resolve(failure('unsupported-command', environment(false)));
      }
      const current = binding;
      return run(id, args, (env) => current!.perform(id, args, env));
    },
  });

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
    defer(id, args) {
      const origin = capture(id);
      return { complete: (write) => run(id, args, () => write(), { origin }) };
    },
    prepare(id) {
      const origin = capture(id);
      return {
        execute: (args) => {
          const current = binding;
          return run(id, args, (env) => current!.perform(id, args, env), { origin });
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
export function pptxCommandController(store: PptxCommandStore): PptxCommandController | null {
  return controllers.get(store) ?? null;
}

/** A store without an editor: every command reports `editor-unavailable`. */
export const UNAVAILABLE_PPTX_COMMANDS: PptxCommandStore = createPptxCommandController().store;
