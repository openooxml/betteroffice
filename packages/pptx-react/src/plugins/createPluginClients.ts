import type {
  PptxEditRefusal,
  PptxEditRequest,
  PptxEditResult,
  PresentationHandle,
} from '@betteroffice/pptx';
import { grantsCommand, grantsEditBatch, grantsWrite } from '../../../../shared/plugin-host/grants';
import type { InvocationRefusal, PluginInvocation } from '../../../../shared/plugin-host/runtime';
import {
  UNAVAILABLE_PPTX_COMMANDS,
  type PptxCommandController,
  type PptxCommandScope,
} from '../commands/createPptxCommandStore';
import { isPluginCommandId, PPTX_COMMAND_DESCRIPTORS } from '../commands/descriptors';
import type { PptxCommandStore } from '../commands/types';
import type {
  PptxPluginCommandClient,
  PptxPluginEditClient,
  PptxPluginFailureCode,
  PptxPluginGrant,
  PptxPluginNavigation,
  PptxPluginNavigationFailureCode,
  PptxPluginNavigationOptions,
  PptxPluginNavigationResult,
  PptxPluginReadClient,
  PptxPluginRefusal,
  PptxPluginSnapshot,
} from './types';

/** Work the editor ran after pending input, or why it could not. */
export type PptxAdmission<T> =
  | { ok: true; value: T }
  | { ok: false; code: 'document-replaced' | 'input-failed'; error: unknown };

/** Selection and navigation of the editor, run inside `admit`; each returns why it failed. */
export interface PptxPluginNavigator {
  goToSlide(
    handle: PresentationHandle,
    target: { slideId: string },
    focus: boolean
  ): PptxPluginNavigationFailureCode | null;
  selectShape(
    handle: PresentationHandle,
    target: { slideId: string; shapeId: string },
    focus: boolean
  ): PptxPluginNavigationFailureCode | null;
  selectText(
    handle: PresentationHandle,
    target: { slideId: string; shapeId: string; storyId: string; anchor: number; focus: number },
    focus: boolean
  ): PptxPluginNavigationFailureCode | null;
}

/** The editor as plugin clients reach it; every member is read at call time. */
export interface PptxPluginEditorAccess {
  /** The open presentation, or null. */
  handle(): PresentationHandle | null;
  /** Runs `operation` after pending input while `handle` is still the open presentation. */
  admit<T>(handle: PresentationHandle, operation: () => T): Promise<PptxAdmission<T>>;
  /** The editor's refusal of writes while read-only, or null. */
  readOnlyRefusal(handle: PresentationHandle): PptxEditRefusal | null;
  /**
   * The editor's batch path, run inside `admit`: `authorize`, the read-only refusal, one
   * transaction inside `commit`, then one refresh when it applied.
   */
  applyEdits<Refusal>(
    handle: PresentationHandle,
    request: PptxEditRequest,
    authorize: () => Refusal | null,
    commit: <T>(write: () => T) => T
  ): PptxEditResult | Refusal;
  commands(): PptxCommandController | null;
  navigator: PptxPluginNavigator;
}

const MESSAGES: Record<PptxPluginFailureCode, string> = {
  'plugin-unavailable': 'The plugin is no longer active',
  'document-replaced': 'The presentation was replaced',
  aborted: 'This plugin context was superseded',
  'permission-denied': 'The host has not granted this operation',
  'read-only': 'The editor is read-only',
  'input-failed': 'Pending input could not be applied',
  'unsupported-policy': 'Plugins cannot run this operation yet',
  'plugin-failed': 'The plugin failed',
};

const NAVIGATION_MESSAGES: Record<PptxPluginNavigationFailureCode, string> = {
  'stale-version': 'The presentation changed after that version',
  'missing-target': 'The target is not in the presentation',
  'ambiguous-target': 'The target does not identify one object',
  'layout-unavailable': 'The slide cannot be shown',
  unsupported: 'The editor cannot show this target now',
};

export function pluginRefusal(
  code: PptxPluginFailureCode,
  message = MESSAGES[code]
): PptxPluginRefusal {
  return { ok: false, failure: { code, message } };
}

function refusalOf(invocation: PluginInvocation<PptxPluginSnapshot>): PptxPluginRefusal | null {
  const refusal: InvocationRefusal | null = invocation.refusal();
  return refusal ? pluginRefusal(refusal) : null;
}

/** The command scope of one invocation: its activation and the plugin's live grant. */
export function pluginCommandScope(
  invocation: PluginInvocation<PptxPluginSnapshot>,
  grant: () => PptxPluginGrant,
  subscribe: (listener: () => void) => () => void
): PptxCommandScope {
  return {
    deny(id) {
      const refusal = invocation.refusal();
      if (refusal) return refusal;
      if (isPluginCommandId(id)) {
        return id.startsWith(`plugin:${invocation.pluginId}/`) ? null : 'permission-denied';
      }
      const current = grant();
      if (!grantsCommand(current, id)) return 'permission-denied';
      if (PPTX_COMMAND_DESCRIPTORS[id].mutatesDocument && !grantsWrite(current)) {
        return 'permission-denied';
      }
      return null;
    },
    subscribe,
  };
}

export interface PptxPluginClients {
  read: PptxPluginReadClient;
  commands: PptxPluginCommandClient;
  edits: PptxPluginEditClient | null;
  navigation: PptxPluginNavigation;
}

type Admitted<T> = { refused: PptxPluginRefusal } | { value: T };

/**
 * `work`'s outcome, or null as soon as a signal aborts before `started()`; once the work has
 * started its own outcome is returned, so a change it committed is never reported as aborted.
 */
function unlessAborted<T>(
  signals: readonly AbortSignal[],
  work: Promise<T>,
  started: () => boolean
): Promise<T | null> {
  if (!started() && signals.some((signal) => signal.aborted)) return Promise.resolve(null);
  return new Promise<T | null>((resolve, reject) => {
    const abort = () => {
      if (started()) return;
      detach();
      resolve(null);
    };
    const detach = () => {
      for (const signal of signals) signal.removeEventListener('abort', abort);
    };
    for (const signal of signals) signal.addEventListener('abort', abort, { once: true });
    work.then(
      (value) => {
        detach();
        resolve(value);
      },
      (error: unknown) => {
        detach();
        reject(error);
      }
    );
  });
}

/** Clients bound to one invocation: they refuse once it is superseded or its plugin ends. */
export function createPluginClients(
  invocation: PluginInvocation<PptxPluginSnapshot>,
  access: PptxPluginEditorAccess,
  grant: () => PptxPluginGrant,
  store: PptxCommandStore
): PptxPluginClients {
  /**
   * Runs `use` after pending input, checking this invocation in the same step. A superseded or
   * ended invocation refuses at once rather than waiting for the input ahead of it.
   */
  const whenAdmitted = async <T>(
    use: (handle: PresentationHandle) => T
  ): Promise<T | PptxPluginRefusal> => {
    const before = refusalOf(invocation);
    if (before) return before;
    const handle = access.handle();
    if (!handle) return pluginRefusal('plugin-unavailable');
    let started = false;
    const admitted = await unlessAborted(
      [invocation.signal, invocation.lifetimeSignal],
      access.admit(handle, (): Admitted<T> => {
        const refused = refusalOf(invocation);
        if (refused) return { refused };
        started = true;
        return { value: use(handle) };
      }),
      () => started
    );
    if (!admitted) return refusalOf(invocation) ?? pluginRefusal('aborted');
    if (!admitted.ok) return refusalOf(invocation) ?? pluginRefusal(admitted.code);
    return 'refused' in admitted.value ? admitted.value.refused : admitted.value.value;
  };

  const read: PptxPluginReadClient = {
    version: () => whenAdmitted((handle) => ({ ok: true as const, version: handle.version() })),
    readContent: (request) => whenAdmitted((handle) => handle.readContent(request)),
    findText: (request) => whenAdmitted((handle) => handle.findText(request)),
    validateEdits: (request) =>
      whenAdmitted((handle) => access.readOnlyRefusal(handle) ?? handle.validateEdits(request)),
  };

  const batchDenial = (history: 'separate' | 'none' | undefined): PptxPluginRefusal | null =>
    refusalOf(invocation) ??
    (grantsEditBatch(grant(), history) ? null : pluginRefusal('permission-denied'));

  const edits: PptxPluginEditClient | null = grantsEditBatch(grant(), 'separate')
    ? {
        async applyEdits(request) {
          const denied = batchDenial(request.history);
          if (denied) return denied;
          return whenAdmitted((handle) =>
            access.applyEdits(
              handle,
              request,
              () => batchDenial(request.history),
              (write) => invocation.commit(write)
            )
          );
        },
      }
    : null;

  const commands = {
    getDescriptor: (id: never) => store.getDescriptor(id),
    getState: (id: never, args?: never) => store.getState(id, args),
    subscribe: (listener: () => void) => store.subscribe(listener),
    execute: async (id: never, args: never) => refusalOf(invocation) ?? store.execute(id, args),
  } as unknown as PptxPluginCommandClient;

  const navigate = async (
    options: PptxPluginNavigationOptions,
    go: (handle: PresentationHandle, focus: boolean) => PptxPluginNavigationFailureCode | null
  ): Promise<PptxPluginNavigationResult> => {
    const outcome = await whenAdmitted((handle) =>
      handle.version() === options?.expectVersion
        ? go(handle, options.focus === true)
        : 'stale-version'
    );
    if (outcome === null) return { ok: true };
    if (typeof outcome === 'object') return outcome;
    return { ok: false, failure: { code: outcome, message: NAVIGATION_MESSAGES[outcome] } };
  };

  const navigation: PptxPluginNavigation = {
    goToSlide: (target, options) =>
      navigate(options, (handle, focus) => access.navigator.goToSlide(handle, target, focus)),
    selectShape: (target, options) =>
      navigate(options, (handle, focus) => access.navigator.selectShape(handle, target, focus)),
    selectText: (target, options) =>
      navigate(options, (handle, focus) => access.navigator.selectText(handle, target, focus)),
  };

  return { read, commands, edits, navigation };
}

/** The store plugin clients use: the editor's commands scoped to the plugin, or none. */
export function scopedCommandStore(
  controller: PptxCommandController | null,
  scope: PptxCommandScope
): PptxCommandStore {
  return controller?.scoped(scope) ?? UNAVAILABLE_PPTX_COMMANDS;
}
