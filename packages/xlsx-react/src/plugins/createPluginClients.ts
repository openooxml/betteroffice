import type {
  Selection,
  WorkbookHandle,
  XlsxEditRefusal,
  XlsxEditRequest,
  XlsxEditResult,
} from '@betteroffice/xlsx';
import { grantsCommand, grantsEditBatch, grantsWrite } from '../../../../shared/plugin-host/grants';
import type { InvocationRefusal, PluginInvocation } from '../../../../shared/plugin-host/runtime';
import {
  UNAVAILABLE_XLSX_COMMANDS,
  type XlsxCommandController,
  type XlsxCommandScope,
} from '../commands/createXlsxCommandStore';
import { isPluginCommandId, XLSX_COMMAND_DESCRIPTORS } from '../commands/descriptors';
import type { XlsxCommandStore } from '../commands/types';
import type {
  XlsxPluginCommandClient,
  XlsxPluginEditClient,
  XlsxPluginFailureCode,
  XlsxPluginGrant,
  XlsxPluginNavigation,
  XlsxPluginNavigationFailureCode,
  XlsxPluginNavigationOptions,
  XlsxPluginNavigationResult,
  XlsxPluginReadClient,
  XlsxPluginRefusal,
  XlsxPluginSnapshot,
} from './types';

/** Work the editor ran after pending input, or why it could not. */
export type XlsxAdmission<T> =
  | { ok: true; value: T }
  | { ok: false; code: 'document-replaced' | 'input-failed'; error: unknown };

export type XlsxRevealAlignment = 'nearest' | 'start' | 'center';

/**
 * Selection and scrolling of the editor, run inside `admit`; each returns why it failed. Work
 * deferred to a later frame runs only while `live` still holds.
 */
export interface XlsxPluginNavigator {
  selectCells(
    handle: WorkbookHandle,
    target: { sheetId: string; selection: Selection },
    focus: boolean,
    live: () => boolean
  ): XlsxPluginNavigationFailureCode | null;
  scrollToCell(
    handle: WorkbookHandle,
    target: { sheetId: string; row: number; col: number },
    align: XlsxRevealAlignment,
    live: () => boolean
  ): XlsxPluginNavigationFailureCode | null;
}

/** The editor as plugin clients reach it; every member is read at call time. */
export interface XlsxPluginEditorAccess {
  /** The open workbook, or null. */
  handle(): WorkbookHandle | null;
  /** Runs `operation` after pending input while `handle` is still the open workbook. */
  admit<T>(handle: WorkbookHandle, operation: () => T): Promise<XlsxAdmission<T>>;
  /** The editor's refusal of writes while read-only, or null. */
  readOnlyRefusal(handle: WorkbookHandle): XlsxEditRefusal | null;
  /**
   * The editor's batch path, run inside `admit`: `authorize`, the read-only refusal, one
   * committed batch, then one `onChange` when it applied.
   */
  applyEdits<Refusal>(
    handle: WorkbookHandle,
    request: XlsxEditRequest,
    authorize: () => Refusal | null
  ): XlsxEditResult | Refusal;
  commands(): XlsxCommandController | null;
  navigator: XlsxPluginNavigator;
}

const MESSAGES: Record<XlsxPluginFailureCode, string> = {
  'plugin-unavailable': 'The plugin is no longer active',
  'document-replaced': 'The workbook was replaced',
  aborted: 'This plugin context was superseded',
  'permission-denied': 'The host has not granted this operation',
  'read-only': 'The editor is read-only',
  'input-failed': 'Pending input could not be applied',
  'unsupported-policy': 'Plugins cannot run this operation yet',
  'plugin-failed': 'The plugin failed',
};

const NAVIGATION_MESSAGES: Record<XlsxPluginNavigationFailureCode, string> = {
  'stale-version': 'The workbook changed after that version',
  'missing-target': 'The target is not in the workbook',
  'ambiguous-target': 'The target does not identify one place',
  'layout-unavailable': 'The sheet cannot be shown',
  unsupported: 'The editor cannot show this target now',
};

const ALIGNMENTS: ReadonlySet<string> = new Set(['nearest', 'start', 'center']);

export function pluginRefusal(
  code: XlsxPluginFailureCode,
  message = MESSAGES[code]
): XlsxPluginRefusal {
  return { ok: false, failure: { code, message } };
}

function refusalOf(invocation: PluginInvocation<XlsxPluginSnapshot>): XlsxPluginRefusal | null {
  const refusal: InvocationRefusal | null = invocation.refusal();
  return refusal ? pluginRefusal(refusal) : null;
}

/** The command scope of one invocation: its activation and the plugin's live grant. */
export function pluginCommandScope(
  invocation: PluginInvocation<XlsxPluginSnapshot>,
  grant: () => XlsxPluginGrant,
  subscribe: (listener: () => void) => () => void
): XlsxCommandScope {
  return {
    deny(id) {
      const refusal = invocation.refusal();
      if (refusal) return refusal;
      if (isPluginCommandId(id)) {
        return id.startsWith(`plugin:${invocation.pluginId}/`) ? null : 'permission-denied';
      }
      const current = grant();
      if (!grantsCommand(current, id)) return 'permission-denied';
      if (XLSX_COMMAND_DESCRIPTORS[id].mutatesDocument && !grantsWrite(current)) {
        return 'permission-denied';
      }
      return null;
    },
    subscribe,
  };
}

export interface XlsxPluginClients {
  read: XlsxPluginReadClient;
  commands: XlsxPluginCommandClient;
  edits: XlsxPluginEditClient | null;
  navigation: XlsxPluginNavigation;
}

type Admitted<T> = { refused: XlsxPluginRefusal } | { value: T };

/**
 * `work`'s outcome, or null when a signal aborts before `started` turns true. Once the work has
 * started its outcome is authoritative, so an abort it causes itself cannot hide it.
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
  invocation: PluginInvocation<XlsxPluginSnapshot>,
  access: XlsxPluginEditorAccess,
  grant: () => XlsxPluginGrant,
  store: XlsxCommandStore
): XlsxPluginClients {
  /**
   * Runs `use` after pending input, checking this invocation in the same step. A superseded or
   * ended invocation refuses at once rather than waiting for the input ahead of it; once `use`
   * has started, its result stands.
   */
  const whenAdmitted = async <T>(
    use: (handle: WorkbookHandle) => T
  ): Promise<T | XlsxPluginRefusal> => {
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

  const read: XlsxPluginReadClient = {
    version: () => whenAdmitted((handle) => ({ ok: true as const, version: handle.version() })),
    readCells: (request) => whenAdmitted((handle) => handle.readCells(request)),
    findText: (request) => whenAdmitted((handle) => handle.findText(request)),
    validateEdits: (request) =>
      whenAdmitted((handle) => access.readOnlyRefusal(handle) ?? handle.validateEdits(request)),
  };

  const batchDenial = (history: 'separate' | 'none' | undefined): XlsxPluginRefusal | null =>
    refusalOf(invocation) ??
    (grantsEditBatch(grant(), history) ? null : pluginRefusal('permission-denied'));

  const edits: XlsxPluginEditClient | null = grantsEditBatch(grant(), 'separate')
    ? {
        async applyEdits(request) {
          const denied = batchDenial(request?.history);
          if (denied) return denied;
          return whenAdmitted((handle) =>
            access.applyEdits(handle, request, () => batchDenial(request.history))
          );
        },
      }
    : null;

  const commands = {
    getDescriptor: (id: never) => store.getDescriptor(id),
    getState: (id: never, args?: never) => store.getState(id, args),
    subscribe: (listener: () => void) => store.subscribe(listener),
    execute: async (id: never, args: never) => refusalOf(invocation) ?? store.execute(id, args),
  } as unknown as XlsxPluginCommandClient;

  const live = () => invocation.refusal() === null;

  const navigate = async (
    options: XlsxPluginNavigationOptions | undefined,
    go: (handle: WorkbookHandle) => XlsxPluginNavigationFailureCode | null
  ): Promise<XlsxPluginNavigationResult> => {
    const outcome = await whenAdmitted((handle) =>
      handle.version() === options?.expectVersion ? go(handle) : 'stale-version'
    );
    if (outcome === null) return { ok: true };
    if (typeof outcome === 'object') return outcome;
    return { ok: false, failure: { code: outcome, message: NAVIGATION_MESSAGES[outcome] } };
  };

  const navigation: XlsxPluginNavigation = {
    selectCells: (target, options) =>
      navigate(options, (handle) =>
        access.navigator.selectCells(handle, target, options.focus === true, live)
      ),
    scrollToCell: (target, options) =>
      navigate(options, (handle) => {
        const align = options.align ?? 'nearest';
        if (!ALIGNMENTS.has(align)) return 'unsupported';
        return access.navigator.scrollToCell(handle, target, align, live);
      }),
  };

  return { read, commands, edits, navigation };
}

/** The store plugin clients use: the editor's commands scoped to the plugin, or none. */
export function scopedCommandStore(
  controller: XlsxCommandController | null,
  scope: XlsxCommandScope
): XlsxCommandStore {
  return controller?.scoped(scope) ?? UNAVAILABLE_XLSX_COMMANDS;
}
