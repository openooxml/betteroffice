import { grantsEditBatch } from '../../../../shared/plugin-host/grants';
import type { PluginInvocation } from '../../../../shared/plugin-host/runtime';
import type { XlsxCommandStore } from '../commands/types';
import {
  ALIGNMENTS, NAVIGATION_MESSAGES, pluginRefusal, unlessAborted, type XlsxPluginClients, type XlsxPluginEditorAccess,
} from './createPluginClients';
import type { XlsxPluginDocumentAccess } from './createPluginDocumentAccess';
import type {
  XlsxPluginCommandClient, XlsxPluginEditClient, XlsxPluginGrant, XlsxPluginNavigation,
  XlsxPluginNavigationFailureCode, XlsxPluginNavigationOptions, XlsxPluginNavigationResult,
  XlsxPluginReadClient, XlsxPluginRefusal, XlsxPluginSnapshot,
} from './types';

export function createPeerlessPluginClients(
  invocation: PluginInvocation<XlsxPluginSnapshot>,
  access: XlsxPluginDocumentAccess | XlsxPluginEditorAccess,
  grant: () => XlsxPluginGrant,
  store: XlsxCommandStore
): XlsxPluginClients {
  if (!('identity' in access)) throw new Error('Async plugin clients require document access');
  const refusal = () => {
    const code = invocation.refusal();
    return code ? pluginRefusal(code) : null;
  };
  const whenAdmitted = async <T>(use: (identity: object) => T, fence = true): Promise<Awaited<T> | XlsxPluginRefusal> => {
    const before = refusal();
    if (before) return before;
    const identity = access.identity();
    if (!identity) return pluginRefusal('plugin-unavailable');
    let started = false;
    const admitted = await unlessAborted([invocation.signal, invocation.lifetimeSignal],
      access.admit(identity, (): { refused: XlsxPluginRefusal } | { value: T } => {
        const denied = refusal();
        if (denied) return { refused: denied };
        started = true;
        return { value: use(identity) };
      }, fence), () => started);
    if (!admitted) return refusal() ?? pluginRefusal('aborted');
    if (!admitted.ok) return refusal() ?? pluginRefusal(admitted.code);
    if ('refused' in admitted.value) return admitted.value.refused;
    try {
      const value = await admitted.value.value;
      return access.identity() === identity ? value : pluginRefusal('document-replaced');
    } catch { return refusal() ?? pluginRefusal('input-failed'); }
  };
  const read: XlsxPluginReadClient = {
    version: () => whenAdmitted((identity) => access.readVersion(identity)),
    readCells(request) {
      const input = structuredClone(request);
      return whenAdmitted((identity) => access.readCells(identity, input));
    },
    findText(request) {
      const input = structuredClone(request);
      return whenAdmitted((identity) => access.findText(identity, input));
    },
    validateEdits(request) {
      const input = structuredClone(request);
      return whenAdmitted((identity) => access.validateEdits(identity, input));
    },
  };
  const batchDenial = (history: 'separate' | 'none' | undefined) =>
    refusal() ?? (grantsEditBatch(grant(), history) ? null : pluginRefusal('permission-denied'));
  const commit = <T>(write: () => T) => invocation.commit(write);
  const edits: XlsxPluginEditClient | null = grantsEditBatch(grant(), 'separate') ? {
    async applyEdits(request) {
      const denied = batchDenial(request?.history);
      if (denied) return denied;
      if (!access.admitEdits) return whenAdmitted((identity) => access.applyEdits(identity, request,
        () => batchDenial(request.history), commit));
      const identity = access.identity();
      if (!identity) return pluginRefusal('plugin-unavailable');
      const input = structuredClone(request);
      const admitted = await access.admitEdits(identity, input,
        (recovering) => recovering
          ? grantsEditBatch(grant(), input.history) ? null : pluginRefusal('permission-denied')
          : batchDenial(input.history),
        commit, [invocation.signal, invocation.lifetimeSignal]);
      return admitted.ok ? admitted.value : pluginRefusal(admitted.code);
    },
  } : null;
  const commands = {
    getDescriptor: (id: never) => store.getDescriptor(id),
    getState: (id: never, args?: never) => store.getState(id, args),
    subscribe: (listener: () => void) => store.subscribe(listener),
    execute: async (id: never, args: never) => refusal() ?? store.execute(id, args),
  } as unknown as XlsxPluginCommandClient;
  const live = () => invocation.refusal() === null;
  const navigate = async (options: XlsxPluginNavigationOptions,
    go: (identity: object) => XlsxPluginNavigationFailureCode | null | Promise<XlsxPluginNavigationFailureCode | null>
  ): Promise<XlsxPluginNavigationResult> => {
    const outcome = await whenAdmitted((identity) =>
      access.version(identity) === options?.expectVersion ? go(identity) : 'stale-version');
    if (outcome === null) return { ok: true };
    if (typeof outcome === 'object') return outcome;
    return { ok: false, failure: { code: outcome, message: NAVIGATION_MESSAGES[outcome] } };
  };
  const navigation: XlsxPluginNavigation = {
    selectCells(target, options) {
      const input = structuredClone(target);
      const at = structuredClone(options);
      return navigate(at, (identity) => access.navigator.selectCells(identity, input, at.focus === true, live, at.expectVersion));
    },
    scrollToCell(target, options) {
      const input = structuredClone(target);
      const at = structuredClone(options);
      return navigate(at, (identity) => {
        const align = at.align ?? 'nearest';
        if (!ALIGNMENTS.has(align)) return 'unsupported';
        return access.navigator.scrollToCell(identity, input, align, live, at.expectVersion);
      });
    },
  };
  return { read, commands, edits, navigation };
}
