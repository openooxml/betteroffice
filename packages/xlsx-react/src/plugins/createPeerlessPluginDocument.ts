import type { XlsxEditFailureCode, XlsxEditRequest, XlsxEditResult } from '@betteroffice/xlsx';
import { XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import { PeerlessWorkerCommandRefusal, type PeerlessWorkerInputCoordinator } from '../commands/peerlessWorkerInputCoordinator';
import type { PeerlessWorkbookSession } from '../worker/usePeerlessSessionWorkbook';
import { pluginRefusal, unlessAborted, type XlsxAdmission } from './createPluginClients';
import type { XlsxPluginDocumentAccess } from './createPluginDocumentAccess';
import { createPluginGeometry, type PaintedGrid } from './geometry';
import type { XlsxPluginGeometry, XlsxPluginLayout, XlsxPluginRefusal, XlsxPluginSelection } from './types';

interface PeerlessPluginOptions {
  owner(): PeerlessWorkbookSession | null;
  coordinator(): PeerlessWorkerInputCoordinator | null;
  readOnly(): boolean;
  commands: XlsxPluginDocumentAccess['commands'];
  selection(owner: PeerlessWorkbookSession): XlsxPluginSelection;
  navigator: XlsxPluginDocumentAccess['navigator'];
}

const EDIT_FAILURES: ReadonlySet<string> = new Set<XlsxEditFailureCode>([
  'stale-version', 'missing-target', 'content-mismatch', 'overlapping-steps', 'locked-target',
  'read-only', 'unsupported', 'invalid-step', 'limit-exceeded',
]);

export function createPeerlessPluginDocument(options: PeerlessPluginOptions): XlsxPluginDocumentAccess {
  const owner = (identity: object): PeerlessWorkbookSession => {
    const current = options.owner();
    if (current !== identity || !current?.current || current.retiring) {
      throw new XlsxCommandAdmissionError('document-replaced');
    }
    if (current.failure) throw current.failure;
    return current;
  };
  const input = () => {
    const coordinator = options.coordinator();
    if (!coordinator) throw new XlsxCommandAdmissionError('editor-unavailable');
    return coordinator;
  };
  const admissionFailure = <T>(error: unknown): XlsxAdmission<T> => ({ ok: false,
    code: error instanceof XlsxCommandAdmissionError && error.code === 'document-replaced'
      ? 'document-replaced' : 'input-failed', error });
  const refused = (error: unknown): error is XlsxPluginRefusal =>
    typeof error === 'object' && error !== null && 'ok' in error && error.ok === false && 'failure' in error;
  const navigations = new WeakMap<PeerlessWorkbookSession, Promise<void>>();
  const navigate = (identity: object, use: () => ReturnType<XlsxPluginDocumentAccess['navigator']['selectCells']>) => {
    const opened = owner(identity);
    const work = (navigations.get(opened) ?? Promise.resolve()).then(() => {
      owner(identity);
      return use();
    });
    navigations.set(opened, work.then(() => {}, () => {}));
    return work;
  };
  const document: XlsxPluginDocumentAccess = {
    identity: () => {
      const current = options.owner();
      return current?.ready && current.current && !current.retiring && !current.failure ? current : null;
    },
    version: (identity) => document.identity() === identity ? owner(identity).session.confirmedBasis.version : null,
    currentVersion: (identity) => options.coordinator()?.pending ? null : document.version(identity),
    subscribe(identity, listener) {
      const opened = owner(identity);
      const offOwner = opened.subscribe(() => {
        if (document.identity() !== identity) listener();
      });
      const offInput = input().subscribe((event) => {
        if (event.type === 'coverage' && document.identity() === identity) listener();
      });
      return () => { offInput(); offOwner(); };
    },
    selection: (identity) => options.selection(owner(identity)),
    readOnly: (identity) => owner(identity).session.editState.readOnly === true,
    async admit<T>(identity: object, operation: () => T, fence = false) {
      try {
        owner(identity);
        return await input().runAfterPendingInput(() => {
          owner(identity);
          const value = operation();
          owner(identity);
          return { ok: true as const, value };
        }, fence);
      } catch (error) { return admissionFailure<T>(error); }
    },
    readVersion: async (identity) => ({ ok: true as const, version: await owner(identity).session.call.version() }),
    readCells: (identity, request) => owner(identity).session.call.readCells(request),
    findText: (identity, request) => owner(identity).session.call.findText(request),
    validateEdits: (identity, request) => document.readOnlyRefusal(identity) ?? owner(identity).session.call.validateEdits(request),
    readOnlyRefusal: (identity) => options.readOnly() || owner(identity).session.editState.readOnly ? {
      ok: false, version: owner(identity).session.confirmedBasis.version,
      failure: { code: 'read-only', message: 'The editor is read-only' },
    } : null,
    applyEdits: () => { throw pluginRefusal('unsupported-policy'); },
    async admitEdits<Refusal>(identity: object, request: XlsxEditRequest,
      authorize: (recovering: boolean) => Refusal | null, commit: <T>(write: () => T) => T,
      signals: readonly AbortSignal[] = []) {
      let started = false;
      let opened: PeerlessWorkbookSession | null = null;
      try {
        opened = owner(identity);
        const coordinator = input();
        const work = coordinator.runAfterPendingInput<XlsxEditResult | Refusal>(() => {
          owner(identity);
          const denial = authorize(false);
          if (denial) return denial;
          const readOnly = document.readOnlyRefusal(identity);
          if (readOnly) return readOnly;
          return coordinator.command({ kind: 'applyEdits', payload: { args: [request] } }, {
            acknowledge: commit,
            dispatch(write) {
              owner(identity);
              const denied = authorize(false) ?? document.readOnlyRefusal(identity);
              if (denied) throw denied;
              started = true;
              return commit(write);
            },
          }).result.then((status) => {
            if (!status.result) throw pluginRefusal('input-failed');
            return status.result as XlsxEditResult;
          });
        });
        const value = await unlessAborted(signals, work, () => started);
        if (value === null) {
          const denial = authorize(false);
          if (denial) return { ok: true as const, value: denial };
          throw pluginRefusal('aborted');
        }
        if (!started) owner(identity);
        return { ok: true as const, value };
      } catch (error) {
        if (refused(error)) return { ok: true as const, value: error as Refusal };
        if (error instanceof PeerlessWorkerCommandRefusal) {
          if (error.status.result) return { ok: true as const, value: error.status.result as XlsxEditResult };
          const code = error.status.reason;
          if (code && EDIT_FAILURES.has(code) && opened) return { ok: true as const, value: {
            ok: false as const, version: opened.session.confirmedBasis.version,
            failure: { code: code as XlsxEditFailureCode, message: error.message },
          } };
        }
        return admissionFailure<XlsxEditResult | Refusal>(error);
      }
    },
    commands: options.commands,
    navigator: {
      selectCells: (identity, ...args) => navigate(identity, () => options.navigator.selectCells(identity, ...args)),
      scrollToCell: (identity, ...args) => navigate(identity, () => options.navigator.scrollToCell(identity, ...args)),
    },
  };
  return document;
}

export function createPeerlessPluginGeometry(
  layout: XlsxPluginLayout, painted: PaintedGrid, canvas: Element, layer: HTMLElement,
  current: () => { version: string | null; painted: PaintedGrid | null }
): XlsxPluginGeometry {
  return createPluginGeometry(layout, painted, canvas, layer, () => {
    const projection = current();
    return projection.version === layout.version && projection.painted === painted;
  });
}
