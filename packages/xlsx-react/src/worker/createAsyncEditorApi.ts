import type {
  Proposal, ProposalEdit, WorkbookHandle, XlsxReadRequest, XlsxReadResult, XlsxFindRequest,
  XlsxFindResult, XlsxEditRequest, XlsxEditResult, XlsxValidationResult,
} from '@betteroffice/xlsx';
import type { XlsxPluginDocumentAccess } from '../plugins/createPluginDocumentAccess';
import type { XlsxAdmission } from '../plugins/createPluginClients';
import { guardAsyncEditorApi } from './guardAsyncEditorApi';
import { XlsxCommandAdmissionError, xlsxCommandController } from '../commands/createXlsxCommandStore';
import type { XlsxCommandFailureCode, XlsxCommandResult, XlsxCommandStore } from '../commands/types';
import { commandReason } from '../commands/evaluate';
import { PeerlessWorkerCommandRefusal, type PeerlessWorkerInputCoordinator } from '../commands/peerlessWorkerInputCoordinator';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import type { AsyncCollaborationReplica } from '@betteroffice/xlsx/collaboration';

export interface XlsxProposalRequest {
  agentId: string;
  note?: string | null;
  edits: ProposalEdit[];
}

export type XlsxProposalResult =
  | { ok: true; proposal: Proposal }
  | Extract<XlsxCommandResult, { ok: false }>;

/**
 * Host reads, edits and saves commit pending input first. Admission rejects with
 * XlsxCommandAdmissionError for input-failed, document-replaced or editor-unavailable.
 * Session failures reject with the session error.
 * @experimental
 */
export interface XlsxAsyncEditorApi {
  readonly hydrated: boolean;
  readonly failure: Error | null;
  whenHydrated(): Promise<void>;
  recoverySave(): Promise<{ bytes: Uint8Array; recovery: true }>;
  /** Commits pending input before reading the workbook version. */
  version(): Promise<string>;
  /** Reads cells after pending input commits. */
  readCells(request: XlsxReadRequest): Promise<XlsxReadResult>;
  /** Searches display text after pending input commits. */
  findText(request: XlsxFindRequest): Promise<XlsxFindResult>;
  /** Refuses with `read-only` while the editor is read-only. */
  validateEdits(request: XlsxEditRequest): Promise<XlsxValidationResult>;
  /** Applies against `expectVersion` after pending input commits; notifies once if applied. */
  applyEdits(request: XlsxEditRequest): Promise<XlsxEditResult>;
  /** Commits pending input, returns bytes and calls onSave without downloading. */
  save(): Promise<Uint8Array>;
  readonly commands: XlsxCommandStore;
  readonly collaboration: AsyncCollaborationReplica | null;
  /** Stages tracked cell edits after pending input drains; available on worker editors. */
  proposeChanges?(request: XlsxProposalRequest): Promise<XlsxProposalResult>;
  focus(): void;
  renderPng(...args: Parameters<WorkbookHandle['renderPng']>): Promise<ReturnType<WorkbookHandle['renderPng']>>;
  renderRangePng(...args: Parameters<WorkbookHandle['renderRangePng']>): Promise<ReturnType<WorkbookHandle['renderRangePng']>>;
  exportStructured(...args: Parameters<WorkbookHandle['exportStructured']>): Promise<ReturnType<WorkbookHandle['exportStructured']>>;
  exportMarkdown(...args: Parameters<WorkbookHandle['exportMarkdown']>): Promise<ReturnType<WorkbookHandle['exportMarkdown']>>;
  printDisplayList(...args: Parameters<WorkbookHandle['printDisplayList']>): Promise<ReturnType<WorkbookHandle['printDisplayList']>>;
}

export { XlsxAsyncApiSyncAccessError } from './asyncApiErrors';

export interface XlsxAsyncExportBridge {
  editorDocument?(): XlsxPluginDocumentAccess | null;
  saved?(bytes: Uint8Array): void;
  exportSheet(): number;
  exportCoordinator(): PeerlessWorkerInputCoordinator | null;
}

export function createAsyncEditorApi(
  session: PeerlessWorkbookSession,
  commands: XlsxCommandStore,
  bridge: () => WorkerEditorApiBridge & Partial<XlsxAsyncExportBridge>
): XlsxAsyncEditorApi {
  const assertCurrent = () => {
    if (!session.current || session.retiring) throw new XlsxCommandAdmissionError('document-replaced');
  };
  const exported = <T>(capture: () => () => Promise<T>): Promise<T> => {
    try {
      assertCurrent();
      if (session.failure) throw session.failure;
      if (!session.ready) throw new XlsxCommandAdmissionError('editor-unavailable');
      const operation = capture();
      const input = bridge().exportCoordinator?.();
      if (!input) throw new XlsxCommandAdmissionError('editor-unavailable');
      return input.runAfterPendingInput(async () => {
        assertCurrent();
        const result = await operation();
        assertCurrent();
        return result;
      }, true);
    } catch (error) { return Promise.reject(error); }
  };
  const replica = session.session.replica?.async;
  const assertProposalReady = () => {
    assertCurrent();
    if (session.failure || !session.ready) throw new XlsxCommandAdmissionError('editor-unavailable');
  };
  const proposalRefusal = (code: XlsxCommandFailureCode): XlsxProposalResult => ({
    ok: false, failure: commandReason(code, null),
  });
  const call = async <T>(operation: () => Promise<T>): Promise<T> => {
    assertCurrent();
    if (session.failure) throw session.failure;
    let result: T;
    try { result = await operation(); }
    catch (error) {
      assertCurrent();
      if (session.failure) throw session.failure;
      throw error;
    }
    assertCurrent();
    if (session.failure) throw session.failure;
    return result;
  };
  const editorDocument = () => {
    assertCurrent();
    if (session.failure) throw session.failure;
    if (!session.ready) throw new XlsxCommandAdmissionError('editor-unavailable');
    const document = bridge().editorDocument?.();
    if (!document) throw new XlsxCommandAdmissionError('editor-unavailable');
    if (document.identity() !== session) throw new XlsxCommandAdmissionError('document-replaced');
    return document;
  };
  const admittedValue = <T>(admitted: XlsxAdmission<T>): T => {
    if (admitted.ok) return admitted.value;
    throw admitted.error instanceof XlsxCommandAdmissionError
      ? admitted.error : new XlsxCommandAdmissionError(admitted.code);
  };
  const hostRead = <T>(operation: (document: XlsxPluginDocumentAccess) => Promise<T>): Promise<T> => call<T>(async () => {
    const document = editorDocument();
    return await admittedValue(await document.admit(session, () => operation(document), true));
  });
  const collaboration: AsyncCollaborationReplica | null = replica ? {
    clientId: replica.clientId,
    encodeStateVector: () => call(() => session.session.call.collaborationStateVector()),
    encodeStateAsUpdate: (vector) => {
      const owned = vector?.slice();
      return call(() => session.session.call.collaborationStateAsUpdate(owned));
    },
    applyUpdate: (update) => {
      const owned = update.slice();
      return call(() => replica.applyUpdate(owned));
    },
    ...replica.applyUpdateChunks && {
      applyUpdateChunks: (updates: readonly Uint8Array[]) => {
        const owned = updates.map((update) => update.slice());
        return call(() => replica.applyUpdateChunks!(owned));
      },
    },
    onUpdate(listener) {
      assertCurrent();
      return replica.onUpdate((update, origin) => {
        if (session.current && !session.retiring && !session.failure) listener(update, origin);
      });
    },
  } : null;
  const api: XlsxAsyncEditorApi = {
    get hydrated() { return session.ready; },
    get failure() { return session.failure; },
    whenHydrated: () => session.whenReady(),
    recoverySave: () => session.recoverySave(),
    version: () => hostRead(async (document) => {
      const result = await document.readVersion(session);
      if (!result.ok) throw new XlsxCommandAdmissionError('input-failed');
      return result.version;
    }),
    readCells(request) {
      const captured = structuredClone(request);
      return hostRead(async (document) => {
        const result = await document.readCells(session, captured);
        if (!result.ok && !('version' in result)) throw new XlsxCommandAdmissionError('input-failed');
        return result;
      });
    },
    findText(request) {
      const captured = structuredClone(request);
      return hostRead(async (document) => {
        const result = await document.findText(session, captured);
        if (!result.ok && !('version' in result)) throw new XlsxCommandAdmissionError('input-failed');
        return result;
      });
    },
    validateEdits(request) {
      const captured = structuredClone(request);
      return hostRead(async (document) => {
        const result = await document.validateEdits(session, captured);
        if (!result.ok && !('version' in result)) throw new XlsxCommandAdmissionError('input-failed');
        return result;
      });
    },
    applyEdits(request) {
      const captured = structuredClone(request);
      return call(async () => {
        const document = editorDocument();
        if (!document.admitEdits) throw new XlsxCommandAdmissionError('editor-unavailable');
        return admittedValue(await document.admitEdits<never>(session, captured, () => null, (write) => write()));
      });
    },
    save: () => call(async () => {
      if (!session.ready) throw new XlsxCommandAdmissionError('editor-unavailable');
      const input = bridge().exportCoordinator?.();
      if (!input) throw new XlsxCommandAdmissionError('editor-unavailable');
      const bytes = await input.save();
      assertCurrent();
      if (session.failure) throw session.failure;
      bridge().saved?.(bytes);
      return bytes;
    }),
    commands: xlsxCommandController(commands)!.guarded(assertCurrent),
    collaboration,
    async proposeChanges(request) {
      try {
        assertProposalReady();
        const captured = structuredClone(request);
        const input = bridge().exportCoordinator?.();
        if (!input) throw new XlsxCommandAdmissionError('editor-unavailable');
        return await input.runAfterPendingInput(async () => {
          assertProposalReady();
          if (input.readOnly) return proposalRefusal('read-only');
          const ticket = input.command({ kind: 'propose', payload: {
            args: [captured.agentId, captured.note ?? null, captured.edits],
          } });
          const status = await ticket.result;
          assertProposalReady();
          return { ok: true as const, proposal: status.result as Proposal };
        }, 'posted');
      } catch (error) {
        if (!session.current || session.retiring) return proposalRefusal('document-replaced');
        if (session.failure || !session.ready) return proposalRefusal('editor-unavailable');
        if (error instanceof PeerlessWorkerCommandRefusal) {
          if (error.status.reason === 'read-only') return proposalRefusal('read-only');
          if (error.status.reason === 'unsupported') return proposalRefusal('unsupported-command');
          if (error.status.reason === 'invalid-command' || error.status.reason === 'invalid-step') {
            return proposalRefusal('invalid-arguments');
          }
          return proposalRefusal('input-failed');
        }
        if (error instanceof XlsxCommandAdmissionError) return proposalRefusal(error.code);
        return proposalRefusal('command-failed');
      }
    },
    renderPng: (viewport) => exported(() => {
      const captured = structuredClone(viewport);
      const sheet = bridge().exportSheet?.() ?? session.session.state.activeSheet;
      return () => session.session.call.renderPng(captured, { sheet });
    }),
    renderRangePng: (options) => exported(() => {
      const captured = structuredClone(options);
      const sheet = bridge().exportSheet?.() ?? session.session.state.activeSheet;
      return () => session.session.call.renderRangePng(captured, { sheet });
    }),
    exportStructured: (options) => exported(() => {
      const captured = structuredClone(options);
      return () => session.session.call.exportStructured(captured);
    }),
    exportMarkdown: (options, markdownOptions) => exported(() => {
      const captured = structuredClone([options, markdownOptions] as const);
      return () => session.session.call.exportMarkdown(...captured);
    }),
    printDisplayList: (...args) => exported(() => {
      const captured = structuredClone(args);
      return () => session.session.call.printDisplayList(...captured);
    }),
    focus() {
      assertCurrent();
      if (!session.failure) bridge().focus();
    },
  };
  return guardAsyncEditorApi(api);
}
