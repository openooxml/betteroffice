import type { Proposal, ProposalEdit, WorkbookHandle } from '@betteroffice/xlsx';
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

/** @experimental */
export interface XlsxAsyncEditorApi {
  readonly hydrated: boolean;
  readonly failure: Error | null;
  whenHydrated(): Promise<void>;
  recoverySave(): Promise<{ bytes: Uint8Array; recovery: true }>;
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
