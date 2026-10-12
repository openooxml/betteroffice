import type { WorkbookHandle } from '@betteroffice/xlsx';
import { XlsxAsyncApiSyncAccessError } from './asyncApiErrors';
import { XlsxCommandAdmissionError, xlsxCommandController } from '../commands/createXlsxCommandStore';
import type { XlsxCommandStore } from '../commands/types';
import type { PeerlessWorkerInputCoordinator } from '../commands/peerlessWorkerInputCoordinator';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import type { AsyncCollaborationReplica } from '@betteroffice/xlsx/collaboration';

/** @experimental */
export interface XlsxAsyncEditorApi {
  readonly hydrated: boolean;
  readonly failure: Error | null;
  whenHydrated(): Promise<void>;
  recoverySave(): Promise<{ bytes: Uint8Array; recovery: true }>;
  readonly commands: XlsxCommandStore;
  readonly collaboration: AsyncCollaborationReplica | null;
  focus(): void;
  renderPng(...args: Parameters<WorkbookHandle['renderPng']>): Promise<ReturnType<WorkbookHandle['renderPng']>>;
  renderRangePng(...args: Parameters<WorkbookHandle['renderRangePng']>): Promise<ReturnType<WorkbookHandle['renderRangePng']>>;
  exportStructured(...args: Parameters<WorkbookHandle['exportStructured']>): Promise<ReturnType<WorkbookHandle['exportStructured']>>;
  exportMarkdown(...args: Parameters<WorkbookHandle['exportMarkdown']>): Promise<ReturnType<WorkbookHandle['exportMarkdown']>>;
  printDisplayList(...args: Parameters<WorkbookHandle['printDisplayList']>): Promise<ReturnType<WorkbookHandle['printDisplayList']>>;
}

export { XlsxAsyncApiSyncAccessError } from './asyncApiErrors';

const legacyAlternatives: Readonly<Record<string, string>> = {
  save: 'commands.execute("save", null)', saveAsync: 'commands.execute("save", null)', handle: 'commands',
  cell: 'plugin.read.readCells', cellAsync: 'plugin.read.readCells',
  rangeCells: 'plugin.read.readCells', rangeCellsAsync: 'plugin.read.readCells',
  readCells: 'plugin.read.readCells', readCellsSync: 'plugin.read.readCells',
  editCell: 'plugin.edits.applyEdits', editCellAsync: 'plugin.edits.applyEdits', applyEdits: 'plugin.edits.applyEdits',
  validateEdits: 'plugin.read.validateEdits', version: 'plugin.read.version', findText: 'plugin.read.findText',
  selectCells: 'plugin.navigation.selectCells', selectCellsAsync: 'plugin.navigation.selectCells',
  clearSelection: 'plugin.navigation.selectCells', getPositionAtPoint: 'plugin.geometry.getPositionAtPoint',
  undo: 'commands.execute("undo", null)', redo: 'commands.execute("redo", null)',
  refreshProposals: 'commands.execute("proposalsPanel", { open: true })',
  flush: 'commands.execute("save", null)', flushPendingInput: 'commands.execute("save", null)',
};

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
  return new Proxy(api, {
    get(target, key, receiver) {
      const alternative = typeof key === 'string' && Object.hasOwn(legacyAlternatives, key) ? legacyAlternatives[key] : null;
      if (alternative) throw new XlsxAsyncApiSyncAccessError(String(key), alternative);
      return Reflect.get(target, key, receiver);
    },
  });
}
