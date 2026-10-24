import { XlsxAsyncApiSyncAccessError } from './asyncApiErrors';
import type { XlsxAsyncEditorApi } from './createAsyncEditorApi';

const legacyAlternatives: Readonly<Record<string, string>> = {
  saveAsync: 'save()', handle: 'commands',
  cell: 'readCells(request)', cellAsync: 'readCells(request)',
  rangeCells: 'readCells(request)', rangeCellsAsync: 'readCells(request)',
  readCellsSync: 'readCells(request)',
  editCell: 'applyEdits(request)', editCellAsync: 'applyEdits(request)',
  selectCells: 'plugin.navigation.selectCells', selectCellsAsync: 'plugin.navigation.selectCells',
  clearSelection: 'plugin.navigation.selectCells', getPositionAtPoint: 'plugin.geometry.getPositionAtPoint',
  undo: 'commands.execute("undo", null)', redo: 'commands.execute("redo", null)',
  refreshProposals: 'commands.execute("proposalsPanel", { open: true })',
  flush: 'commands.execute("save", null)', flushPendingInput: 'commands.execute("save", null)',
};

export function guardAsyncEditorApi(api: XlsxAsyncEditorApi): XlsxAsyncEditorApi {
  return new Proxy(api, {
    get(target, key, receiver) {
      const alternative = typeof key === 'string' && Object.hasOwn(legacyAlternatives, key) ? legacyAlternatives[key] : null;
      if (alternative) throw new XlsxAsyncApiSyncAccessError(String(key), alternative);
      return Reflect.get(target, key, receiver);
    },
  });
}
