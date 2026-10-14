import { XlsxAsyncApiSyncAccessError } from './asyncApiErrors';
import type { XlsxAsyncEditorApi } from './createAsyncEditorApi';

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

export function guardAsyncEditorApi(api: XlsxAsyncEditorApi): XlsxAsyncEditorApi {
  return new Proxy(api, {
    get(target, key, receiver) {
      const alternative = typeof key === 'string' && Object.hasOwn(legacyAlternatives, key) ? legacyAlternatives[key] : null;
      if (alternative) throw new XlsxAsyncApiSyncAccessError(String(key), alternative);
      return Reflect.get(target, key, receiver);
    },
  });
}
