import {
  createSessionHost,
  transferable,
  type MethodHandlers,
  type SessionHost,
  type SessionTransport,
} from '../../../../shared/office-session';
import {
  initWasm, openWorkbook, workbookDisplayListJson,
  type SheetInfo, type WorkbookHandle,
} from '../wasm/loader';
import {
  WORKBOOK_SESSION_POLICIES,
  type WorkbookSessionEvents,
  type WorkbookSessionMethods,
  type WorkbookSheetSummary,
} from './methods';

type Events = { [K in keyof WorkbookSessionEvents]: WorkbookSessionEvents[K] };

function sheets(info: SheetInfo): WorkbookSheetSummary[] {
  return info.sheetIds.map((id, index) => ({ id, index, name: info.sheetNames[index] }));
}

export function createWorkbookSessionHost(
  transport: SessionTransport,
  options: { initWasm?: (source?: ArrayBuffer | WebAssembly.Module) => Promise<void> } = {}
): SessionHost<Events> {
  let handle: WorkbookHandle | undefined;
  let disposed = false;
  let version = 0;
  let dirty = false;
  let epoch = 0;
  const encoder = new TextEncoder();

  function workbook(): WorkbookHandle {
    if (disposed) throw new Error('Workbook session is disposed');
    if (!handle) throw new Error('Workbook session is not open');
    return handle;
  }

  function dispose(): void {
    disposed = true;
    const opened = handle;
    handle = undefined;
    opened?.dispose();
  }

  function checkSheet(opened: WorkbookHandle, sheet: number): void {
    if (!Number.isInteger(sheet) || sheet < 0 || sheet >= opened.sheetCount()) {
      throw new RangeError('Sheet index is out of range');
    }
  }

  const handlers: MethodHandlers<WorkbookSessionMethods, null> = {
    async open(_, bytes, input = {}) {
      if (disposed) throw new Error('Workbook session is disposed');
      if (handle) throw new Error('Workbook session is already open');
      await (options.initWasm ?? initWasm)(input.wasm);
      if (disposed) throw new Error('Workbook session is disposed');
      const opened = openWorkbook(new Uint8Array(bytes), {
        collaborative: input.collaborative,
        clientId: input.clientId,
      });
      try {
        const info = opened.sheetInfo();
        const summaries = sheets(info);
        handle = opened;
        return { format: 'xlsx', stage: 'ready', version, dirty,
          sheets: summaries, activeSheet: info.activeSheet };
      } catch (error) {
        opened.dispose();
        throw error;
      }
    },
    version: () => workbook().version(),
    readCells: (_, request) => workbook().readCells(request),
    findText: (_, request) => workbook().findText(request),
    validateEdits: (_, request) => workbook().validateEdits(request),
    applyEdits(_, request) {
      const result = workbook().applyEdits(request);
      if (result.ok && result.applied) {
        version += 1;
        dirty = true;
        host.emit('changed', { version, dirty });
      }
      return result;
    },
    frame(_, viewport, options = {}) {
      const opened = workbook();
      const sheet = options.sheet === undefined ? opened.sheetInfo().activeSheet : options.sheet;
      checkSheet(opened, sheet);
      const json = workbookDisplayListJson(opened, viewport, options.sheet);
      const mergedRanges = opened.visibleMergedRanges(sheet, viewport);
      const buffer = encoder.encode(json).buffer;
      epoch += 1;
      return transferable({
        displayList: buffer, version: opened.version(), epoch, sheet, viewport, mergedRanges,
      }, [buffer]);
    },
    sheetView(_, sheet) {
      const opened = workbook();
      checkSheet(opened, sheet);
      const { contentWidth, contentHeight, frozenRows, frozenCols, initialScrollX, initialScrollY } =
        opened.sheetInfoFor(sheet);
      const edge = opened.cellRect(sheet, Math.max(0, frozenRows - 1), Math.max(0, frozenCols - 1));
      return {
        sheet, version: opened.version(), contentWidth, contentHeight, frozenRows, frozenCols,
        initialScrollX, initialScrollY,
        frozenWidth: frozenCols === 0 ? 0 : edge.x + edge.w,
        frozenHeight: frozenRows === 0 ? 0 : edge.y + edge.h,
      };
    },
    cellGeometry(_, sheet, row, col) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return {
        sheet, version: opened.version(), rect: opened.cellRect(sheet, row, col),
        scrollPosition: opened.cellPosition(sheet, row, col),
      };
    },
    cellInputs(_, sheet, range) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return { sheet, version: opened.version(), cells: opened.rangeCells(sheet, range) };
    },
    sheets: () => sheets(workbook().sheetInfo()),
    calculationStatus: () => workbook().calculationStatus(),
    save() {
      const bytes = workbook().save();
      const buffer = bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 &&
        bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer : new Uint8Array(bytes).buffer;
      return transferable(buffer, [buffer]);
    },
    dispose() {
      workbook();
      dispose();
    },
  };

  const host = createSessionHost<WorkbookSessionMethods, Events, null>(transport, {
    handlers, policies: WORKBOOK_SESSION_POLICIES, context: null, onDispose: dispose,
  });
  return host;
}
