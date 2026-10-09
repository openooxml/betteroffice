import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { initWasm, openWorkbook } from '@betteroffice/xlsx';
import {
  XlsxEditor,
  defineXlsxPlugin,
  type XlsxEditorApi,
  type XlsxPluginContext,
  type XlsxPluginNavigationResult,
  type XlsxPluginSelection,
} from '@betteroffice/xlsx-react';
import workbookUrl from '../../../packages/xlsx/test-fixtures/sample.xlsx?url';

interface ProbeWindow {
  editor: XlsxEditorApi | null;
  navigation: XlsxPluginNavigationResult | null;
  selection: XlsxPluginSelection;
  marks: number;
}

const probe: ProbeWindow = { editor: null, navigation: null, selection: null, marks: 0 };
(window as unknown as { __probe: ProbeWindow }).__probe = probe;

type ProbeContext = XlsxPluginContext<null>;

function ProbePanel({ context }: { context: ProbeContext }) {
  return (
    <div>
      <button
        type="button"
        data-testid="probe-navigate"
        onClick={() =>
          void context.run(async (action) => {
            probe.navigation = await action.navigation.selectCells(
              {
                sheetId: 'sheet:1',
                selection: { anchor: { row: 3, col: 2 }, focus: { row: 1, col: 1 } },
              },
              { expectVersion: action.snapshot.version }
            );
          })
        }
      >
        Select on the second sheet
      </button>
      <input data-testid="probe-input" aria-label="Probe note" />
      {createPortal(
        <input
          data-testid="portal-input"
          aria-label="Portal note"
          style={{ position: 'fixed', right: 8, bottom: 8 }}
        />,
        document.body
      )}
    </div>
  );
}

const alignment = defineXlsxPlugin<null>({
  id: 'probe.alignment',
  createState: () => null,
  onEvent(_context, event) {
    if (event.type === 'selection-change') probe.selection = event.selection;
  },
  panel: { title: 'Probe', placement: 'right', render: ProbePanel },
  overlay: ({ context, geometry }) => {
    const selection = context.snapshot.selection;
    const cells = selection?.cells;
    const focus = cells?.focus;
    const cell =
      selection && cells
        ? geometry.getRangeRect({
            sheetId: selection.sheetId,
            range: {
              top: Math.min(cells.anchor.row, cells.focus.row),
              left: Math.min(cells.anchor.col, cells.focus.col),
              bottom: Math.max(cells.anchor.row, cells.focus.row),
              right: Math.max(cells.anchor.col, cells.focus.col),
            },
          })
        : null;
    const corner = geometry.getCellRect({ sheetId: geometry.layout.sheetId, row: 0, col: 0 });
    const target = geometry.getCellRect({ sheetId: 'sheet:0', row: 4, col: 2 });
    return (
      <>
        {corner && (
          <input
            data-testid="overlay-input"
            aria-label="Overlay note"
            style={{
              position: 'absolute',
              left: corner.x + 4,
              top: corner.y + 2,
              width: 120,
              pointerEvents: 'auto',
            }}
          />
        )}
        {target && (
          <div
            data-probe-target=""
            style={{
              position: 'absolute',
              left: target.x,
              top: target.y,
              width: target.width,
              height: target.height,
              background: 'rgba(0, 128, 255, 0.08)',
            }}
          />
        )}
        {cell && (
          <div
            data-probe-cell={`${focus!.row}:${focus!.col}`}
            data-probe-sheet={geometry.layout.sheetId}
            data-zoom={geometry.layout.zoom}
            style={{
              position: 'absolute',
              left: cell.x,
              top: cell.y,
              width: cell.width,
              height: cell.height,
              outline: '1px solid rgba(255, 0, 0, 0.4)',
            }}
          />
        )}
      </>
    );
  },
  commands: [
    {
      id: 'mark',
      label: 'Probe mark',
      mutatesDocument: false,
      shortcuts: ['Mod+Shift+K'],
      execute() {
        probe.marks += 1;
        return { ok: true, status: 'executed' };
      },
    },
  ],
  toolbar: ['mark'],
});

const second = defineXlsxPlugin({
  id: 'probe.second',
  createState: () => null,
  panel: {
    title: 'Second',
    placement: 'right',
    render: () => <p data-testid="second-panel">Second panel</p>,
  },
});

const PLUGINS = [alignment, second];

/** The sample workbook with its first two rows and first column frozen. */
async function frozenWorkbook(): Promise<Uint8Array> {
  await initWasm();
  const bytes = new Uint8Array(await (await fetch(workbookUrl)).arrayBuffer());
  const workbook = openWorkbook(bytes);
  try {
    workbook.applyOps([
      {
        type: 'setFreezePane',
        sheet: 0,
        pane: { rows: 2, cols: 1, top_left: { row: 2, col: 1 } },
      },
    ]);
    return workbook.save();
  } finally {
    workbook.dispose();
  }
}

function Harness() {
  const [file, setFile] = useState<Uint8Array | null>(null);
  useEffect(() => {
    void frozenWorkbook().then(setFile);
  }, []);
  if (!file) return null;
  return (
    <div style={{ height: '100%' }}>
      <XlsxEditor
        file={file}
        plugins={PLUGINS}
        onReady={(api) => {
          probe.editor = api;
        }}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
