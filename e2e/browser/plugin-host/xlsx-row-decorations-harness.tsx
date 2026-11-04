import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
import { initWasm, openWorkbook } from '@betteroffice/xlsx';
import {
  XlsxEditor,
  defineXlsxPlugin,
  type XlsxAsyncEditorApi,
  type XlsxEditorApi,
  type XlsxPluginRowDecoration,
} from '@betteroffice/xlsx-react';
import workbookUrl from '../../../packages/xlsx/test-fixtures/sample.xlsx?url';

interface RowProbe {
  editor: XlsxEditorApi | XlsxAsyncEditorApi | null;
  activations: string[];
}

const probe: RowProbe = { editor: null, activations: [] };
(window as unknown as { __rowProbe: RowProbe }).__rowProbe = probe;
const workerOnly = new URLSearchParams(location.search).get('engine') === 'worker';

const ITEMS: Omit<XlsxPluginRowDecoration, 'sheetId'>[] = [
  { id: 'a', row: 2, kind: 'badge', text: 'ABCD', tone: 'info', tooltip: 'First' },
  { id: 'b', row: 2, kind: 'badge', text: 'EFGH', tone: 'warning', tooltip: 'Second' },
  { id: 'c', row: 2, kind: 'badge', text: 'IJKL', tone: 'danger', tooltip: 'Third' },
  { id: 'd', row: 4, kind: 'dot', tone: 'success', tooltip: 'Dot' },
  { id: 'e', row: 4, kind: 'badge', text: 'OK' },
  { id: 'f', row: 6, kind: 'badge', text: 'WIDE' },
  ...Array.from({ length: 9_000 }, (_, index): Omit<XlsxPluginRowDecoration, 'sheetId'> => ({
    id: `short${index}`, row: 8, kind: index % 2 ? 'dot' : 'badge', text: index % 2 ? undefined : 'NOTE',
    tooltip: `Short ${index}`,
  })),
];

/** The sample workbook with row 9 at 7.5 pt. */
async function shortRowWorkbook(): Promise<Uint8Array> {
  await initWasm();
  const bytes = new Uint8Array(await (await fetch(workbookUrl)).arrayBuffer());
  const workbook = openWorkbook(bytes);
  try {
    workbook.applyOps([{ type: 'setRowHeight', sheet: 0, row: 8, height: 7.5 }]);
    return workbook.save();
  } finally {
    workbook.dispose();
  }
}

const rows = defineXlsxPlugin<null>({
  id: 'probe.rows',
  createState: () => null,
  onEvent(_context, event) {
    if (event.type === 'row-decoration-activate') probe.activations.push(event.id);
  },
  rowDecorations: (context) => ({
    version: context.snapshot.version,
    items: ITEMS.map((item) => ({ ...item, sheetId: 'sheet:0' })),
  }),
});

const PLUGINS = [rows];

function Harness() {
  const [file, setFile] = useState<Uint8Array | null>(null);
  useEffect(() => {
    void shortRowWorkbook().then(setFile);
  }, []);
  if (!file) return null;
  return (
    <div style={{ height: '100%' }}>
      <XlsxEditor
        {...(workerOnly ? {} : { legacySyncApi: true, experimentalWorkerOpen: false })}
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
