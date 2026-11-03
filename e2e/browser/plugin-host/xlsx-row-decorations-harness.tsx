import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
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
];

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
    void fetch(workbookUrl)
      .then((response) => response.arrayBuffer())
      .then((bytes) => setFile(new Uint8Array(bytes)));
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
