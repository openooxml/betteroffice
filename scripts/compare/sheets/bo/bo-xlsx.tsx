import { createRoot } from 'react-dom/client';
import { XlsxEditor, defineXlsxPlugin, type XlsxPluginContext, type XlsxPluginError } from '@betteroffice/xlsx-react';
import { errors, expose, rectOf, workbookName } from '../common/probe';

let context: XlsxPluginContext<null> | null = null;
const plugin = defineXlsxPlugin({ id: 'cmp', createState: () => null, initialize(value) { context = value; } });
async function use<T>(action: (current: XlsxPluginContext<null>) => Promise<T>): Promise<T> {
  if (!context) throw new Error('plugin unavailable');
  let out!: T;
  await context.run(async (current) => { out = await action(current); });
  return out;
}
const A1 = (a1: string, sheetId: string) => ({ sheetId, range: { kind: 'a1' as const, a1 } });

const response = await fetch(`./docs/${workbookName}.xlsx`);
if (!response.ok) throw new Error(`fetch ${workbookName}: ${response.status}`);
const file = new Uint8Array(await response.arrayBuffer());
const t0 = performance.now();
let finishReady!: () => void;
const ready = new Promise<void>((done) => { finishReady = done; });
const scroll = () => document.querySelector('[data-testid="xlsx-scroll"]');

expose('bo-xlsx', t0, {
  gridRect: () => { const element = scroll(); return element ? rectOf(element) : null; },
  active: () => use(async (current) => current.snapshot.selection.cells?.anchor ?? null),
  read: (a1) => use(async (current) => {
    const result = await current.read.readCells({ ranges: [A1(a1, current.snapshot.selection.sheetId)] });
    if (!result.ok) return { refused: result };
    return result.ranges[0]?.cells[0]?.[0]?.value;
  }),
  setA1: (value) => use(async (current) => {
    if (!current.edits) throw new Error('no edit grant');
    const result = await current.edits.applyEdits({ expectVersion: current.snapshot.version, steps: [
      { op: 'setCellInputs', target: A1('A1', current.snapshot.selection.sheetId), inputs: [[String(value)]] },
    ] });
    if (!result.ok) throw new Error(`applyEdits refused: ${JSON.stringify(result).slice(0, 200)}`);
    return true;
  }),
  scrollToRow: (row) => use(async (current) => {
    const result = await current.navigation.scrollToCell({ sheetId: current.snapshot.selection.sheetId, row, col: 0 },
      { expectVersion: current.snapshot.version, align: 'start' });
    if (!result.ok) throw new Error(`scrollToCell refused: ${JSON.stringify(result).slice(0, 200)}`);
    return true;
  }),
  activateSheet: async (name) => {
    const tab = [...document.querySelectorAll<HTMLElement>('[role="tab"]')].find((t) => t.textContent === name);
    tab?.click();
    return Boolean(tab);
  },
  editorText: () => (document.querySelector('[data-testid="xlsx-cell-editor"]') as HTMLInputElement | null)?.value ?? '',
}, ready, { bytes: file.byteLength });

createRoot(document.getElementById('root')!).render(<XlsxEditor file={file} readOnly={false}
  plugins={[plugin]} pluginGrants={{ cmp: { document: 'write', editBatches: true } }}
  onError={(error: Error) => errors.push(error.message)}
  onPluginError={(error: XlsxPluginError) => errors.push(JSON.stringify(error).slice(0, 300))}
  onReady={(api: { whenHydrated?: () => Promise<void> }) => {
    void (async () => {
      try {
        if (api && typeof api.whenHydrated === 'function') await api.whenHydrated();
        finishReady();
      } catch (error) { errors.push(String(error)); }
    })();
  }} />);
