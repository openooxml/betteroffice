import '@univerjs/preset-sheets-core/lib/index.css';
import * as presets from '@univerjs/presets';
import * as sheetsCore from '@univerjs/preset-sheets-core';
import enUS from '@univerjs/preset-sheets-core/locales/en-US';
import { errors, expose, largestCanvas, nextFrame, rectOf, workbookName } from '../common/probe';

const { createUniver, LocaleType, mergeLocales } = presets;
const any = (value: unknown) => value as any;

const response = await fetch(`./docs/${workbookName}.json`);
if (!response.ok) throw new Error(`fetch ${workbookName}: ${response.status}`);
const data = await response.json();
const container = document.getElementById('root')!;
const t0 = performance.now();
const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: { [LocaleType.EN_US]: mergeLocales(enUS) },
  presets: [sheetsCore.UniverSheetsCorePreset({ container })],
});
const workbook = any(univerAPI).createWorkbook(data);
const createMs = Math.round(performance.now() - t0);
const injector = any(univerAPI)._injector;
const sheet = () => any(univerAPI).getActiveWorkbook().getActiveSheet();
const used: Record<string, string> = {};
const lifecycle: Record<string, number> = {};
try {
  any(univerAPI).addEvent(any(univerAPI).Event.LifeCycleChanged, ({ stage }: { stage: number }) => {
    lifecycle[String(stage)] = Math.round(performance.now() - t0);
  });
} catch (error) { errors.push(`lifecycle: ${String(error)}`); }

function editorText(): string | null {
  try {
    const Bridge = any(sheetsCore).IEditorBridgeService;
    const Instances = any(presets).IUniverInstanceService;
    const unitId = any(presets).DOCS_NORMAL_EDITOR_UNIT_ID_KEY ?? '__INTERNAL_EDITOR__DOCS_NORMAL';
    if (!Bridge || !Instances) return null;
    if (!injector.get(Bridge).isVisible().visible) return '';
    const stream: string = injector.get(Instances).getUnit(unitId)?.getBody?.()?.dataStream ?? '';
    return stream.replace(/[\r\n]+$/g, '');
  } catch { return null; }
}

const ready = (async () => {
  const deadline = performance.now() + 120_000;
  while (!largestCanvas(container) || largestCanvas(container)!.clientHeight < 100) {
    if (performance.now() > deadline) throw new Error('no grid canvas');
    await nextFrame();
  }
  await nextFrame();
})();

expose('univer-sheets', t0, {
  gridRect: () => { const canvas = largestCanvas(container); return canvas ? rectOf(canvas) : null; },
  async active() {
    const range = any(univerAPI).getActiveWorkbook()?.getActiveRange?.();
    return range ? { row: range.getRow(), col: range.getColumn() } : null;
  },
  async read(a1) { return sheet().getRange(a1).getValue(); },
  async setA1(value) {
    const range = sheet().getRange('A1');
    const out = range.setValue(value);
    if (out && typeof out.then === 'function') await out;
    return true;
  },
  async scrollToRow(row) {
    const ws = sheet();
    if (typeof ws.scrollToCell === 'function') { used.scroll = 'scrollToCell'; ws.scrollToCell(row, 0); return true; }
    used.scroll = 'activate';
    ws.getRange(row, 0).activate();
    return true;
  },
  async activateSheet(name) {
    const wb = any(univerAPI).getActiveWorkbook();
    const ws = wb?.getSheetByName(name);
    if (ws) wb.setActiveSheet(ws);
    return Boolean(ws);
  },
  editorText,
}, ready, { createMs, lifecycle, used, rows: data.sheets?.sheet1?.rowCount, workbookId: workbook?.getId?.() });
