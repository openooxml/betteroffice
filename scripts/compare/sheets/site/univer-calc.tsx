// Recalc page: loads a converted cache-cleared snapshot with forced initial computing, waits for the formula engine
// to go idle, then reads target cells through the facade. Exposes window.calc. Registers Univer's OSS table preset
// so converted Excel tables resolve structured references.
import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-table/lib/index.css';
import '@univerjs/preset-sheets-hyper-link/lib/index.css';
import * as presets from '@univerjs/presets';
import * as sheetsCore from '@univerjs/preset-sheets-core';
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table';
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link';
import enUS from '@univerjs/preset-sheets-core/locales/en-US';
import tableEnUS from '@univerjs/preset-sheets-table/locales/en-US';
import hyperlinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US';

const any = (value: unknown) => value as any;
const errors: string[] = [];
addEventListener('error', (event) => errors.push(String(event.message).slice(0, 300)));
addEventListener('unhandledrejection', (event) => errors.push(String(event.reason).slice(0, 300)));
const sleep = (ms: number) => new Promise((done) => setTimeout(done, ms));
const name = new URLSearchParams(location.search).get('wb')!;
const capMs = Number(new URLSearchParams(location.search).get('cap') ?? 120_000);
const mode = any(sheetsCore).CalculationMode?.FORCED ?? 0;

const ready = (async () => {
  const response = await fetch(`./docs/${name}.json`);
  if (!response.ok) throw new Error(`fetch ${name}: ${response.status}`);
  const data = await response.json();
  const t0 = performance.now();
  const { univerAPI } = presets.createUniver({
    locale: presets.LocaleType.EN_US,
    locales: { [presets.LocaleType.EN_US]: presets.mergeLocales(enUS, tableEnUS, hyperlinkEnUS) },
    presets: [sheetsCore.UniverSheetsCorePreset({ container: document.getElementById('root')!, formula: { initialFormulaComputing: mode } }), UniverSheetsTablePreset(), UniverSheetsHyperLinkPreset()],
  });
  const formula = any(univerAPI).getFormula();
  const state = { starts: 0, ends: 0, lastEvent: performance.now(), endStates: [] as unknown[] };
  formula.calculationStart(() => { state.starts++; state.lastEvent = performance.now(); });
  formula.calculationEnd((s: unknown) => { state.ends++; state.lastEvent = performance.now(); if (state.endStates.length < 5) state.endStates.push(s); });
  const workbook = any(univerAPI).createWorkbook(data);
  const createMs = performance.now() - t0;
  let timedOut = false;
  for (;;) {
    const now = performance.now();
    const idle = state.starts > 0 ? state.ends >= state.starts && now - state.lastEvent > 1000 : now - t0 > 8000;
    if (idle) break;
    if (now - t0 > capMs) { timedOut = true; break; }
    await sleep(100);
  }
  await new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done)));
  return { univerAPI, workbook, info: { createMs: Math.round(createMs), calcMs: Math.round(performance.now() - t0), timedOut, starts: state.starts, ends: state.ends, endStates: state.endStates } };
})();
ready.catch((error) => errors.push(`ready: ${String(error).slice(0, 300)}`));

function kind(cell: any) {
  if (!cell) return { kind: 'empty' };
  const v = cell.v;
  if (v === undefined || v === null) return { kind: 'empty' };
  if (cell.t === 3 || typeof v === 'boolean') return { kind: 'bool', value: v === true || v === 1 || v === 'TRUE' || v === '1' };
  if (typeof v === 'string' && /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|FIELD!|UNKNOWN!|BUSY!|CYCLE!)$/.test(v)) return { kind: 'error', value: v };
  if (typeof v === 'number' || cell.t === 2) { const n = Number(v); return Number.isFinite(n) ? { kind: 'number', value: n } : { kind: 'text', value: String(v) }; }
  return cell.viaDisplay ? { kind: 'text', value: String(v), viaDisplay: true } : { kind: 'text', value: String(v) };
}

/** targets: [{sheet: name, cells: [[row, col], ...]}]; returns values in the same order. */
async function read(targets: { sheet: string; cells: [number, number][] }[]) {
  const { workbook } = await ready;
  const out: unknown[][] = [];
  for (const group of targets) {
    const ws = workbook.getSheetByName(group.sheet);
    if (!ws) { out.push(group.cells.map(() => ({ kind: 'missing-sheet' }))); continue; }
    const unitId = workbook.getId(), sheetId = ws.getSheetId();
    let arrays: any = null;
    try { arrays = (await ready).univerAPI._injector.get(any(sheetsCore).FormulaDataModel).getArrayFormulaCellData(); } catch (error) { errors.push(`arrays: ${String(error).slice(0, 120)}`); }
    const maxRow = ws.getMaxRows() - 1, maxCol = ws.getMaxColumns() - 1;
    const inside = group.cells.filter(([r, c]) => r <= maxRow && c <= maxCol);
    if (inside.length < group.cells.length) errors.push(`${group.sheet}: ${group.cells.length - inside.length} targets outside ${maxRow + 1}x${maxCol + 1}`);
    if (!inside.length) { out.push(group.cells.map(() => ({ kind: 'empty', outside: true }))); continue; }
    const rows = inside.map(([r]) => r), cols = inside.map(([, c]) => c);
    const lo = (xs: number[]) => xs.reduce((m, x) => Math.min(m, x), Infinity), hi = (xs: number[]) => xs.reduce((m, x) => Math.max(m, x), -Infinity);
    const r0 = lo(rows), r1 = hi(rows), c0 = lo(cols), c1 = hi(cols);
    const values: unknown[] = [];
    const span = c1 - c0 + 1;
    const chunk = Math.max(1, Math.floor(400_000 / span));
    const byRow = new Map<number, [number, number][]>();
    for (const [r, c] of inside) { if (!byRow.has(r)) byRow.set(r, []); byRow.get(r)!.push([r, c]); }
    const grid = new Map<string, unknown>();
    for (let start = r0; start <= r1; start += chunk) {
      const end = Math.min(r1, start + chunk - 1);
      if (![...byRow.keys()].some((r) => r >= start && r <= end)) continue;
      const range = ws.getRange(start, c0, end - start + 1, span);
      const datas = range.getCellDatas();
      let shown: unknown[][] | null = null;
      for (let r = start; r <= end; r++) for (const [, c] of byRow.get(r) ?? []) {
        const raw = datas?.[r - start]?.[c - c0];
        let cell = raw;
        if (!raw || raw.v === undefined || raw.v === null) {
          const spill = arrays?.[unitId]?.[sheetId]?.[r]?.[c];
          if (spill && spill.v !== undefined && spill.v !== null) cell = spill;
          else {
            shown ??= range.getValues();
            const v = shown?.[r - start]?.[c - c0];
            if (v !== undefined && v !== null) cell = { v, viaDisplay: true };
          }
        }
        grid.set(`${r}:${c}`, kind(cell));
      }
    }
    for (const [r, c] of group.cells) values.push(grid.get(`${r}:${c}`) ?? { kind: 'empty', outside: r > maxRow || c > maxCol });
    out.push(values);
  }
  return out;
}

Object.assign(window, { calc: { ready: ready.then((x) => x.info), read, errors } });
