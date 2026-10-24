import type { ChartEdit, ChartInfo, ChartMutation } from '@betteroffice/xlsx';
import { DocumentToolError } from './types';
import { integer, textWindow } from './prototype';
import { batch, failure } from './xlsx-operations';
import type { XlsxAgentWorkbook } from './xlsx';

export type XlsxChartEdit = ({ id: string; sheet?: string; chart?: never } | { sheet: string; chart: number; id?: never }) & ChartEdit;
export type XlsxChartTextField = 'title' | 'type' | 'name' | 'nameRef' | 'categoryRef' | 'valueRef';
export interface XlsxChartTextRead { id?: string; sheet?: string; chart?: number; field?: XlsxChartTextField; series?: number; start?: number; length?: number }

function resolveChart(book: XlsxAgentWorkbook, target: { id?: string; sheet?: string; chart?: number }) {
  const sheets = target.sheet === undefined ? book.catalog() : [book.resolveSheet(target.sheet)];
  const matches = sheets.flatMap(sheet => book.handle.charts(sheet.index).map(chart => ({ sheet, chart })));
  if (target.id !== undefined && target.chart !== undefined) throw new DocumentToolError('INVALID_ARGUMENT', 'Address a chart with id, or with sheet and chart (zero-based), exclusively.');
  const found = target.id !== undefined ? matches.filter(item => item.chart.id === target.id) : matches.filter(item => item.sheet.sheetId === target.sheet || item.sheet.name === target.sheet).filter(item => item.chart.index === target.chart);
  if (found.length !== 1) throw new DocumentToolError('UNKNOWN_CHART', `Chart ${JSON.stringify(target.id ?? target.chart)} is unknown or ambiguous. Use id from xlsx_list_charts, or sheet and zero-based chart. Valid choices: ${matches.slice(0, 30).map(item => `${item.sheet.sheetId} (${item.sheet.name}), chart ${item.chart.index}`).join('; ')}.`);
  if (!found[0].chart.editable) throw new DocumentToolError('UNSUPPORTED', `Chart ${found[0].chart.index} on ${found[0].sheet.name} cannot be read: ${found[0].chart.error}. Choose a chart with editable true from xlsx_list_charts.`);
  return found[0];
}

function seriesIndex(chart: ChartInfo, value: number) {
  if (!chart.series.length) throw new DocumentToolError('UNKNOWN_SERIES', 'series has no valid indices on this chart. Choose a chart with series from xlsx_list_charts.');
  return integer(value, 0, chart.series.length - 1, 'series');
}

const preview = (value: string | null, length: number) => value === null ? null : textWindow(value, { length }).text;

export function listCharts(book: XlsxAgentWorkbook, options: { sheet: string; offset?: number; limit?: number; seriesOffset?: number }) {
  const sheet = book.resolveSheet(options.sheet);
  const all = book.handle.charts(sheet.index);
  const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
  const limit = integer(options.limit ?? 10, 1, 100, 'limit');
  const seriesOffset = integer(options.seriesOffset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'seriesOffset');
  const items = [];
  let size = 0;
  for (const chart of all.slice(offset, offset + limit)) {
    const series = [];
    const base = { ...chart, title: preview(chart.title, 400), type: preview(chart.type, 80)!, sheet: sheet.sheetId, seriesCount: chart.series.length };
    let itemSize = JSON.stringify({ ...base, series: [] }).length;
    for (const source of chart.series.slice(seriesOffset, seriesOffset + 20)) {
      const item = { ...source, name: preview(source.name, 120), nameRef: preview(source.nameRef, 400), categoryRef: preview(source.categoryRef, 400), valueRef: preview(source.valueRef, 400), truncated: (source.name?.length ?? 0) > 120 || [source.nameRef, source.categoryRef, source.valueRef].some(value => (value?.length ?? 0) > 400) };
      const cost = JSON.stringify(item).length;
      if (series.length && itemSize + cost > 10000) break;
      itemSize += cost;
      series.push(item);
    }
    const item = { ...base, series, nextSeriesOffset: seriesOffset + series.length < chart.series.length ? seriesOffset + series.length : null, truncated: (chart.title?.length ?? 0) > 400 || chart.type.length > 80 || series.some(item => item.truncated) };
    const cost = JSON.stringify(item).length;
    if (items.length && size + cost > 12000) break;
    size += cost;
    items.push(item);
  }
  return { version: book.handle.version(), items, total: all.length, nextOffset: offset + items.length < all.length ? offset + items.length : null };
}

export function readChartText(book: XlsxAgentWorkbook, options: XlsxChartTextRead) {
  const { sheet, chart } = resolveChart(book, options);
  const field = options.field ?? 'title';
  const fields = ['title', 'type', 'name', 'nameRef', 'categoryRef', 'valueRef'];
  if (!fields.includes(field)) throw new DocumentToolError('INVALID_ARGUMENT', `field must be one of ${fields.join(', ')}.`);
  const series = field === 'title' || field === 'type' ? undefined : seriesIndex(chart, options.series ?? 0);
  const value = field === 'title' ? chart.title : field === 'type' ? chart.type : chart.series[series!][field];
  const length = integer(options.length ?? 2000, 1, 2000, 'length');
  return { version: book.handle.version(), sheet: sheet.sheetId, id: chart.id, field, ...(series === undefined ? {} : { series }), present: value !== null, ...textWindow(value ?? '', { start: options.start, length }) };
}

export function editCharts(book: XlsxAgentWorkbook, options: { version: string; edits: XlsxChartEdit[] }) {
  book.checkVersion(options.version);
  batch(options.edits);
  const targets: Array<{ sheet: string; chart: number; id: string }> = [];
  const edits: ChartMutation[] = options.edits.map((edit, index) => {
    try {
      const { sheet, chart } = resolveChart(book, edit);
      targets.push({ sheet: sheet.sheetId, chart: chart.index, id: chart.id });
      if (edit.action === 'set_chart_title') return { sheet: sheet.index, chart: chart.index, action: edit.action, title: edit.title };
      if (edit.action !== 'set_series') throw new DocumentToolError('INVALID_ARGUMENT', 'action must be set_chart_title or set_series.');
      seriesIndex(chart, edit.series);
      if (edit.name === undefined && edit.categories === undefined && edit.values === undefined) throw new DocumentToolError('INVALID_ARGUMENT', 'set_series needs at least one field: name, categories, values.');
      return { sheet: sheet.index, chart: chart.index, action: edit.action, series: edit.series, ...(edit.name === undefined ? {} : { name: edit.name }), ...(edit.categories === undefined ? {} : { categories: edit.categories }), ...(edit.values === undefined ? {} : { values: edit.values }) };
    } catch (error) { failure(error, index); }
  });
  let result;
  try { book.checkVersion(options.version); result = book.handle.editCharts(edits); } catch (error) { failure(error); }
  if (result.applied) book.invalidateCells();
  return { baseVersion: options.version, version: book.handle.version(), applied: result.applied, items: targets.map((target, index) => ({ index, action: options.edits[index].action, sheet: target.sheet, chart: target.chart, ...(target.id.length <= 1024 ? { id: target.id } : {}) })) };
}
