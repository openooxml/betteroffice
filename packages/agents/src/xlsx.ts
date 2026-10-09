import { initWasm, isPngExportAvailable, openWorkbook, type WorkbookHandle, type WorkbookEditPeer, type XlsxCellRead, type XlsxEditStep } from '@betteroffice/xlsx';
import { DocumentToolError, type GrepOptions, type TextEdit } from './types';
import { applyXlsxFormatting, editXlsxGrid, editXlsxSheets, findXlsx, outlineXlsx, previewXlsx, sortXlsx, writeXlsx, type XlsxFormatEdit, type XlsxGridEdit, type XlsxSheetEdit, type XlsxWriteEdit } from './xlsx-operations';
import { initialize, integer, plainText, PrototypeDocument, textWindow, unwrap, type PrototypeOptions } from './prototype';

export interface XlsxOptions extends PrototypeOptions { editPeer?: WorkbookEditPeer }

export interface CellEdit { cell: string; input: string }
export interface CellChange { ref: string; sheet: string; a1: string; before: XlsxCellRead; input: string }
interface CellTarget { sheet: string; a1: string; before: XlsxCellRead; version: string }

export function address(row: number, col: number) {
  let letters = '';
  for (let n = col + 1; n; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + (n - 1) % 26) + letters;
  return `${letters}${row + 1}`;
}

export function corners(range: string) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(range);
  if (!match) throw new DocumentToolError('INVALID_RANGE', 'Use one A1 cell or a rectangular A1:B2 range, without a sheet prefix.');
  const col = (letters: string) => [...letters.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const start = { row: Number(match[2]) - 1, col: col(match[1]) };
  const end = { row: Number(match[4] ?? match[2]) - 1, col: col(match[3] ?? match[1]) };
  if (start.row > end.row || start.col > end.col || end.row >= 1048576 || end.col >= 16384) throw new DocumentToolError('INVALID_RANGE', 'Range must be ordered and inside the Excel grid.');
  return { start, end };
}

export class XlsxAgentWorkbook extends PrototypeDocument<CellChange, XlsxEditStep> {
  readonly format = 'xlsx' as const;
  private readonly cells = new Map<string, CellTarget>();
  private readonly refs = new Map<string, { sheet: string; a1: string }>();
  private nextCell = 1;

  private sheetCatalog: Array<{ sheetId: string; name: string }> = [];
  private nextSheet = 0;

  private constructor(readonly handle: WorkbookHandle, readonly name = 'workbook.xlsx', private readonly ownsHandle = true, private readonly editPeer?: WorkbookEditPeer) {
    super();
    this.catalog();
  }

  static attach(handle: WorkbookHandle, options: XlsxOptions = {}) {
    const attached = options.editPeer ? new Proxy(handle, { get: (target, key) => key === 'applyEdits' || key === 'applyOps' ? options.editPeer![key] : Reflect.get(target, key) }) : handle;
    return new XlsxAgentWorkbook(attached, options.name, false, options.editPeer);
  }

  catalog() {
    this.alive();
    const info = this.handle.sheetInfo();
    this.sheetCatalog = info.sheetNames.map(name => this.sheetCatalog.find(sheet => sheet.name === name) ?? { sheetId: `sheet:${this.nextSheet++}`, name });
    return this.sheetCatalog.map((sheet, index) => ({ ...sheet, index, nativeId: info.sheetIds[index] }));
  }

  resolveSheet(sheet: string) {
    const catalog = this.catalog();
    const found = catalog.find(entry => entry.sheetId === sheet) ?? catalog.find(entry => entry.name === sheet);
    if (!found) throw new DocumentToolError('UNKNOWN_SHEET', `Unknown sheet ${JSON.stringify(sheet)}. Use a sheetId or exact name from xlsx_outline. Valid sheets: ${catalog.slice(0, 30).map(entry => `${entry.sheetId} (${entry.name})`).join(', ')}.`, { sheets: catalog.slice(0, 30).map(({ sheetId, name }) => ({ sheetId, name })), total: catalog.length });
    return found;
  }

  adoptSheets(sheets: Array<{ sheetId: string; name: string }>) {
    const assigned = new Map<string, string>();
    this.sheetCatalog = sheets.map(sheet => {
      const sheetId = sheet.sheetId.startsWith('added:') ? `sheet:${this.nextSheet++}` : sheet.sheetId;
      assigned.set(sheet.sheetId, sheetId);
      return { sheetId, name: sheet.name };
    });
    return assigned;
  }

  checkVersion(version: string) {
    this.alive();
    if (version !== this.handle.version()) throw new DocumentToolError('STALE_VERSION', 'Workbook changed. Read the range or outline again and use its version.', { version: this.handle.version() });
  }

  invalidateCells() { this.cells.clear(); this.refs.clear(); }

  override async export(id?: string) {
    await this.editPeer?.flush();
    return super.export(id);
  }

  outline(options: { offset?: number; limit?: number; sheet?: string; nameOffset?: number; tableOffset?: number } = {}) { return outlineXlsx(this, options); }
  writeRanges(options: { version: string; edits: XlsxWriteEdit[] }) { return writeXlsx(this, options); }
  editGrid(options: { version: string; edits: XlsxGridEdit[] }) { return editXlsxGrid(this, options); }
  editSheets(options: { version: string; edits: XlsxSheetEdit[] }) { return editXlsxSheets(this, options); }
  formatRanges(options: { version: string; edits: XlsxFormatEdit[] }) { return applyXlsxFormatting(this, options); }
  findReplace(options: { query: string; sheet?: string; range?: string; caseSensitive?: boolean; offset?: number; limit?: number; replacement?: string; version?: string }) { return findXlsx(this, options); }
  sortRange(options: { version: string; sheet: string; range: string; keys: Array<{ column: string; order: 'asc' | 'desc' }>; header?: boolean }) { return sortXlsx(this, options); }
  previewRange(options: { sheet: string; range: string; scale?: number }) { return previewXlsx(this, options); }

  static async open(bytes: Uint8Array, options: PrototypeOptions = {}) {
    await initialize('xlsx', initWasm, options.wasm);
    return new XlsxAgentWorkbook(openWorkbook(bytes), options.name);
  }

  overview() {
    this.alive();
    const read = unwrap(this.handle.readCells({ ranges: [] }));
    return { name: this.name, format: this.format, version: read.version, sheets: read.sheets.length, capabilities: { grep: 'literal', proposals: 'cell inputs and formulas', render: isPngExportAvailable(), export: true, batchEdits: true, formatting: true, structure: true, sort: true, attachment: 'WorkbookHandle, including an editor edit peer' }, workflow: 'xlsx_outline lists stable sheet IDs and version; xlsx_read_range reads A1 ranges; write tools require that version and return changed cells. xlsx_preview renders a range. office_export saves a new file.' };
  }

  list(options: { offset?: number; limit?: number; story?: string; headingsOnly?: boolean } = {}) {
    this.alive();
    if (options.headingsOnly) throw new DocumentToolError('UNSUPPORTED', 'XLSX outline lists sheets; omit headingsOnly.');
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 30, 1, 100, 'limit');
    const read = unwrap(this.handle.readCells({ ranges: [] }));
    const catalog = this.catalog();
    const nativeStory = options.story ? this.resolveSheet(options.story).nativeId : undefined;
    const sheets = read.sheets.filter(s => !nativeStory || s.sheetId === nativeStory);
    const items = sheets.map(sheet => ({ ...sheet, sheetId: catalog.find(entry => entry.nativeId === sheet.sheetId)!.sheetId }));
    return { version: read.version, items: items.slice(offset, offset + limit), nextOffset: offset + limit < sheets.length ? offset + limit : null, total: sheets.length };
  }

  readCells(options: { sheet: string; range: string; offset?: number; limit?: number }) {
    this.alive();
    const sheet = this.resolveSheet(options.sheet);
    const { start, end } = corners(options.range);
    const width = end.col - start.col + 1;
    const total = width * (end.row - start.row + 1);
    const offset = integer(options.offset ?? 0, 0, total, 'offset');
    const limit = integer(options.limit ?? 50, 1, 100, 'limit');
    const positions = [];
    for (let index = offset; index < Math.min(total, offset + limit); index++) {
      positions.push({ sheet: sheet.sheetId, a1: address(start.row + Math.floor(index / width), start.col + index % width) });
    }
    const read = this.readTargets(positions);

    const cells = [];
    let size = 0;
    for (const target of read.targets) {
      const visible = this.visible(target);
      const cost = JSON.stringify(visible).length;
      if (cells.length && size + cost > 16000) break;
      size += cost;
      cells.push(visible);
    }
    return { version: read.version, sheet: sheet.sheetId, range: `${address(start.row, start.col)}:${address(end.row, end.col)}`, cells, total, nextOffset: offset + cells.length < total ? offset + cells.length : null };
  }

  grep(options: GrepOptions) {
    this.checkQuery(options);
    if (options.caseSensitive === false) throw new DocumentToolError('UNSUPPORTED', 'XLSX search is case-sensitive. Omit caseSensitive or set it to true.');
    const found = unwrap(this.handle.findText({ text: options.query, sheetIds: options.story ? [this.resolveSheet(options.story).nativeId] : undefined, limit: 10000 }));
    const page = this.searchPage(options, found.matches.map(hit => ({ ref: this.ref(this.resolveNative(hit.cell.sheetId), hit.cell.a1), sheet: this.resolveNative(hit.cell.sheetId), a1: hit.cell.a1, text: hit.text.slice(0, 400), length: hit.text.length })), found.truncated);
    const read = this.readTargets(page.matches);
    return { ...page, matches: page.matches.map((hit, index) => ({ ...hit, cell: this.hold(read.targets[index]) })) };
  }

  read(ref: string, options: { start?: number; length?: number; field?: 'displayText' | 'formula' | 'value' } = {}) {
    this.alive();
    const at = this.refs.get(ref);
    if (!at) throw new DocumentToolError('UNKNOWN_REF', 'Use a ref from grep or readCells.');
    const target = this.cellAt(at.sheet, at.a1);
    const field = options.field ?? 'displayText';
    const text = field === 'formula' ? target.before.formula ?? '' : field === 'value' ? ('value' in target.before.value ? String(target.before.value.value) : '') : target.before.displayText;
    return { ...this.visible(target), field, ...textWindow(text, options) };
  }

  propose(_input: { author: string; note?: string; edits: TextEdit[] }): never {
    throw new DocumentToolError('UNSUPPORTED', 'XLSX edits use proposeCells / office_propose_cells with {cell, input}. Read cells first; input replaces the entire cell.');
  }

  proposeCells(input: { author: string; note?: string; edits: CellEdit[] }) {
    this.alive();
    if (!Array.isArray(input.edits) || input.edits.length < 1 || input.edits.length > 32) throw new DocumentToolError('INVALID_ARGUMENT', 'Supply 1–32 cell edits.');
    const version = this.currentVersion();
    const changes: CellChange[] = [];
    const steps: XlsxEditStep[] = [];
    for (const edit of input.edits) {
      const target = this.cells.get(edit.cell);
      if (!target) throw new DocumentToolError('UNKNOWN_CELL', 'Cell handle expired or belongs to another workbook. Read cells or grep again.');
      if (target.version !== version) throw new DocumentToolError('STALE_TARGET', 'Workbook changed. Read fresh cell handles.');
      plainText(edit.input);
      if (edit.input.length > 16000) throw new DocumentToolError('EDIT_LIMIT', 'Cell input must fit in 16000 UTF-16 units.');
      changes.push({ ref: this.ref(target.sheet, target.a1), sheet: target.sheet, a1: target.a1, before: target.before, input: edit.input });
      steps.push({ op: 'setCellInputs', target: { sheetId: this.resolveSheet(target.sheet).nativeId, range: { kind: 'a1', a1: target.a1 } }, inputs: [[edit.input]], expect: { cells: [[{ value: target.before.value, formula: target.before.formula, displayText: target.before.displayText }]] } });
    }
    return this.stage(input, changes, steps, version);
  }

  protected currentVersion() { return this.handle.version(); }
  protected validate(steps: XlsxEditStep[], version: string) { unwrap(this.handle.validateEdits({ expectVersion: version, steps, source: 'agent' })); }
  protected apply(steps: XlsxEditStep[], version: string) {
    const result = unwrap(this.handle.applyEdits({ expectVersion: version, steps, source: 'agent' }));
    return { applied: result.applied, version: result.version, changedSheets: result.changedSheets.length, calculation: { changedCells: result.calculation.changed.length, cycleCells: result.calculation.cycleCells.length, limitedCells: result.calculation.limitedCells.length, truncated: result.calculation.truncated } };
  }
  protected save(steps?: XlsxEditStep[]) {
    if (!steps) return this.handle.save();
    const draft = openWorkbook(this.handle.save());
    try { unwrap(draft.applyEdits({ expectVersion: draft.version(), steps, source: 'agent' })); return draft.save(); }
    finally { draft.dispose(); }
  }
  protected reopen(bytes: Uint8Array) { const reopened = openWorkbook(bytes); reopened.dispose(); }
  protected dispose() { this.cells.clear(); this.refs.clear(); if (this.ownsHandle) this.handle.dispose(); }

  private resolveNative(nativeId: string) { return this.catalog().find(sheet => sheet.nativeId === nativeId)!.sheetId; }

  private ref(sheet: string, a1: string) {
    const ref = `${sheet}!${a1}`;
    if (this.refs.size >= 10000 && !this.refs.has(ref)) this.refs.delete(this.refs.keys().next().value!);
    this.refs.set(ref, { sheet, a1 });
    return ref;
  }

  private cellAt(sheet: string, a1: string): CellTarget {
    return this.readTargets([{ sheet, a1 }]).targets[0];
  }

  private readTargets(positions: Array<{ sheet: string; a1: string }>) {
    const read = unwrap(this.handle.readCells({ ranges: positions.map(({ sheet, a1 }) => ({ sheetId: this.resolveSheet(sheet).nativeId, range: { kind: 'a1', a1 } })) }));
    const targets: CellTarget[] = positions.map(({ sheet, a1 }, index) => ({ sheet, a1, before: read.ranges[index].cells[0][0], version: read.version }));
    return { version: read.version, sheets: read.sheets, targets };
  }

  private hold(target: CellTarget) {
    const cell = `${this.prefix}:cell${this.nextCell++}`;
    if (this.cells.size >= 1024) this.cells.delete(this.cells.keys().next().value!);
    this.cells.set(cell, target);
    return cell;
  }

  private visible(target: CellTarget) {
    const { before, sheet, a1, version } = target;
    const value = before.value.kind === 'text' ? { kind: 'text', value: before.value.value.slice(0, 400) } : before.value;
    return { ref: this.ref(sheet, a1), cell: this.hold(target), sheet, a1, version, value, formula: before.formula?.slice(0, 400) ?? null, displayText: before.displayText.slice(0, 400), numberFormat: (this.handle.selectionFormatting(this.resolveSheet(sheet).index, a1).numberFormatPattern ?? 'General').slice(0, 400), computed: value, truncated: before.displayText.length > 400 || (before.formula?.length ?? 0) > 400 || before.value.kind === 'text' && before.value.value.length > 400 };
  }
}

export async function openXlsx(bytes: Uint8Array, options: PrototypeOptions = {}) {
  return XlsxAgentWorkbook.open(bytes, options);
}

/** Attach a live handle; pass its edit peer to synchronize a worker editor. */
export function attachXlsx(handle: WorkbookHandle, options: XlsxOptions = {}) {
  return XlsxAgentWorkbook.attach(handle, options);
}
