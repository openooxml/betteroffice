import { initWasm, openWorkbook, type WorkbookHandle, type XlsxCellRead, type XlsxEditStep } from '@betteroffice/xlsx';
import { DocumentToolError, type GrepOptions, type TextEdit } from './types';
import { initialize, integer, plainText, PrototypeDocument, textWindow, unwrap, type PrototypeOptions } from './prototype';

export interface CellEdit { cell: string; input: string }
export interface CellChange { ref: string; sheet: string; a1: string; before: XlsxCellRead; input: string }
interface CellTarget { sheet: string; a1: string; before: XlsxCellRead; version: string }

function address(row: number, col: number) {
  let letters = '';
  for (let n = col + 1; n; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + (n - 1) % 26) + letters;
  return `${letters}${row + 1}`;
}

function corners(range: string) {
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

  private constructor(private readonly handle: WorkbookHandle, readonly name = 'workbook.xlsx') { super(); }

  static async open(bytes: Uint8Array, options: PrototypeOptions = {}) {
    await initialize('xlsx', initWasm, options.wasm);
    return new XlsxAgentWorkbook(openWorkbook(bytes), options.name);
  }

  overview() {
    this.alive();
    const read = unwrap(this.handle.readCells({ ranges: [] }));
    return { name: this.name, format: this.format, version: read.version, sheets: read.sheets.length, capabilities: { grep: 'literal, case-sensitive', proposals: 'cell inputs and formulas', render: false, export: true, prototype: true }, workflow: 'outline lists sheet IDs; grep finds displayed values; cells reads A1 ranges and returns cell handles; proposeCells with cell and input (prefix = for a formula); review; verify; export.' };
  }

  list(options: { offset?: number; limit?: number; story?: string; headingsOnly?: boolean } = {}) {
    this.alive();
    if (options.headingsOnly) throw new DocumentToolError('UNSUPPORTED', 'XLSX outline lists sheets; omit headingsOnly.');
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 30, 1, 100, 'limit');
    const read = unwrap(this.handle.readCells({ ranges: [] }));
    if (options.story && !read.sheets.some(s => s.sheetId === options.story)) throw new DocumentToolError('UNKNOWN_STORY', 'Use a sheetId from outline.');
    const sheets = read.sheets.filter(s => !options.story || s.sheetId === options.story);
    return { version: read.version, items: sheets.slice(offset, offset + limit), nextOffset: offset + limit < sheets.length ? offset + limit : null, total: sheets.length };
  }

  readCells(options: { sheet: string; range: string; offset?: number; limit?: number }) {
    this.alive();
    const { start, end } = corners(options.range);
    const width = end.col - start.col + 1;
    const total = width * (end.row - start.row + 1);
    const offset = integer(options.offset ?? 0, 0, total, 'offset');
    const limit = integer(options.limit ?? 50, 1, 100, 'limit');
    const cells = [];
    let size = 0;
    for (let index = offset; index < Math.min(total, offset + limit); index++) {
      const a1 = address(start.row + Math.floor(index / width), start.col + index % width);
      const target = this.cellAt(options.sheet, a1);
      const visible = this.visible(target);
      const cost = JSON.stringify(visible).length;
      if (cells.length && size + cost > 16000) break;
      size += cost;
      cells.push(visible);
    }
    return { version: this.currentVersion(), sheet: options.sheet, range: options.range, cells, total, nextOffset: offset + cells.length < total ? offset + cells.length : null };
  }

  grep(options: GrepOptions) {
    this.checkQuery(options);
    if (options.caseSensitive === false) throw new DocumentToolError('UNSUPPORTED', 'XLSX search is case-sensitive. Omit caseSensitive or set it to true.');
    const found = unwrap(this.handle.findText({ text: options.query, sheetIds: options.story ? [options.story] : undefined, limit: 10000 }));
    const page = this.searchPage(options, found.matches.map(hit => ({ ref: this.ref(hit.cell.sheetId, hit.cell.a1), sheet: hit.cell.sheetId, a1: hit.cell.a1, text: hit.text.slice(0, 400), length: hit.text.length })), found.truncated);
    return { ...page, matches: page.matches.map(hit => ({ ...hit, cell: this.hold(this.cellAt(hit.sheet, hit.a1)) })) };
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
      steps.push({ op: 'setCellInputs', target: { sheetId: target.sheet, range: { kind: 'a1', a1: target.a1 } }, inputs: [[edit.input]], expect: { cells: [[{ value: target.before.value, formula: target.before.formula, displayText: target.before.displayText }]] } });
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
  protected dispose() { this.cells.clear(); this.refs.clear(); this.handle.dispose(); }

  private ref(sheet: string, a1: string) {
    const ref = `${sheet}!${a1}`;
    if (this.refs.size >= 10000 && !this.refs.has(ref)) this.refs.delete(this.refs.keys().next().value!);
    this.refs.set(ref, { sheet, a1 });
    return ref;
  }

  private cellAt(sheet: string, a1: string): CellTarget {
    const read = unwrap(this.handle.readCells({ ranges: [{ sheetId: sheet, range: { kind: 'a1', a1 } }] }));
    return { sheet, a1, before: read.ranges[0].cells[0][0], version: read.version };
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
    return { ref: this.ref(sheet, a1), cell: this.hold(target), sheet, a1, version, value, formula: before.formula?.slice(0, 400) ?? null, displayText: before.displayText.slice(0, 400), truncated: before.displayText.length > 400 || (before.formula?.length ?? 0) > 400 || before.value.kind === 'text' && before.value.value.length > 400 };
  }
}

export async function openXlsx(bytes: Uint8Array, options: PrototypeOptions = {}) {
  return XlsxAgentWorkbook.open(bytes, options);
}
