import { isPngExportAvailable, type RangeStylePatch, type XlsxCellValue, type XlsxEditStep } from '@betteroffice/xlsx';
import { DocumentToolError } from './types';
import { integer, plainText, unwrap } from './prototype';
import { corners, type XlsxAgentWorkbook } from './xlsx';

export type XlsxScalar = string | number | boolean | null;
export type XlsxWriteEdit = { sheet: string; range: string } & ({ values: XlsxScalar[][] } | { formulas: string[][] });
export interface XlsxGridEdit { sheet: string; axis: 'rows' | 'columns'; action: 'insert' | 'delete'; at: number; count: number }
export type XlsxSheetEdit = { action: 'add'; name: string; index?: number } | { action: 'rename'; sheet: string; name: string } | { action: 'delete'; sheet: string } | { action: 'move'; sheet: string; index: number };
export interface XlsxFormatEdit { sheet: string; range: string; numberFormat?: string; bold?: boolean; italic?: boolean; fill?: string; borders?: NonNullable<RangeStylePatch['border']>; alignment?: { horizontal?: 'left' | 'center' | 'right'; vertical?: 'top' | 'middle' | 'bottom'; wrap?: 'overflow' | 'wrap' | 'clip' }; columnWidth?: number }

function batch(edits: unknown[]) {
  if (!Array.isArray(edits) || !edits.length || edits.length > 32) throw new DocumentToolError('EDIT_LIMIT', 'Supply 1 to 32 edits in a batch.');
  if (JSON.stringify(edits).length > 64000) throw new DocumentToolError('EDIT_LIMIT', 'Batch inputs must fit in 64000 characters. Split the batch and read its new version.');
}

function sized(range: string, max = 1024) {
  const bounds = corners(range);
  const rows = bounds.end.row - bounds.start.row + 1;
  const columns = bounds.end.col - bounds.start.col + 1;
  if (rows * columns > max) throw new DocumentToolError('RANGE_LIMIT', `Use a range of at most ${max} cells. Split the operation into smaller ranges.`);
  return { ...bounds, rows, columns, total: rows * columns };
}

function scalar(value: XlsxScalar): XlsxCellValue {
  if (value === null) return { kind: 'empty' };
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new DocumentToolError('INVALID_VALUE', 'Use a finite JSON number.');
    return { kind: 'number', value };
  }
  if (typeof value === 'boolean') return { kind: 'bool', value };
  plainText(value);
  if (value.length > 16000) throw new DocumentToolError('EDIT_LIMIT', 'Cell text must fit in 16000 characters.');
  return { kind: 'text', value };
}

function failure(error: unknown, index?: number): never {
  if (error instanceof DocumentToolError) {
    const item = index ?? (typeof error.details?.stepIndex === 'number' ? error.details.stepIndex : undefined);
    const fix = error.code === 'overlapping-steps' ? 'Use disjoint ranges or split overlapping writes into separate calls.' : error.code === 'locked-target' ? 'Choose editable cells outside merged followers and array formulas.' : error.code === 'stale-version' ? 'Read the range again and use its current version.' : '';
    throw new DocumentToolError(error.code, `${error.message} ${item === undefined ? '' : `Fix edits[${item}] and retry the whole batch. `}${fix} No edits were applied.`, { ...error.details, ...(item === undefined ? {} : { item }) });
  }
  throw new DocumentToolError('ENGINE_REFUSAL', `${error instanceof Error ? error.message : String(error)}. ${index === undefined ? '' : `Fix edits[${index}]. `}No edits were applied.`, index === undefined ? undefined : { item: index });
}

function editable(book: XlsxAgentWorkbook, sheet: string) {
  const target = book.resolveSheet(sheet);
  const entry = unwrap(book.handle.readCells({ ranges: [] })).sheets.find(entry => entry.sheetId === target.nativeId);
  if (!entry?.editable) throw new DocumentToolError('LOCKED_SHEET', `Sheet ${target.name} is protected or is not a worksheet. Use an editable worksheet from xlsx_outline.`);
  return target;
}

function compactPage(page: ReturnType<XlsxAgentWorkbook['readCells']>) {
  return { ...page, cells: page.cells.map(cell => {
    const value = cell.value.kind === 'text' ? { ...cell.value, value: cell.value.value.slice(0, 120) } : cell.value;
    return { a1: cell.a1, value, formula: cell.formula?.slice(0, 120) ?? null, displayText: cell.displayText.slice(0, 120), numberFormat: cell.numberFormat.slice(0, 120), truncated: cell.truncated || cell.displayText.length > 120 || (cell.formula?.length ?? 0) > 120 || cell.value.kind === 'text' && cell.value.value.length > 120 };
  }) };
}

function calculation(book: XlsxAgentWorkbook, result: { calculation: { changed: Array<{ sheetId: string; a1: string }>; cycleCells: unknown[]; limitedCells: unknown[]; truncated: boolean } }) {
  const catalog = book.catalog();
  return { changedCells: result.calculation.changed.length, results: result.calculation.changed.slice(0, 20).map(cell => {
    const sheet = catalog.find(sheet => sheet.nativeId === cell.sheetId)!;
    return { sheet: sheet.sheetId, ...compactPage(book.readCells({ sheet: sheet.sheetId, range: cell.a1, limit: 1 })).cells[0] };
  }), cycleCells: result.calculation.cycleCells.length, limitedCells: result.calculation.limitedCells.length, truncated: result.calculation.truncated || result.calculation.changed.length > 20 };
}

function editResult(book: XlsxAgentWorkbook, steps: XlsxEditStep[], edits: Array<{ sheet: string; range: string }>, version: string) {
  let result;
  try { result = unwrap(book.handle.applyEdits({ expectVersion: version, source: 'agent', steps })); }
  catch (error) { failure(error); }
  return { version: result.version, baseVersion: result.baseVersion, applied: result.applied, items: result.receipts.map((receipt, index) => ({ index, changed: receipt.changed, changedCells: receipt.changedCells.length, ...compactPage(book.readCells({ ...edits[index], limit: Math.max(1, Math.floor(20 / edits.length)) })) })), calculation: calculation(book, result) };
}

export function writeXlsx(book: XlsxAgentWorkbook, options: { version: string; edits: XlsxWriteEdit[] }) {
  book.checkVersion(options.version);
  batch(options.edits);
  let total = 0;
  const edits = options.edits.map((edit, index) => {
    try {
      const sheet = book.resolveSheet(edit.sheet);
      const bounds = sized(edit.range);
      total += bounds.total;
      if (total > 1024) throw new DocumentToolError('EDIT_LIMIT', 'A batch may write at most 1024 cells.');
      if (('values' in edit) === ('formulas' in edit)) throw new DocumentToolError('INVALID_ARGUMENT', 'Each edit must contain values or formulas, exclusively.');
      const matrix = 'values' in edit ? edit.values : edit.formulas;
      if (!Array.isArray(matrix) || matrix.length !== bounds.rows || matrix.some(row => !Array.isArray(row) || row.length !== bounds.columns)) throw new DocumentToolError('INVALID_MATRIX', `Supply exactly ${bounds.rows} rows and ${bounds.columns} columns for ${edit.range}.`);
      const target = { sheetId: sheet.nativeId, range: { kind: 'a1' as const, a1: edit.range } };
      const step: XlsxEditStep = 'values' in edit ? { op: 'setCellValues', target, values: edit.values.map(row => row.map(scalar)) } : { op: 'setFormulas', target, formulas: edit.formulas.map(row => row.map(formula => {
        plainText(formula);
        if (!formula.length || formula.length > 16000) throw new DocumentToolError('INVALID_FORMULA', 'Supply a nonempty formula of at most 16000 characters, with or without its leading =.');
        return formula.replace(/^=/, '');
      })) };
      return { step, sheet: sheet.sheetId, range: edit.range };
    } catch (error) { failure(error, index); }
  });
  return editResult(book, edits.map(edit => edit.step), edits, options.version);
}

export function outlineXlsx(book: XlsxAgentWorkbook, options: { offset?: number; limit?: number; sheet?: string; nameOffset?: number; tableOffset?: number }) {
  const catalog = book.catalog();
  const entries = unwrap(book.handle.readCells({ ranges: [] })).sheets;
  const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
  const limit = integer(options.limit ?? 10, 1, 100, 'limit');
  const selected = options.sheet ? [book.resolveSheet(options.sheet)] : catalog;
  const nameOffset = integer(options.nameOffset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'nameOffset');
  const tableOffset = integer(options.tableOffset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'tableOffset');
  const names = unwrap(book.handle.exportStructured({ scope: [{ sheet: 0, range: 'A1' }], includeHiddenSheets: true, includeHiddenNames: true, maxCells: 1, maxBytes: 64000 })).content;
  const items = [];
  let size = 0;
  for (const sheet of selected.slice(offset, offset + limit)) {
    const read = unwrap(book.handle.exportStructured({ scope: [{ sheet: sheet.index }], includeHiddenSheets: true, includeHiddenRows: true, includeHiddenColumns: true, includeDefinedNames: false, maxCells: 1, maxBytes: 64000 })).content.sheets[0];
    const bounds = read.usedRange ? corners(read.usedRange) : null;
    const info = book.handle.sheetInfoFor(sheet.index);
    const item = { sheetId: sheet.sheetId, name: sheet.name, index: sheet.index, kind: read.kind, editable: entries.find(entry => entry.sheetId === sheet.nativeId)?.editable ?? false, visibility: read.visibility, usedRange: read.usedRange, dimensions: { rows: bounds ? bounds.end.row + 1 : 0, columns: bounds ? bounds.end.col + 1 : 0, width: info.contentWidth, height: info.contentHeight }, tables: read.tables.slice(tableOffset, tableOffset + 10).map(table => ({ id: table.id, name: table.name.slice(0, 120), range: table.anchor.kind === 'range' ? table.anchor.a1 : null, headerRows: table.headerRows, totalsRows: table.totalsRows, columns: table.columns.slice(0, 10).map(column => column.slice(0, 100)), columnCount: table.columns.length })), tableCount: read.tables.length, metadataTruncated: read.truncated && read.cells.length === 0, nextTableOffset: tableOffset + 10 < read.tables.length ? tableOffset + 10 : null };
    const cost = JSON.stringify(item).length;
    if (items.length && size + cost > 12000) break;
    size += cost;
    items.push(item);
  }
  return { version: book.handle.version(), items, total: selected.length, nextOffset: offset + items.length < selected.length ? offset + items.length : null, definedNames: names.definedNames.slice(nameOffset, nameOffset + 10).map(name => ({ id: name.id, name: name.name.slice(0, 120), formula: name.formula.slice(0, 400), truncated: name.formula.length > 400, hidden: name.hidden, sheet: name.localSheet ? catalog.find(sheet => sheet.index === name.localSheet!.index)?.sheetId : null })), definedNameCount: names.definedNames.length, nextNameOffset: nameOffset + 10 < names.definedNames.length ? nameOffset + 10 : null, metadataTruncated: names.truncated && names.sheets.length === 0 };
}

export function editXlsxGrid(book: XlsxAgentWorkbook, options: { version: string; edits: XlsxGridEdit[] }) {
  book.checkVersion(options.version);
  batch(options.edits);
  const ops = options.edits.map((edit, index) => {
    try {
      const sheet = editable(book, edit.sheet);
      const max = edit.axis === 'rows' ? 1048576 : 16384;
      integer(edit.at, 1, max, 'at');
      integer(edit.count, 1, Math.min(10000, max - edit.at + 1), 'count');
      if (!['rows', 'columns'].includes(edit.axis) || !['insert', 'delete'].includes(edit.action)) throw new DocumentToolError('INVALID_ARGUMENT', 'Use axis rows or columns and action insert or delete.');
      return { type: `${edit.action}${edit.axis === 'rows' ? 'Rows' : 'Cols'}`, sheet: sheet.index, at: edit.at - 1, count: edit.count };
    } catch (error) { failure(error, index); }
  });
  let result;
  try { book.checkVersion(options.version); result = book.handle.applyOps(ops); } catch (error) { failure(error); }
  if (result.applied) book.invalidateCells();
  return { baseVersion: options.version, version: book.handle.version(), applied: result.applied, items: options.edits.map((edit, index) => ({ index, ...edit, changed: result.applied })), calculation: { changedCells: result.changed?.length ?? 0, limitedCells: result.limitedCells?.length ?? 0 } };
}

function sheetName(name: string, names: string[], except?: string) {
  plainText(name);
  if (!name.length || name.length > 31 || /[\\/\[\]:*?]/.test(name) || name.startsWith("'") || name.endsWith("'")) throw new DocumentToolError('INVALID_SHEET_NAME', 'Use 1 to 31 characters without \\, /, [, ], :, *, ?, or a leading/trailing apostrophe.');
  if (names.some(existing => existing !== except && existing.toLowerCase() === name.toLowerCase())) throw new DocumentToolError('DUPLICATE_SHEET_NAME', `Name ${JSON.stringify(name)} already exists. Choose a different name. Valid names: ${names.slice(0, 30).join(', ')}.`);
}

export function editXlsxSheets(book: XlsxAgentWorkbook, options: { version: string; edits: XlsxSheetEdit[] }) {
  book.checkVersion(options.version);
  batch(options.edits);
  const working: Array<{ sheetId: string; name: string }> = book.catalog().map(({ sheetId, name }) => ({ sheetId, name }));
  const items: Array<Record<string, unknown>> = [];
  const ops = options.edits.map((edit, index) => {
    try {
      if (edit.action === 'add') {
        sheetName(edit.name, working.map(sheet => sheet.name));
        const at = integer(edit.index ?? working.length, 0, working.length, 'index');
        working.splice(at, 0, { sheetId: `added:${index}`, name: edit.name });
        items.push({ index, action: edit.action, sheet: `added:${index}`, name: edit.name, position: at });
        return { type: 'addSheet', index: at, name: edit.name };
      }
      const byId = working.findIndex(sheet => sheet.sheetId === edit.sheet);
      const at = byId >= 0 ? byId : working.findIndex(sheet => sheet.name === edit.sheet);
      if (at < 0) throw new DocumentToolError('UNKNOWN_SHEET', `Use a current sheetId or name. Valid sheets: ${working.slice(0, 30).map(sheet => `${sheet.sheetId} (${sheet.name})`).join(', ')}.`);
      const target = working[at];
      if (edit.action === 'rename') {
        sheetName(edit.name, working.map(sheet => sheet.name), target.name);
        items.push({ index, action: edit.action, sheet: target.sheetId, before: target.name, name: edit.name, changed: target.name !== edit.name });
        target.name = edit.name;
        return { type: 'renameSheet', sheet: at, name: edit.name };
      }
      if (edit.action === 'delete') {
        if (working.length === 1) throw new DocumentToolError('LAST_SHEET', 'Keep at least one sheet. Add a new sheet before deleting the last sheet.');
        working.splice(at, 1);
        items.push({ index, action: edit.action, sheet: target.sheetId, name: target.name });
        return { type: 'removeSheet', index: at };
      }
      if (edit.action !== 'move') throw new DocumentToolError('INVALID_ARGUMENT', 'Use action add, rename, delete, or move.');
      const to = integer(edit.index, 0, working.length - 1, 'index');
      working.splice(at, 1);
      working.splice(to, 0, target);
      items.push({ index, action: edit.action, sheet: target.sheetId, from: at, position: to, changed: at !== to });
      return { type: 'moveSheet', from: at, to };
    } catch (error) { failure(error, index); }
  });
  let result;
  try { book.checkVersion(options.version); result = book.handle.applyOps(ops); } catch (error) { failure(error); }
  const assigned = book.adoptSheets(working);
  const catalog = book.catalog();
  if (result.applied) book.invalidateCells();
  return { baseVersion: options.version, version: book.handle.version(), applied: result.applied, items: items.map(item => ({ changed: true, ...item, ...(typeof item.sheet === 'string' && item.sheet.startsWith('added:') ? { sheet: assigned.get(item.sheet) ?? null, deletedInBatch: !assigned.has(item.sheet) } : {}) })), sheets: catalog.slice(0, 100).map(({ sheetId, name, index }) => ({ sheetId, name, index })), sheetCount: catalog.length };
}

export function applyXlsxFormatting(book: XlsxAgentWorkbook, options: { version: string; edits: XlsxFormatEdit[] }) {
  book.checkVersion(options.version);
  batch(options.edits);
  let total = 0;
  const targets = options.edits.map((edit, index) => {
    try {
      const sheet = editable(book, edit.sheet);
      const bounds = sized(edit.range);
      total += bounds.total;
      if (total > 1024) throw new DocumentToolError('EDIT_LIMIT', 'A formatting batch may cover at most 1024 cells.');
      const patch: RangeStylePatch = { ...(edit.bold === undefined ? {} : { bold: edit.bold }), ...(edit.italic === undefined ? {} : { italic: edit.italic }), ...(edit.fill === undefined ? {} : { fillColor: edit.fill }), ...(edit.borders === undefined ? {} : { border: edit.borders }), ...(edit.alignment?.horizontal === undefined ? {} : { horizontalAlignment: edit.alignment.horizontal }), ...(edit.alignment?.vertical === undefined ? {} : { verticalAlignment: edit.alignment.vertical }), ...(edit.alignment?.wrap === undefined ? {} : { textWrapping: edit.alignment.wrap }) };
      for (const color of [edit.fill, edit.borders?.color]) if (color !== undefined && !/^#[0-9a-f]{6}$/i.test(color)) throw new DocumentToolError('INVALID_COLOR', 'Use a color in #RRGGBB format.');
      const target = { sheetId: sheet.nativeId, range: { kind: 'a1' as const, a1: edit.range } };
      const ops: unknown[] = [];
      const guard: XlsxEditStep = { op: 'patchStyle', target, patch };
      unwrap(book.handle.validateEdits({ expectVersion: options.version, source: 'agent', steps: [guard] }));
      const rawRange = { start: bounds.start, end: bounds.end };
      if (Object.keys(patch).length) ops.push({ type: 'patchRangeStyle', sheet: sheet.index, range: rawRange, patch });
      if (edit.numberFormat !== undefined) {
        plainText(edit.numberFormat);
        if (!edit.numberFormat.length || edit.numberFormat.length > 200) throw new DocumentToolError('INVALID_FORMAT', 'Use a number format pattern of 1 to 200 characters, such as 0.00 or yyyy-mm-dd.');
        const format = { type: 'custom' as const, pattern: edit.numberFormat };
        unwrap(book.handle.validateEdits({ expectVersion: options.version, source: 'agent', steps: [{ op: 'setNumberFormat', target, format }] }));
        ops.push({ type: 'setRangeNumberFormat', sheet: sheet.index, range: rawRange, format });
      }
      if (edit.columnWidth !== undefined) {
        if (!Number.isFinite(edit.columnWidth) || edit.columnWidth < 1 || edit.columnWidth > 255) throw new DocumentToolError('INVALID_WIDTH', 'Use columnWidth from 1 to 255 Excel character units.');
        for (let col = bounds.start.col; col <= bounds.end.col; col++) ops.push({ type: 'setColWidth', sheet: sheet.index, col, width: edit.columnWidth });
      }
      if (!ops.length) throw new DocumentToolError('EMPTY_FORMAT', 'Supply numberFormat, bold, italic, fill, borders, alignment, or columnWidth.');
      return { sheet, range: edit.range, ops, before: book.handle.captureFormat(sheet.index, edit.range), widthsBefore: Array.from({ length: bounds.columns }, (_, col) => book.handle.cellRect(sheet.index, bounds.start.row, bounds.start.col + col).w), bounds };
    } catch (error) { failure(error, index); }
  });
  let result;
  try { book.checkVersion(options.version); result = book.handle.applyOps(targets.flatMap(target => target.ops)); } catch (error) { failure(error); }
  return { baseVersion: options.version, version: book.handle.version(), applied: result.applied, items: targets.map((target, index) => ({ index, changed: JSON.stringify(target.before) !== JSON.stringify(book.handle.captureFormat(target.sheet.index, target.range)) || target.widthsBefore.some((width, col) => width !== book.handle.cellRect(target.sheet.index, target.bounds.start.row, target.bounds.start.col + col).w), formatting: book.handle.selectionFormatting(target.sheet.index, target.range), ...(options.edits[index].columnWidth === undefined ? {} : { columnWidth: options.edits[index].columnWidth }), ...compactPage(book.readCells({ sheet: target.sheet.sheetId, range: target.range, limit: Math.max(1, Math.floor(20 / targets.length)) })) })) };
}

export function findXlsx(book: XlsxAgentWorkbook, options: { query: string; sheet?: string; range?: string; caseSensitive?: boolean; offset?: number; limit?: number; replacement?: string; version?: string }) {
  book.catalog();
  plainText(options.query);
  if (!options.query.length || options.query.length > 1000) throw new DocumentToolError('INVALID_QUERY', 'Use a literal query of 1 to 1000 characters.');
  if (options.range && !options.sheet) throw new DocumentToolError('INVALID_ARGUMENT', 'Supply sheet when filtering by range.');
  const sheet = options.sheet ? book.resolveSheet(options.sheet) : undefined;
  const bounds = options.range ? corners(options.range) : undefined;
  const hits = book.handle.searchText(options.query, { caseSensitive: options.caseSensitive ?? false, limit: 10000 });
  const matches = hits.filter(hit => (!sheet || hit.sheet === sheet.index) && (!bounds || hit.row >= bounds.start.row && hit.row <= bounds.end.row && hit.col >= bounds.start.col && hit.col <= bounds.end.col));
  const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
  const limit = integer(options.limit ?? 20, 1, 100, 'limit');
  const catalog = book.catalog();
  const page = [];
  let size = 0;
  for (const hit of matches.slice(offset, offset + limit)) {
    const item = { sheet: catalog[hit.sheet].sheetId, a1: hit.a1, text: hit.text.slice(0, 400), truncated: hit.text.length > 400 };
    const cost = JSON.stringify(item).length;
    if (page.length && size + cost > 16000) break;
    size += cost;
    page.push(item);
  }
  if (options.replacement === undefined) return { version: book.handle.version(), matches: page, total: matches.length, nextOffset: offset + page.length < matches.length ? offset + page.length : null, truncated: hits.length === 10000 };
  if (!options.version) throw new DocumentToolError('VERSION_REQUIRED', 'Find first, then supply its version to replace.');
  book.checkVersion(options.version);
  plainText(options.replacement);
  if (options.replacement.length > 16000) throw new DocumentToolError('EDIT_LIMIT', 'replacement must fit in 16000 characters.');
  if (hits.length === 10000 || matches.length > 32) throw new DocumentToolError('REPLACE_LIMIT', 'Narrow sheet/range/query to at most 32 matching cells, then replace. Replacement always covers every match in that scope.');
  const expression = new RegExp(options.query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), options.caseSensitive ? 'g' : 'gi');
  const edits: XlsxWriteEdit[] = matches.map(hit => {
    const target = catalog[hit.sheet];
    const cell = unwrap(book.handle.readCells({ ranges: [{ sheetId: target.nativeId, range: { kind: 'a1', a1: hit.a1 } }] })).ranges[0].cells[0][0];
    if (cell.formula || cell.value.kind !== 'text') throw new DocumentToolError('REPLACE_VALUE', `Cell ${target.name}!${hit.a1} is a formula or a non-text value. Narrow the scope to text cells or use xlsx_write_range to replace it explicitly. No edits were applied.`);
    return { sheet: target.sheetId, range: hit.a1, values: [[cell.value.value.replace(expression, () => options.replacement!)]] };
  });
  if (!edits.length) return { baseVersion: options.version, version: book.handle.version(), applied: false, items: [], replacedCells: 0 };
  return { ...writeXlsx(book, { version: options.version, edits }), replacedCells: edits.length };
}

function compare(a: XlsxCellValue, b: XlsxCellValue) {
  if (a.kind === 'empty' || b.kind === 'empty') return a.kind === b.kind ? 0 : a.kind === 'empty' ? 1 : -1;
  const rank = { number: 0, text: 1, bool: 2, error: 3 };
  if (a.kind !== b.kind) return rank[a.kind] - rank[b.kind];
  if (a.kind === 'number' && b.kind === 'number') return a.value - b.value;
  return String(a.value).toLowerCase().localeCompare(String(b.value).toLowerCase(), 'en');
}

export function sortXlsx(book: XlsxAgentWorkbook, options: { version: string; sheet: string; range: string; keys: Array<{ column: string; order: 'asc' | 'desc' }>; header?: boolean }) {
  book.checkVersion(options.version);
  const sheet = editable(book, options.sheet);
  const bounds = sized(options.range);
  if (!Array.isArray(options.keys) || !options.keys.length || options.keys.length > 8) throw new DocumentToolError('INVALID_SORT', 'Supply 1 to 8 keys with column letters and order asc or desc.');
  const keys = options.keys.map(key => {
    if (!/^[A-Z]{1,3}$/i.test(key.column) || !['asc', 'desc'].includes(key.order)) throw new DocumentToolError('INVALID_SORT', 'Use column letters such as B and order asc or desc.');
    const col = corners(`${key.column}1`).start.col;
    if (col < bounds.start.col || col > bounds.end.col) throw new DocumentToolError('INVALID_SORT', `Sort column ${key.column} must be inside ${options.range}.`);
    return { col: col - bounds.start.col, order: key.order };
  });
  if (book.handle.mergedRanges(sheet.index, options.range).length) throw new DocumentToolError('MERGED_SORT', 'Sort a range without merged cells.');
  const target = { sheetId: sheet.nativeId, range: { kind: 'a1' as const, a1: options.range } };
  const rows = unwrap(book.handle.readCells({ ranges: [target] })).ranges[0].cells;
  const inputs = book.handle.rangeCells(sheet.index, options.range).map(row => row.map(cell => cell.input));
  unwrap(book.handle.validateEdits({ expectVersion: book.handle.version(), source: 'agent', steps: [{ op: 'setCellInputs', target, inputs }] }));
  const start = options.header ? 1 : 0;
  const order = Array.from({ length: rows.length - start }, (_, index) => index + start).sort((a, b) => {
    for (const key of keys) {
      const av = rows[a][key.col].value;
      const bv = rows[b][key.col].value;
      const cmp = compare(av, bv);
      if (cmp) return av.kind === 'empty' || bv.kind === 'empty' || key.order === 'asc' ? cmp : -cmp;
    }
    return a - b;
  });
  const captured = book.handle.captureFormat(sheet.index, options.range);
  const ops: unknown[] = [];
  for (let dest = start; dest < rows.length; dest++) {
    const source = order[dest - start];
    for (let col = 0; col < bounds.columns; col++) {
      const cell = rows[source][col];
      ops.push({ type: 'setCell', sheet: sheet.index, at: { row: bounds.start.row + dest, col: bounds.start.col + col }, cell: { value: cell.value, formula: cell.formula } });
    }
  }
  const reorderFormats = <T>(formats: readonly T[]): T[] => Array.from({ length: rows.length }, (_, dest) => {
    const source = dest < start ? dest : order[dest - start];
    return formats.slice(source * bounds.columns, (source + 1) * bounds.columns);
  }).flat();
  if (ops.length) ops.push({ type: 'applyRangeFormat', sheet: sheet.index, range: { start: bounds.start, end: bounds.end }, format: {
    ...captured,
    formats: reorderFormats(captured.formats),
    ...(captured.sourceStyles && { sourceStyles: reorderFormats(captured.sourceStyles) }),
    ...(captured.sourceFormats && { sourceFormats: reorderFormats(captured.sourceFormats) }),
  } });
  let result;
  try { book.checkVersion(options.version); result = book.handle.applyOps(ops); } catch (error) { failure(error); }
  return { baseVersion: options.version, version: book.handle.version(), applied: result.applied, sheet: sheet.sheetId, range: options.range, rows: order.slice(0, 100).map((source, index) => ({ from: bounds.start.row + source + 1, to: bounds.start.row + start + index + 1 })), rowCount: order.length, truncated: order.length > 100, formulaPolicy: 'Formula source moves unchanged with its row; references keep their original A1 addresses.', results: book.readCells({ sheet: sheet.sheetId, range: options.range, limit: 20 }) };
}

export function previewXlsx(book: XlsxAgentWorkbook, options: { sheet: string; range: string; scale?: number }) {
  const sheet = book.resolveSheet(options.sheet);
  sized(options.range, 10000);
  const scale = options.scale ?? 1;
  if (!Number.isFinite(scale) || scale < 0.25 || scale > 3) throw new DocumentToolError('INVALID_SCALE', 'Use scale from 0.25 to 3.');
  if (!isPngExportAvailable()) throw new DocumentToolError('PNG_UNAVAILABLE', 'Build the XLSX Wasm engine with its raster feature to enable PNG previews.');
  const bounds = corners(options.range);
  const first = book.handle.cellRect(sheet.index, bounds.start.row, bounds.start.col);
  const last = book.handle.cellRect(sheet.index, bounds.end.row, bounds.end.col);
  const width = Math.ceil((last.x + last.w - first.x) * scale);
  const height = Math.ceil((last.y + last.h - first.y) * scale);
  if (width <= 0 || height <= 0 || width * height > 16000000) throw new DocumentToolError('PREVIEW_LIMIT', 'Preview exceeds 16 megapixels. Select a smaller range or lower scale.');
  const active = book.handle.sheetInfo().activeSheet;
  let png;
  try {
    book.handle.setActiveSheet(sheet.index);
    png = book.handle.renderRangePng({ range: options.range, scale });
  } finally { book.handle.setActiveSheet(active); }
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  return { png, version: book.handle.version(), sheet: sheet.sheetId, range: options.range, width: view.getUint32(16), height: view.getUint32(20), warnings: [] as string[] };
}
