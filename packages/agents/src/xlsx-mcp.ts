import { z } from 'zod';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { DocumentToolError } from './types';
import { XlsxAgentWorkbook } from './xlsx';

type Register = <S extends z.ZodRawShape>(name: string, description: string, shape: S, readOnly: boolean, run: (args: z.infer<z.ZodObject<S>>) => CallToolResult | Promise<CallToolResult>) => void;

const document = z.string().min(1).max(128).describe('Document ID from office_open.');
const sheet = z.string().min(1).max(128).describe('Stable sheetId or exact sheet name from xlsx_outline.');
const range = z.string().min(1).max(40).describe('One A1 cell or rectangle, such as A1:C10, without a sheet prefix.');
const version = z.string().min(1).max(128).describe('Version from the latest outline/read. Refresh after each write.');
const offset = z.number().int().min(0).max(Number.MAX_SAFE_INTEGER).optional();
const limit = z.number().int().min(1).max(100).optional();
const target = { sheet, range };
const edits = <T extends z.ZodType>(item: T) => z.array(item).min(1).max(32);
const scalar = z.union([z.string().max(16000), z.number(), z.boolean(), z.null()]);
const matrix = <T extends z.ZodType>(item: T) => z.array(z.array(item).min(1).max(1024)).min(1).max(1024);
const color = z.string().regex(/^#[0-9a-f]{6}$/i).describe('Color in #RRGGBB format.');

function result(value: object): CallToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(value) }], structuredContent: value as Record<string, unknown> };
}

export function registerXlsxTools(register: Register, workspace: { get(id: string): unknown }, readOnly = false) {
  const workbook = (id: string) => {
    const opened = workspace.get(id);
    if (!(opened instanceof XlsxAgentWorkbook)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'This tool requires an XLSX workbook. Use office_open with an .xlsx path.');
    return opened;
  };
  register('xlsx_outline', 'Inspect sheets, dimensions, used ranges, tables, and defined names. Start here for stable sheet IDs and the current version. Follow pagination offsets.', { document, sheet: sheet.optional(), offset, limit, nameOffset: offset, tableOffset: offset }, true, args => result(workbook(args.document).outline(args)));
  register('xlsx_read_range', 'XLSX: read an A1 range and current version for direct edits with xlsx_write_range or xlsx_format. Includes empty cells, values, formulas and number formats. Follow nextOffset; long text uses office_read.', { document, ...target, offset, limit }, true, args => result(workbook(args.document).readCells(args)));
  register('xlsx_preview', 'View an A1 range as PNG after editing to verify its appearance. At most 10000 cells and 16 megapixels.', { document, ...target, scale: z.number().min(0.25).max(3).optional() }, true, args => {
    const { png, ...metadata } = workbook(args.document).previewRange(args);
    const output = result(metadata);
    output.content.push({ type: 'image', mimeType: 'image/png', data: Buffer.from(png).toString('base64') });
    return output;
  });
  const findShape = { document, query: z.string().min(1).max(1000), sheet: sheet.optional(), range: range.optional(), caseSensitive: z.boolean().optional(), offset, limit };
  if (readOnly) {
    register('xlsx_find_replace', 'Find literal displayed text, case-insensitive by default. Filter sheet/range and follow nextOffset. Replacement is disabled in read-only mode.', findShape, true, args => result(workbook(args.document).findReplace(args)));
    return;
  }
  register('xlsx_find_replace', 'Find literal displayed text, case-insensitive by default. To replace, add replacement and the found version. Replaces all matches in scope, at most 32 text cells; formulas/numbers are refused.', { ...findShape, replacement: z.string().max(16000).optional(), version: version.optional() }, false, args => result(workbook(args.document).findReplace(args)));
  register('xlsx_write_range', 'XLSX: apply 1–32 range edits immediately in memory using the latest read version, at most 1024 cells. Each edit has values OR formulas in a matrix matching its A1 range. Strings are literal; null clears. Export to save. For staged review, use office_propose_cells.', { document, version, edits: edits(z.union([z.object({ ...target, values: matrix(scalar) }).strict(), z.object({ ...target, formulas: matrix(z.string().min(1).max(16000)) }).strict()])) }, false, args => result(workbook(args.document).writeRanges(args)));
  register('xlsx_edit_grid', 'Insert/delete rows or columns atomically. at is one-based, insert is before at. Edits run in order, so later coordinates refer to the changed grid. Returns each change.', { document, version, edits: edits(z.object({ sheet, axis: z.enum(['rows', 'columns']), action: z.enum(['insert', 'delete']), at: z.number().int().min(1).max(1048576), count: z.number().int().min(1).max(10000) }).strict()) }, false, args => result(workbook(args.document).editGrid(args)));
  register('xlsx_edit_sheets', 'Add, rename, delete, or move sheets atomically. index is zero-based; move uses the final position. Sheet IDs remain stable. Later edits see earlier changes. Keep at least one sheet.', { document, version, edits: edits(z.discriminatedUnion('action', [z.object({ action: z.literal('add'), name: z.string().min(1).max(31), index: z.number().int().min(0).optional() }).strict(), z.object({ action: z.literal('rename'), sheet, name: z.string().min(1).max(31) }).strict(), z.object({ action: z.literal('delete'), sheet }).strict(), z.object({ action: z.literal('move'), sheet, index: z.number().int().min(0) }).strict()])) }, false, args => result(workbook(args.document).editSheets(args)));
  register('xlsx_format', 'Format ranges atomically: Excel number format pattern, bold/italic, fill, borders, alignment, or column width in Excel character units. At most 1024 cells. Returns applied formatting and values.', { document, version, edits: edits(z.object({ ...target, numberFormat: z.string().min(1).max(200).optional(), bold: z.boolean().optional(), italic: z.boolean().optional(), fill: color.optional(), borders: z.object({ preset: z.enum(['all', 'inner', 'horizontal', 'vertical', 'outer', 'left', 'top', 'right', 'bottom', 'none']).optional(), style: z.enum(['solid', 'dashed', 'dotted', 'double']).optional(), color: color.optional() }).strict().optional(), alignment: z.object({ horizontal: z.enum(['left', 'center', 'right']).optional(), vertical: z.enum(['top', 'middle', 'bottom']).optional(), wrap: z.enum(['overflow', 'wrap', 'clip']).optional() }).strict().optional(), columnWidth: z.number().min(1).max(255).optional() }).strict()) }, false, args => result(workbook(args.document).formatRanges(args)));
  register('xlsx_sort', 'Sort rows in an A1 range by column letters, at most 1024 cells. header preserves the first row. Values and formatting move together; formula sources keep their original A1 references. Merges are refused.', { document, version, ...target, keys: z.array(z.object({ column: z.string().regex(/^[A-Z]{1,3}$/i), order: z.enum(['asc', 'desc']) }).strict()).min(1).max(8), header: z.boolean().optional() }, false, args => result(workbook(args.document).sortRange(args)));
}
