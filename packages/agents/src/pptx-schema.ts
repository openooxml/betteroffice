import { z } from 'zod';

const id = z.string().min(1).max(1024);
const text = z.string().max(16000);
const color = z.string().regex(/^#[0-9a-fA-F]{6}$/);
const position = z.number().finite().min(-100000).max(100000);
export const pptxRectSchema = z.object({ x: position, y: position, width: z.number().positive().max(100000), height: z.number().positive().max(100000) }).strict();
export const pptxRunStyleSchema = z.object({
  bold: z.boolean().optional(), italic: z.boolean().optional(), underline: z.enum(['none', 'sng', 'dbl']).optional(),
  fontSizePt: z.number().min(1).max(400).optional(), fontFamily: z.string().min(1).max(200).optional(), color: color.optional(),
}).strict();
const paragraph = z.object({
  bullet: z.boolean().optional().describe('true adds a bullet, false removes it; omitted preserves the existing bullet.'),
  alignment: z.enum(['l', 'ctr', 'r', 'just']).optional(),
  runs: z.array(pptxRunStyleSchema.extend({ text: text.describe('Text within one paragraph; use \\u000b for a soft line break.') })).min(1).max(64),
}).strict();
const shapeTarget = { slide: id, shape: id, story: id.optional() };
const tableTarget = {
  slide: id.optional().describe('Owning slide ID; optional when shape or ref identifies the table.'),
  shape: id.optional().describe('Table shape ID from pptx_outline or shapeId from office_outline/office_grep. Omit when using a cell paragraph ref.'),
  ref: id.optional().describe('Ref or paragraph ID of any cell in the anchor row or column. Copy a fresh ref after structural edits.'),
};
const tableIndex = z.union([z.number().int().nonnegative().max(100000), id]).optional();
const tableCells = z.array(text).max(1000).optional().describe('One line of text per new cell, in column order for a row or row order for a column. Missing entries create empty cells with neighbor formatting.');
const tablePosition = z.enum(['before', 'after']).describe('Insert before or after the anchor row or column.');
export const pptxEditSchema = z.discriminatedUnion('op', [
  z.object({ op: z.literal('replace_text'), ...shapeTarget, text: text.describe('Whole-shape text. Use \\n for paragraphs and \\u000b for soft line breaks. Matching paragraph text keeps run formatting; other text uses the first run format. For partial edits use find_replace or office_propose.') }).strict(),
  z.object({ op: z.literal('set_paragraphs'), ...shapeTarget, paragraphs: z.array(paragraph).min(1).max(100) }).strict(),
  z.object({ op: z.literal('find_replace'), query: z.string().min(1).max(1000), replacement: text, caseSensitive: z.boolean().optional(), slide: id.optional() }).strict(),
  z.object({ op: z.literal('add_slide'), index: z.number().int().min(1).max(100000).optional(), layout: id.optional() }).strict(),
  z.object({ op: z.literal('set_slide_layout'), slide: id, layout: id, placeholders: z.array(z.object({ shape: id, layoutPlaceholder: id }).strict()).max(100).optional() }).strict(),
  z.object({ op: z.literal('duplicate_slide'), slide: id, index: z.number().int().min(1).max(100000).optional(), after: id.optional() }).strict(),
  z.object({ op: z.literal('delete_slide'), slide: id }).strict(),
  z.object({
    op: z.literal('move_slide'), slide: id,
    index: z.number().int().min(1).max(100000).optional().describe('One-based final position of the slide in the deck after the move.'),
    after: id.optional().describe('Slide ID to place the slide directly after. Use one of index, after or before.'),
    before: id.optional().describe('Slide ID to place the slide directly before. Use one of index, after or before.'),
  }).strict(),
  z.object({ op: z.literal('add_text_box'), slide: id, name: z.string().min(1).max(200), rect: pptxRectSchema, text, style: pptxRunStyleSchema.optional() }).strict(),
  z.object({ op: z.literal('add_shape'), slide: id, name: z.string().min(1).max(200), rect: pptxRectSchema, geometry: z.enum(['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'line', 'rightArrow', 'chevron']), fill: color.nullable().optional() }).strict(),
  z.object({ op: z.literal('add_image'), slide: id, name: z.string().min(1).max(200), rect: pptxRectSchema, contentType: z.enum(['image/png', 'image/jpeg', 'image/gif', 'image/webp']), base64: z.string().min(4).max(11200000) }).strict(),
  z.object({ op: z.literal('insert_table_row'), ...tableTarget, row: tableIndex.describe('Zero-based anchor row index with shape, or any cell paragraph ref in that row.'), position: tablePosition, cells: tableCells }).strict(),
  z.object({ op: z.literal('delete_table_row'), ...tableTarget, row: tableIndex.describe('Zero-based row index with shape, or any cell paragraph ref in the row to delete.') }).strict(),
  z.object({ op: z.literal('insert_table_column'), ...tableTarget, column: tableIndex.describe('Zero-based anchor column index with shape, or any cell paragraph ref in that column.'), position: tablePosition, cells: tableCells }).strict(),
  z.object({ op: z.literal('delete_table_column'), ...tableTarget, column: tableIndex.describe('Zero-based column index with shape, or any cell paragraph ref in the column to delete.') }).strict(),
  z.object({ op: z.literal('set_notes'), slide: id, text: text.describe('One paragraph per line; a blank line adds an empty paragraph.') }).strict(),
]);
export const pptxBatchSchema = z.object({ version: z.string().min(1).max(256), edits: z.array(pptxEditSchema).min(1).max(32) }).strict();
export type PptxAgentEdit = z.infer<typeof pptxEditSchema>;
export type PptxAgentBatch = z.infer<typeof pptxBatchSchema>;
