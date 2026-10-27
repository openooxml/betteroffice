import { z } from 'zod';

const id = z.string().min(1).max(1024);
export const pptxSlideIdSchema = id.describe('Copy a slide ID from pptx_outline. Opened IDs use slide:<0-based index>:<source ID>; the middle number is the original 0-based index. Numeric addresses are ambiguous and refused.');
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
const shapeTarget = { slide: pptxSlideIdSchema.optional().describe('Optional: inferred from shape. If supplied, must identify the same slide as shape.'), shape: id.describe('Shape ID from pptx_outline or pptx_read_slide; identifies its slide.'), story: id.optional() };
const tableTarget = {
  slide: pptxSlideIdSchema.optional().describe('Optional: inferred from shape or ref. If supplied, must identify the same slide as the table.'),
  shape: id.optional().describe('Table shape ID from pptx_outline or shapeId from office_outline/office_grep. Omit when using a cell paragraph ref.'),
  ref: id.optional().describe('Ref or paragraph ID of any cell in the anchor row or column. Copy a fresh ref after structural edits.'),
};
const tableIndex = z.union([z.number().int().nonnegative().max(100000), id]).optional();
const tableCells = z.array(text).max(1000).optional().describe('One line of text per new cell, in column order for a row or row order for a column. Missing entries create empty cells with neighbor formatting.');
const tablePosition = z.enum(['before', 'after']).describe('Insert before or after the anchor row or column.');
export const pptxEditSchema = z.discriminatedUnion('op', [
  z.object({ op: z.literal('replace_text'), ...shapeTarget, text: text.describe('Whole-shape text. Use \\n for paragraphs and \\u000b for soft line breaks. Matching paragraph text keeps run formatting; other text uses the first run format. For partial edits use find_replace or office_propose.') }).strict(),
  z.object({ op: z.literal('set_paragraphs'), ...shapeTarget, paragraphs: z.array(paragraph).min(1).max(100) }).strict(),
  z.object({ op: z.literal('find_replace'), query: z.string().min(1).max(1000), replacement: text, caseSensitive: z.boolean().optional(), slide: pptxSlideIdSchema.optional(), shape: id.optional().describe('Restrict replacement to this exact shape ID. Slide is inferred if omitted; an explicit slide must agree.') }).strict(),
  z.object({ op: z.literal('add_slide'), index: z.number().int().min(1).max(100000).optional(), layout: id.optional() }).strict(),
  z.object({ op: z.literal('set_slide_layout'), slide: pptxSlideIdSchema, layout: id, placeholders: z.array(z.object({ shape: id.describe('Shape ID on the slide from pptx_outline or pptx_read_slide.'), layoutPlaceholder: id.describe('Placeholder ID from the target layout in pptx_outline.') }).strict()).max(100).optional() }).strict(),
  z.object({ op: z.literal('duplicate_slide'), slide: pptxSlideIdSchema, index: z.number().int().min(1).max(100000).optional(), after: pptxSlideIdSchema.optional() }).strict(),
  z.object({ op: z.literal('delete_slide'), slide: pptxSlideIdSchema }).strict(),
  z.object({
    op: z.literal('move_slide'), slide: pptxSlideIdSchema,
    index: z.number().int().min(1).max(100000).optional().describe('One-based final position of the slide in the deck after the move.'),
    after: pptxSlideIdSchema.optional().describe('Slide ID to place the slide directly after. Use one of index, after or before.'),
    before: pptxSlideIdSchema.optional().describe('Slide ID to place the slide directly before. Use one of index, after or before.'),
  }).strict(),
  z.object({ op: z.literal('add_text_box'), slide: pptxSlideIdSchema, name: z.string().min(1).max(200), rect: pptxRectSchema, text, style: pptxRunStyleSchema.optional() }).strict(),
  z.object({ op: z.literal('add_shape'), slide: pptxSlideIdSchema, name: z.string().min(1).max(200), rect: pptxRectSchema, geometry: z.enum(['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'line', 'rightArrow', 'chevron']), fill: color.nullable().optional() }).strict(),
  z.object({ op: z.literal('add_image'), slide: pptxSlideIdSchema, name: z.string().min(1).max(200), rect: pptxRectSchema, contentType: z.enum(['image/png', 'image/jpeg', 'image/gif', 'image/webp']), base64: z.string().min(4).max(11200000) }).strict(),
  z.object({ op: z.literal('insert_table_row'), ...tableTarget, row: tableIndex.describe('Zero-based anchor row index with shape, or any cell paragraph ref in that row.'), position: tablePosition, cells: tableCells }).strict(),
  z.object({ op: z.literal('delete_table_row'), ...tableTarget, row: tableIndex.describe('Zero-based row index with shape, or any cell paragraph ref in the row to delete.') }).strict(),
  z.object({ op: z.literal('insert_table_column'), ...tableTarget, column: tableIndex.describe('Zero-based anchor column index with shape, or any cell paragraph ref in that column.'), position: tablePosition, cells: tableCells, width: z.enum(['split', 'copy']).optional().describe('Column width mode. split (default) halves the anchor width and keeps the table width; copy preserves existing widths, copies the anchor width and refits the table frame to the grid.') }).strict(),
  z.object({ op: z.literal('delete_table_column'), ...tableTarget, column: tableIndex.describe('Zero-based column index with shape, or any cell paragraph ref in the column to delete.') }).strict(),
  z.object({
    op: z.literal('add_comment'), slide: pptxSlideIdSchema,
    author: z.string().min(1).max(200).describe('Comment author name.'), initials: z.string().max(32).optional().describe('Author initials; defaults to empty.'),
    text: text.min(1).describe('Plain comment text.'), date: z.string().datetime().optional().describe('ISO date; defaults to now.'),
    position: z.object({ x: position, y: position }).strict().optional().describe('Anchor position in slide points; defaults to the origin.'),
  }).strict().describe('Add a comment to a slide in the deck comment flavor.'),
  z.object({
    op: z.literal('reply_comment'), comment: id.describe("Comment ID from office_outline story 'comments'; a reply targets its thread root."),
    author: z.string().min(1).max(200).describe('Reply author name.'), initials: z.string().max(32).optional().describe('Author initials; defaults to empty.'),
    text: text.min(1).describe('Plain reply text.'), date: z.string().datetime().optional().describe('ISO date; defaults to now.'),
  }).strict().describe('Reply to a legacy or modern slide comment.'),
  z.object({
    op: z.literal('resolve_comment'), comment: id.describe("Comment ID from office_outline story 'comments'; targets the thread root."),
    resolved: z.boolean().optional().describe('Defaults to true; false reopens the thread. Modern comments only.'),
    done: z.boolean().optional().describe('Alias of resolved, as in docx_edit.'),
  }).strict().describe('Resolve or reopen a modern comment; legacy comments cannot represent status.'),
  z.object({ op: z.literal('set_notes'), slide: pptxSlideIdSchema, text: text.describe('One paragraph per line; a blank line adds an empty paragraph.') }).strict(),
]);
export const pptxBatchSchema = z.object({ version: z.string().min(1).max(256), edits: z.array(pptxEditSchema).min(1).max(32) }).strict();
export type PptxAgentEdit = z.infer<typeof pptxEditSchema>;
export type PptxAgentBatch = z.infer<typeof pptxBatchSchema>;
