import { z } from 'zod';
import { dateSchema } from './tool-input';

const ref = z.string().min(1).max(128);
const text = z.string().max(16000);
const style = z.string().min(1).max(256).describe('Paragraph style ID or name defined in the document, such as Heading2.');
const twips = z.number().int().nonnegative();
const sectionProperties = {
  orientation: z.enum(['portrait', 'landscape']).optional(),
  pageWidth: z.number().int().positive().optional(),
  pageHeight: z.number().int().positive().optional(),
  margins: z.object({ top: twips.optional(), bottom: twips.optional(), left: twips.optional(), right: twips.optional(), header: twips.optional(), footer: twips.optional(), gutter: twips.optional() }).strict().optional(),
};
const mark = z.boolean().optional().describe('true sets, false clears, omitted keeps.');

export const docxEditSchema = z.discriminatedUnion('op', [
  z.object({ op: z.literal('insert_section_break'), before: ref.optional(), after: ref.optional(), kind: z.enum(['nextPage', 'continuous', 'evenPage', 'oddPage']) }).strict(),
  z.object({ op: z.literal('set_section'), section: z.number().int().positive(), ...sectionProperties }).strict(),
  z.object({ op: z.literal('set_paragraph_style'), paragraph: ref, style }).strict(),
  z.object({ op: z.literal('format_text'), match: ref.describe('Match ID from office_grep.'), bold: mark, italic: mark, underline: mark }).strict(),
  z.object({
    op: z.literal('add_comment'), match: ref.describe('Match ID from office_grep.'), author: z.string().min(1).max(200),
    text: text.min(1).describe('Plain text; each line becomes a comment paragraph.'), date: dateSchema.optional().describe('ISO timestamp or YYYY-MM-DD (midnight UTC); defaults to now.'),
  }).strict(),
  z.object({
    op: z.literal('reply_comment'), comment: ref.describe("Comment ID from office_outline story 'comments'."), author: z.string().min(1).max(200),
    text: text.min(1), date: dateSchema.optional(),
  }).strict(),
  z.object({ op: z.literal('resolve_comment'), comment: ref.describe("Comment ID from office_outline story 'comments'."), done: z.boolean().optional().describe('Defaults to true; false reopens the thread.') }).strict(),
  z.object({ op: z.literal('insert_paragraphs'), after: ref, paragraphs: z.array(z.object({ text, style: style.optional() }).strict()).min(1).max(100), runFormat: z.enum(['style', 'anchor']).optional().describe('Text formatting: style (default), or anchor first plain run; incompatible with list.'), list: z.enum(['continue']).optional().describe('continue: after a list item, the new paragraphs join its list at its level with its style and paragraph formatting.') }).strict(),
  z.object({ op: z.literal('delete_paragraph'), paragraph: ref }).strict(),
  z.object({ op: z.literal('insert_table_row'), row: ref.describe('Ref of any paragraph in the anchor row.'), position: z.enum(['before', 'after']), cells: z.array(text).min(1).max(64).describe('One text per table column.') }).strict(),
  z.object({ op: z.literal('delete_table_row'), row: ref.describe('Ref of any paragraph in the row.') }).strict(),
  z.object({ op: z.literal('merge_cells'), from: ref.describe('Ref of any paragraph in one corner cell.'), to: ref.describe('Ref of any paragraph in the opposite corner cell of the same table.') }).strict(),
]).superRefine((edit, context) => {
  if (edit.op === 'insert_section_break' && Number(edit.before !== undefined) + Number(edit.after !== undefined) !== 1) context.addIssue({ code: 'custom', message: 'Supply exactly one before or after paragraph ref.' });
  if (edit.op === 'set_section' && edit.orientation === undefined && edit.pageWidth === undefined && edit.pageHeight === undefined && !Object.keys(edit.margins ?? {}).length) context.addIssue({ code: 'custom', message: 'Supply at least one section property.' });
});
export const docxVersionSchema = z.number().int().nonnegative().describe('version from office_outline or office_grep; refresh after each edit.');
export const docxBatchSchema = z.object({ version: docxVersionSchema, edits: z.array(docxEditSchema).min(1).max(32) }).strict();
export type DocxAgentEdit = z.infer<typeof docxEditSchema>;
export type DocxAgentBatch = z.infer<typeof docxBatchSchema>;

/** What one edit did: its target ref, the resolved style, the formatted or commented range, the new comment ID, new paragraph refs, or whether the table went with its last row. */
export interface DocxEditOutcome {
  index: number;
  op: DocxAgentEdit['op'];
  ref?: string;
  style?: string;
  start?: number;
  end?: number;
  comment?: string;
  parentId?: string | null;
  done?: boolean;
  refs?: Array<string | null>;
  deletedTable?: boolean;
  rowSpan?: number;
  colSpan?: number;
}
