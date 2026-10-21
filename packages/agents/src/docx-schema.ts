import { z } from 'zod';

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
  z.object({ op: z.literal('insert_paragraphs'), after: ref, paragraphs: z.array(z.object({ text, style: style.optional() }).strict()).min(1).max(100) }).strict(),
  z.object({ op: z.literal('delete_paragraph'), paragraph: ref }).strict(),
  z.object({ op: z.literal('insert_table_row'), row: ref.describe('Ref of any paragraph in the anchor row.'), position: z.enum(['before', 'after']), cells: z.array(text).min(1).max(64).describe('One text per table column.') }).strict(),
  z.object({ op: z.literal('delete_table_row'), row: ref.describe('Ref of any paragraph in the row.') }).strict(),
]).superRefine((edit, context) => {
  if (edit.op === 'insert_section_break' && Number(edit.before !== undefined) + Number(edit.after !== undefined) !== 1) context.addIssue({ code: 'custom', message: 'Supply exactly one before or after paragraph ref.' });
  if (edit.op === 'set_section' && edit.orientation === undefined && edit.pageWidth === undefined && edit.pageHeight === undefined && !Object.keys(edit.margins ?? {}).length) context.addIssue({ code: 'custom', message: 'Supply at least one section property.' });
});
export const docxVersionSchema = z.number().int().nonnegative().describe('version from office_outline or office_grep; refresh after each edit.');
export const docxBatchSchema = z.object({ version: docxVersionSchema, edits: z.array(docxEditSchema).min(1).max(32) }).strict();
export type DocxAgentEdit = z.infer<typeof docxEditSchema>;
export type DocxAgentBatch = z.infer<typeof docxBatchSchema>;

/** What one edit did: its target ref, the resolved style, the formatted range, new paragraph refs, or whether the table went with its last row. */
export interface DocxEditOutcome {
  index: number;
  op: DocxAgentEdit['op'];
  ref?: string;
  style?: string;
  start?: number;
  end?: number;
  refs?: Array<string | null>;
  deletedTable?: boolean;
}
