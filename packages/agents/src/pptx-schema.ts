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
  alignment: z.enum(['l', 'ctr', 'r', 'just']).optional(),
  runs: z.array(pptxRunStyleSchema.extend({ text })).min(1).max(64),
}).strict();
const shapeTarget = { slide: id, shape: id, story: id.optional() };
export const pptxEditSchema = z.discriminatedUnion('op', [
  z.object({ op: z.literal('replace_text'), ...shapeTarget, text }).strict(),
  z.object({ op: z.literal('set_paragraphs'), ...shapeTarget, paragraphs: z.array(paragraph).min(1).max(100) }).strict(),
  z.object({ op: z.literal('find_replace'), query: z.string().min(1).max(1000), replacement: text, caseSensitive: z.boolean().optional(), slide: id.optional() }).strict(),
  z.object({ op: z.literal('add_slide'), index: z.number().int().min(1).max(100000).optional(), layout: id.optional() }).strict(),
  z.object({ op: z.literal('duplicate_slide'), slide: id, index: z.number().int().min(1).max(100000).optional() }).strict(),
  z.object({ op: z.literal('delete_slide'), slide: id }).strict(),
  z.object({ op: z.literal('move_slide'), slide: id, index: z.number().int().min(1).max(100000) }).strict(),
  z.object({ op: z.literal('add_text_box'), slide: id, name: z.string().min(1).max(200), rect: pptxRectSchema, text, style: pptxRunStyleSchema.optional() }).strict(),
  z.object({ op: z.literal('add_shape'), slide: id, name: z.string().min(1).max(200), rect: pptxRectSchema, geometry: z.enum(['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'line', 'rightArrow', 'chevron']), fill: color.nullable().optional() }).strict(),
  z.object({ op: z.literal('add_image'), slide: id, name: z.string().min(1).max(200), rect: pptxRectSchema, contentType: z.enum(['image/png', 'image/jpeg', 'image/gif', 'image/webp']), base64: z.string().min(4).max(11200000) }).strict(),
  z.object({ op: z.literal('set_notes'), slide: id, text }).strict(),
]);
export const pptxBatchSchema = z.object({ version: z.string().min(1).max(256), edits: z.array(pptxEditSchema).min(1).max(32) }).strict();
export type PptxAgentEdit = z.infer<typeof pptxEditSchema>;
export type PptxAgentBatch = z.infer<typeof pptxBatchSchema>;
