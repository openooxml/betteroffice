import { z } from 'zod';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { PptxAgentPresentation } from './pptx';
import { pptxEditSchema, pptxRectSchema, pptxSlideIdSchema } from './pptx-schema';
import { readPptxImage } from './pptx-image';
import { DocumentToolError } from './types';
import { readLimit } from './tool-input';
import type { FileWorkspace } from './workspace';

export type PptxToolRegistrar = <S extends z.ZodRawShape>(name: string, description: string, shape: S, readOnly: boolean, run: (args: z.infer<z.ZodObject<S>>) => Promise<CallToolResult> | CallToolResult) => void;
const document = z.string().min(1).max(128).describe('Document ID from office_open.');
const id = z.string().min(1).max(1024);
const offset = z.number().int().nonnegative().max(Number.MAX_SAFE_INTEGER).optional();
const limit = readLimit(100);
const imageFile = z.object({
  op: z.literal('add_image_file'), slide: pptxSlideIdSchema, name: z.string().min(1).max(200), rect: pptxRectSchema,
  path: z.string().min(1).max(4096).describe('Image path relative to the workspace root.'),
}).strict();
function result(value: object): CallToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(value) }], structuredContent: value as Record<string, unknown> };
}

export function registerPptxTools(tool: PptxToolRegistrar, workspace: FileWorkspace, readOnly = false) {
  function deck(document: string) {
    const opened = workspace.get(document);
    if (!(opened instanceof PptxAgentPresentation)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'This tool requires PPTX. Use office_open on a .pptx file.');
    return opened;
  }
  tool('pptx_outline', 'PPTX: list slides, titles, layout IDs/names and their placeholders, and shape IDs/types/positions/text summaries. Slide positions are one-based; the middle number in an opened slide ID is its original 0-based index. Follow nextOffset, nextShapeOffset, nextLayoutOffset, and each layout’s nextPlaceholderOffset (pass layoutPlaceholderOffset).', { document, offset, limit, shapeOffset: offset, shapeLimit: limit, layoutOffset: offset, layoutPlaceholderOffset: offset }, true, args => result(deck(args.document).outline(args)));
  tool('pptx_read_slide', 'PPTX: read the shape tree, placeholders, paragraphs, formatted runs, and speaker notes. Text uses \\u000b for a soft line break (a:br); \\n separates paragraphs. Copy slide/shape/story IDs. Supply slide or shape; slide is inferred from shape and must agree if both are given. The middle number in an opened slide ID is its original 0-based index. Follow nextOffset, nextStoryOffset, nextParagraphOffset, nextRunOffset, and nextStart for full content.', {
    document, slide: pptxSlideIdSchema.optional(), shape: id.optional(), story: id.optional(), offset, limit, storyOffset: offset, paragraphOffset: offset, runOffset: offset,
    textStart: offset, textLength: readLimit(4000), notesStart: offset,
  }, true, args => result(deck(args.document).readSlide(args)));
  tool('pptx_preview', 'PPTX: view a slide as PNG after editing. Copy its slide ID from pptx_outline. scale defaults to 1; lower it for large slides. Returns preview warnings.', { document, slide: pptxSlideIdSchema, scale: z.number().min(0.25).max(3).optional() }, true, async args => {
    const { png, ...metadata } = await deck(args.document).preview(args.slide, args);
    const output = result(metadata);
    output.content.push({ type: 'image', mimeType: 'image/png', data: Buffer.from(png).toString('base64') });
    return output;
  });
  if (!readOnly) tool('pptx_edit', 'PPTX: apply 1-32 edits atomically using version from a fresh read. Ops: replace_text {shape, text, slide?, story?}; set_paragraphs {shape, paragraphs:[{runs:[{text}], bullet?}], slide?, story?}; find_replace {query, replacement, shape?, slide?, caseSensitive?}; add_slide {index?, layout?}; set_slide_layout {slide, layout, placeholders:[{shape, layoutPlaceholder}]?}; duplicate_slide {slide, index? or after?}; delete_slide {slide}; move_slide {slide, index or after or before}; add_text_box {slide, name, rect, text, style?}; add_shape {slide, name, rect, geometry, fill?}; add_image {slide, name, rect, contentType, base64}; add_image_file {slide, name, rect, path}; set_notes {slide, text}; add_comment {slide, author, text, initials?, date?, position?}; reply_comment {comment, author, text, initials?, date?}; resolve_comment {comment, resolved?}; insert_table_row {shape+row or ref, position, cells?, slide?}; delete_table_row {shape+row or ref, slide?}; insert_table_column {shape+column or ref, position, cells?, slide?}; delete_table_column {shape+column or ref, slide?}. Shape targets infer slide; an explicit slide must agree. Table edits take shape plus a zero-based row/column index, or any cell paragraph ref in ref (also accepted in row/column); missing cells are empty. Structural edits refresh table refs; read again. Merged tables and deleting the last row/column are refused. Table structure edits are unavailable in collaborative sessions. List discussions with office_outline story comments and read full text with office_read. add_comment takes slide, author, text, optional initials/date/position in points. reply_comment targets a comment ID and its thread root in either flavor; resolve_comment defaults resolved to true, false reopens, and requires modern comments. duplicate_slide copies an independent discussion. replace_text replaces the whole shape: \\n starts a paragraph, \\u000b inserts a soft line break. Unchanged paragraph text, including removal of trailing soft breaks, keeps run formatting; other text uses the first run\'s format for that paragraph. For partial edits use find_replace or office_propose. Copy shape and layout placeholder IDs from pptx_outline for placeholders:[{shape, layoutPlaceholder}]; unmapped shapes stay as they are. Rectangles are points; positions are one-based, while the middle number in an opened slide ID is the original 0-based index. duplicate_slide defaults to immediately after its source; supply either index or after. add_slide is empty and links a layout; it does not clone placeholders. Each add_text_box creates one shape. Put a bullet list in one body shape with set_paragraphs and bullet:true. add_image_file reads a workspace path. Returns changed IDs; export to save.', {
    document, version: z.string().min(1).max(256), edits: z.array(z.discriminatedUnion('op', [...pptxEditSchema.options, imageFile])).min(1).max(32),
  }, false, async args => {
    const opened = deck(args.document);
    const edits = [];
    let imageBytes = 0;
    for (const [index, edit] of args.edits.entries()) {
      if (edit.op !== 'add_image_file') {
        if (edit.op === 'add_image') imageBytes += edit.base64.length;
        edits.push(edit);
      } else {
        try {
          const { path, ...shape } = edit;
          const image = await readPptxImage(workspace.root, path);
          imageBytes += image.base64.length;
          edits.push({ ...shape, op: 'add_image' as const, ...image });
        } catch (error) {
          const failure = error instanceof DocumentToolError ? error : new DocumentToolError('INVALID_IMAGE', 'Image file cannot be read. Choose an existing workspace image path.');
          throw new DocumentToolError(failure.code, `${failure.message} No batch edits were applied.`, { ...failure.details, failedIndex: index, results: args.edits.map((item, index) => ({ index, op: item.op, applied: false })) });
        }
      }
      if (imageBytes > 11200000) throw new DocumentToolError('EDIT_LIMIT', 'Batch images exceed 8 MiB. Split the batch; no edits were applied.');
    }
    const receipt = opened.edit({ version: args.version, edits });
    return result({ ...receipt, results: receipt.results.map((item, index) => ({ ...item, op: args.edits[index].op })) });
  });
}
