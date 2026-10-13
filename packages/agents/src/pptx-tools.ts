import { z } from 'zod';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { PptxAgentPresentation } from './pptx';
import { pptxEditSchema, pptxRectSchema } from './pptx-schema';
import { readPptxImage } from './pptx-image';
import { DocumentToolError } from './types';
import type { FileWorkspace } from './workspace';

export type PptxToolRegistrar = <S extends z.ZodRawShape>(name: string, description: string, shape: S, readOnly: boolean, run: (args: z.infer<z.ZodObject<S>>) => Promise<CallToolResult> | CallToolResult) => void;
const document = z.string().min(1).max(128).describe('Document ID from office_open.');
const id = z.string().min(1).max(1024);
const offset = z.number().int().nonnegative().max(Number.MAX_SAFE_INTEGER).optional();
const limit = z.number().int().min(1).max(100).optional();
const imageFile = z.object({
  op: z.literal('add_image_file'), slide: id, name: z.string().min(1).max(200), rect: pptxRectSchema,
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
  tool('pptx_outline', 'PPTX: list slides, titles, layout IDs/names, and shape IDs/types/positions/text summaries. Indices are one-based. Follow nextOffset, nextShapeOffset, and nextLayoutOffset.', { document, offset, limit, shapeOffset: offset, shapeLimit: limit, layoutOffset: offset }, true, args => result(deck(args.document).outline(args)));
  tool('pptx_read_slide', 'PPTX: read the shape tree, placeholders, paragraphs, formatted runs, and speaker notes. Copy slide/shape/story IDs. Follow nextOffset, nextStoryOffset, nextParagraphOffset, nextRunOffset, and nextStart for full content.', {
    document, slide: id, shape: id.optional(), story: id.optional(), offset, limit, storyOffset: offset, paragraphOffset: offset, runOffset: offset,
    textStart: offset, textLength: z.number().int().min(1).max(4000).optional(), notesStart: offset,
  }, true, args => result(deck(args.document).readSlide(args)));
  tool('pptx_preview', 'PPTX: view a slide as PNG after editing. Copy its slide ID from pptx_outline. scale defaults to 1; lower it for large slides. Returns preview warnings.', { document, slide: id, scale: z.number().min(0.25).max(3).optional() }, true, async args => {
    const { png, ...metadata } = await deck(args.document).preview(args.slide, args);
    const output = result(metadata);
    output.content.push({ type: 'image', mimeType: 'image/png', data: Buffer.from(png).toString('base64') });
    return output;
  });
  if (!readOnly) tool('pptx_edit', 'PPTX: apply 1–32 edits atomically using version from a fresh read. Replace shape text, set paragraphs/runs (bullet:true adds real bullets), find_replace, add/duplicate/delete/move slides, add_text_box/add_shape/add_image_file, or set_notes. Rectangles are points; indices are one-based final positions. add_slide is empty and links a layout ID from outline; it does not clone placeholders. Each add_text_box creates one shape. Put a bullet list in one body shape with set_paragraphs. add_image_file reads a workspace path; add_image accepts raster base64. Returns each edit’s changed IDs; export to save.', {
    document, version: z.string().min(1).max(256), edits: z.array(z.union([pptxEditSchema, imageFile])).min(1).max(32),
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
