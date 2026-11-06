import { pptxProposalEditSchema } from './pptx-schema';
import { PptxAgentPresentation } from './pptx';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { CallToolRequestSchema, ListToolsRequestSchema, type CallToolResult, type Tool } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import { clampingNotes, dateSchema, readLimit, toolInputIssues } from './tool-input';
import { DocumentToolError, type DocumentRenderer } from './types';
import { XlsxAgentWorkbook } from './xlsx';
import { registerXlsxTools } from './xlsx-mcp';
import { FileWorkspace } from './workspace';
import { registerPptxTools } from './pptx-tools';
import { DocxAgentDocument } from './document';
import { docxEditSchema, docxVersionSchema } from './docx-schema';
import { editLimit } from './edit-limits';
import { version as packageVersion } from '../package.json';
import { buildInstructions, isToolEnabled, normalizeFormats, normalizeTools, type OfficeFormat } from './tool-filter';

export type { OfficeFormat } from './tool-filter';

const identifier = z.string().min(1).max(128);
const path = z.string().min(1).max(4096);
const document = identifier.describe('Open document ID returned by office_open, such as doc1.');
const proposal = identifier.describe('Proposal ID returned by office_propose.');
const proposalTextEdit = z.object({ match: identifier.describe('Match ID from office_grep (not the text).'), newText: z.string().max(16000) }).strict();
const proposalEdit = z.union([proposalTextEdit, pptxProposalEditSchema]).meta({
  properties: z.toJSONSchema(proposalTextEdit, { io: 'input' }).properties,
});
const limit = readLimit(100);
const offset = z.number().int().nonnegative().max(Number.MAX_SAFE_INTEGER).optional();

function result(value: object): CallToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(value) }], structuredContent: value as Record<string, unknown> };
}

function failure(error: unknown, schema?: unknown, input?: unknown): CallToolResult {
  if (error instanceof z.ZodError) {
    const issues = toolInputIssues(error, schema, input);
    const shown = issues.slice(0, 5).map(issue => `${issue.path.map(String).join('.')}: ${issue.message}`);
    return { ...result({ code: 'INVALID_ARGUMENT', message: [...shown, ...issues.length > 5 ? [`+${issues.length - 5} more`] : []].join('; '), issues }), isError: true };
  }
  const system = error as { code?: string };
  const code = error instanceof DocumentToolError ? error.code : system?.code ?? 'ENGINE_ERROR';
  const message = code === 'EEXIST' ? 'Output already exists. Choose a new export path.' : error instanceof Error ? error.message : String(error);
  return { ...result({ code, message, ...(error instanceof DocumentToolError && error.details ? { details: error.details } : {}) }), isError: true };
}

export async function createOfficeMcpServer(options: { root: string; renderer?: DocumentRenderer; readOnly?: boolean; formats?: readonly OfficeFormat[]; tools?: readonly string[]; maxEdits?: number }) {
  const maxEdits = editLimit(options.maxEdits);
  const formats = normalizeFormats(options.formats);
  const toolFilter = normalizeTools(options.tools);
  const workspace = await FileWorkspace.create(options.root, options.renderer, formats, maxEdits);
  const server = new Server({ name: 'betteroffice', version: packageVersion }, {
    capabilities: { tools: {} },
    instructions: buildInstructions(formats),
  });
  let queue = Promise.resolve();
  const tools: Tool[] = [];
  const handlers = new Map<string, (args: unknown) => Promise<CallToolResult>>();
  function tool<S extends z.ZodRawShape>(name: string, description: string, shape: S, readOnly: boolean, run: (args: z.infer<z.ZodObject<S>>) => Promise<CallToolResult> | CallToolResult) {
    if (!isToolEnabled(name, formats, toolFilter)) return;
    const schema = z.object(shape).strict();
    const inputSchema = z.toJSONSchema(schema, { io: 'input' }) as Tool['inputSchema'];
    tools.push({ name, description, inputSchema, annotations: { readOnlyHint: readOnly, destructiveHint: false, openWorldHint: false } });
    handlers.set(name, async args => {
      const task = queue.then(async () => {
        try {
          const parsed = schema.parse(args);
          const output = await run(parsed);
          const clamping = clampingNotes(args, parsed);
          if (!clamping.length || !output.structuredContent || output.isError) return output;
          const annotated = result({ ...output.structuredContent, clamping });
          annotated.content.push(...output.content.filter(block => block.type !== 'text'));
          return { ...output, ...annotated };
        } catch (error) { return failure(error, inputSchema, args); }
      });
      queue = task.then(() => undefined, () => undefined);
      return task;
    });
  }
  registerXlsxTools(tool, workspace, options.readOnly);
  tool('office_files', 'List DOCX, XLSX, and PPTX files and subdirectories under the workspace, plus open document IDs. Start here when the path is unknown.', { directory: path.optional(), offset }, true, async args => result(await workspace.files(args.directory, args.offset)));
  tool('office_open', 'Open an Office file and return its document ID, format, and capabilities. Paths are relative to the configured workspace root.', { path }, true, async args => result(await workspace.open(args.path)));
  tool('office_outline', 'List DOCX paragraphs, PPTX slide paragraphs, or XLSX sheets (sheetId and name). For DOCX/PPTX comments use story \'comments\' to list IDs, parentId, author, date and bounded text. PPTX also lists initials, resolved/status, slideId, slide number and position in points; office_read accepts the comment ref for full text. DOCX includes done and anchor refs. Filter story; headingsOnly is DOCX-only. DOCX table cell paragraphs include table {story,index,row,column,rows,columns}, one-based, plus rowSpan/colSpan for merged cells; list paragraphs include list {level}, zero-based. DOCX sections lists one-based numbers, first/last paragraph refs, orientation, size and margins in twips; paginate independently with sectionOffset, sectionLimit and sections.nextOffset. Paginate paragraphs with nextOffset.', { document, story: path.optional(), headingsOnly: z.boolean().optional(), offset, limit, sectionOffset: offset, sectionLimit: limit }, true, args => result(workspace.get(args.document).list(args)));
  tool('office_grep', 'Find literal text with bounded context. DOCX/PPTX return match IDs for office_propose; XLSX returns cell handles for office_propose_cells, searching displayed values case-sensitively. story filters a DOCX/PPTX story or XLSX sheetId. Follow nextCursor; if truncated remains true, narrow query or story.', { document, query: z.string().min(1).max(1000), caseSensitive: z.boolean().optional(), story: path.optional(), limit, cursor: identifier.optional() }, true, args => result(workspace.get(args.document).grep(args)));
  tool('office_read', 'Read a ref from outline/grep/cells: a DOCX/PPTX paragraph, PPTX comment, or XLSX cell display text. Default 4000 UTF-16 units; follow nextStart. DOCX includes formatting in points and the table position of cell paragraphs; XLSX includes a bounded value/formula summary; field selects displayText (default), formula, or value for paginated full text.', { document, ref: identifier, start: offset, length: readLimit(16000), field: z.enum(['displayText', 'formula', 'value']).optional() }, true, args => {
    const opened = workspace.get(args.document);
    if (args.field && !(opened instanceof XlsxAgentWorkbook)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'field is an XLSX-only read option.');
    return result(opened.read(args.ref, args));
  });
  tool('docx_read_runs', 'DOCX: read paragraphs with stable run IDs and a deduplicated formats dictionary for docx_edit set_runs. IDs require the paragraph revision returned here. Fixed tokens {id,kind} cannot be edited. Default 50 paragraphs, maximum 200; maxChars defaults to 32000 UTF-16 units and stops at a paragraph boundary, returning at least one paragraph. Follow nextOffset.', {
    document, story: path.optional(), offset, limit: z.number().int().min(1).max(200).optional(), maxChars: z.number().int().min(1).max(Number.MAX_SAFE_INTEGER).optional(),
  }, true, args => {
    const opened = workspace.get(args.document);
    if (!(opened instanceof DocxAgentDocument)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'docx_read_runs requires a DOCX document.');
    return result(opened.readRuns(args));
  });
  tool('office_cells', 'XLSX: read an A1 range and cell handles for staged edits with office_propose_cells. For direct range edits, use xlsx_read_range and xlsx_write_range. Includes empty cells, values and formulas; follow nextOffset. Long text uses office_read.', { document, sheet: identifier, range: z.string().min(1).max(40), offset, limit }, true, args => {
    const workbook = workspace.get(args.document);
    if (!(workbook instanceof XlsxAgentWorkbook)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'office_cells requires an XLSX workbook.');
    return result(workbook.readCells(args));
  });
  tool('office_render', 'DOCX only: view a page as PNG using the document engine. page is one-based. Supply proposal to see its proposed result without changing the live document. Returns pageCount and render warnings.', { document, page: z.number().int().min(1).max(100000).default(1), proposal: proposal.optional() }, true, async args => {
    const opened = workspace.get(args.document);
    if (opened.overview().format === 'pptx') throw new DocumentToolError('UNSUPPORTED', 'Use pptx_preview with a slide ID from pptx_outline.');
    const { png, ...metadata } = await opened.render(args.page, args.proposal);
    const output = result(metadata);
    output.content.push({ type: 'image', mimeType: 'image/png', data: Buffer.from(png).toString('base64') });
    return output;
  });
  tool('office_verify', 'Save and reopen the current document or a pending proposal in memory. Reports completed checks explicitly; does not assert semantic or visual correctness. DOCX also compares the saved package with the original (structure check, on unless structure is false): parts, element counts, field instructions, drawings and relationships. structure.lost lists losses the edits do not explain and sets ok false; structure.unattributed lists drops it cannot attribute after edits that removed content, to review even when ok is true; structure.summary says so. Also reports keptUnchanged and structureDropped like office_export.', { document, proposal: proposal.optional(), structure: z.boolean().optional() }, true, async args => result(await workspace.get(args.document).verify(args.proposal, { structure: args.structure })));
  tool('office_close', 'Release an open document and its in-memory proposals. Export any work you want to keep first.', { document }, false, args => result(workspace.close(args.document)));
  if (!options.readOnly) {
    tool('office_propose', `DOCX/PPTX: stage 1 to ${maxEdits} text replacements without changing the document. XLSX uses office_propose_cells. First grep for the exact text you want to replace, then copy the desired occurrence\'s match ID. Supply edits [{match: ID_FROM_GREP, newText: REPLACEMENT}]. To change part of a sentence, grep that part; do not reuse a match for the whole sentence. New text inherits formatting at its start. No paragraph/embedded-content/tracked-change crossings. PPTX also accepts slide edits with type insertSlide {ref?, afterSlideId, layoutPartPath?}, duplicateSlide {ref?, slideId, afterSlideId?}, deleteSlide {slideId}, moveSlide {slideId, afterSlideId}. afterSlideId null means first; omitted on duplicateSlide means after its source. PPTX also accepts addTextBox, addShape, addPicture with {ref?, slideId, draft}, and removeShape with {slideId, shapeId}. Draft rectangles use EMU (12700 per point), and picture drafts carry contentType and mediaBase64 with an 8 MB decoded total limit. Later slide IDs, shape IDs and anchors may use @ref. replaceText, formatText and setParagraphAlignment may target storyId @ref for a created shape first text story. PPTX also accepts setStoryParagraphs {storyId, paragraphs: [{alignment?, bullet?, runs: [{text, style?}]}]} for whole-story rewrites and bullet lists. It accepts 1 to 100 paragraphs, up to 64 runs per paragraph and 16000 UTF-16 text units total. storyId may be @ref for a created shape first text story. PPTX also accepts setSlideLayout {slideId, layoutPartPath, mapping?: [[shapeId, layoutPlaceholderId]]} and setSlideBackground {slideId, background: ShapeFill | null}. Null inherits the layout or master background. Both accept slideId @ref. PPTX also accepts insertTableRow and insertTableColumn {slideId, shapeId, index, position: before | after, cells?}, deleteTableRow and deleteTableColumn {slideId, shapeId, index}. Indices are zero-based. Column inserts accept width: split (default) | copy. Table targets accept slideId and shapeId @ref where available; cells follow row or column order. PPTX also accepts reorderShape {slideId, shapeId, to: front | back | forward | backward} for top-level shapes. Both IDs may use @ref from earlier edits. Boundary reorders are accepted without changes.`, {
      document, author: z.string().min(1).max(200), note: z.string().max(2000).optional(),
      edits: z.array(proposalEdit).min(1).max(maxEdits),
    }, false, async args => {
      const target = workspace.get(args.document);
      if (target instanceof PptxAgentPresentation) return result(target.propose(args));
      if (args.edits.some(edit => 'type' in edit)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'Slide proposals require PPTX.');
      const edits = args.edits.filter((edit): edit is { match: string; newText: string } => 'match' in edit);
      return result(target instanceof DocxAgentDocument ? await target.proposeAsync({ ...args, edits }) : target.propose({ ...args, edits }));
    });
    tool('office_propose_cells', 'XLSX: stage 1–32 whole-cell edits using handles from office_cells or office_grep. input parses typed numbers, text and =formulas; empty clears. Review then office_accept, or office_export with proposal. For immediate matrix writes, use xlsx_write_range.', {
      document, author: z.string().min(1).max(200), note: z.string().max(2000).optional(),
      edits: z.array(z.object({ cell: identifier, input: z.string().max(16000) }).strict()).min(1).max(32),
    }, false, args => {
      const workbook = workspace.get(args.document);
      if (!(workbook instanceof XlsxAgentWorkbook)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'office_propose_cells requires an XLSX workbook.');
      return result(workbook.proposeCells(args));
    });
    tool('office_review', 'List proposals, or supply proposal to inspect exact before/after changes, attribution, status, and staleness. A stale proposal must be re-created from fresh reads.', { document, proposal: proposal.optional() }, true, args => result(args.proposal ? workspace.get(args.document).review(args.proposal) : { proposals: workspace.get(args.document).listProposals() }));
    tool('office_accept', 'Apply a proposal to the open document atomically. Stale targets are rejected. DOCX only: use tracked=true to retain native Word tracked changes; date is an ISO timestamp or YYYY-MM-DD (midnight UTC) and defaults to now. Export afterwards to save a file.', { document, proposal, tracked: z.boolean().optional(), date: dateSchema.optional() }, false, async args => result(await workspace.get(args.document).accept(args.proposal, args)));
    tool('office_reject', 'Discard a pending proposal without editing the document.', { document, proposal }, false, async args => {
      const opened = workspace.get(args.document);
      return result(opened instanceof DocxAgentDocument ? await opened.rejectAsync(args.proposal) : opened.reject(args.proposal));
    });
    tool('docx_edit', `DOCX: apply 1-${maxEdits} edits atomically and directly to the open document (no proposal, no tracked changes) using version from office_outline, office_grep or docx_read_runs. set_paragraph_style takes a style ID or name defined in the document. set_language sets or clears proofing language slots on one paragraph, grep match, style or document defaults. format_text sets (true) or clears (false) bold/italic/underline on exactly one grep match. add_comment anchors a Word comment on exactly one grep match in the body or a table; each text line becomes a comment paragraph; date defaults to now. List comments via office_outline story \'comments\', then reply_comment or resolve_comment by comment ID. Both target the thread root; resolve_comment defaults done to true, false reopens. replace_comment_text replaces every occurrence of find in one comment\'s text, keeping its formatting. delete_comment removes a comment, or the whole thread for a root; it cannot share a batch with other comment operations. insert_paragraphs adds plain paragraphs after a ref in the body default style unless style is given, or before a ref copying that paragraph\'s style and paragraph formatting; runFormat: anchor copies the ref\'s first plain run instead of style formatting (default), incompatible with list. With list set to continue after a list item (outline list) they join its list at its level, with its formatting. delete_paragraph keeps at least one paragraph per cell or story. insert_table_row and delete_table_row take the ref of any paragraph in the row (cell paragraphs show table in outline/read); cells has one text per column and copies the row and cell formatting. merge_cells merges the rectangle between the cells holding from and to (refs of any paragraph in two corner cells of one table) into its top-left cell, keeping every paragraph in reading order; refs lists the paragraphs of the merged cell. insert_section_break requires exactly one before or after body paragraph ref and kind nextPage, continuous, evenPage or oddPage; before the first paragraph is refused. set_section uses one-based section numbers from the supplied outline version and optional orientation, pageWidth, pageHeight and margins in twips. Refresh the outline to address new sections. set_runs takes {paragraph, revision, runs:[{run,text}]} from docx_read_runs, replacing only listed text runs in place with their formatting; empty text deletes, fixed tokens are refused. Use one set_runs per paragraph. Returns new refs, comment IDs and the version. Use office_propose for staged text replacements.`, {
      document, version: docxVersionSchema, edits: z.array(docxEditSchema).min(1).max(maxEdits),
    }, false, async args => {
      const opened = workspace.get(args.document);
      if (!(opened instanceof DocxAgentDocument)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'docx_edit requires a DOCX document.');
      return result(await opened.edit({ version: args.version, edits: args.edits }));
    });
    tool('office_export', 'Write a new file of the opened format (.docx, .xlsx, .pptx) inside the workspace. Existing files are never overwritten. Supply proposal to export a proposed result without accepting it; omit proposal to export accepted edits. DOCX: keptUnchanged.structure counts paragraphs written as they were, without their new formatting, because writing it would have dropped a hyperlink, field or drawing. keptUnchanged.paragraphs names up to 200 of them: ref and text, the dropped elements, and scope "part" when the whole part kept its source; tell the user, or retry a narrower range. structureDropped lists parts that had to be written whole with fewer structural elements than their source, and the paragraphs affected; your edits may account for them, otherwise tell the user.', { document, path, proposal: proposal.optional() }, false, async args => result(await workspace.export(args.document, args.path, args.proposal)));
  }
  registerPptxTools(tool, workspace, options.readOnly);
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));
  server.setRequestHandler(CallToolRequestSchema, async request => {
    const handler = handlers.get(request.params.name);
    return handler ? handler(request.params.arguments ?? {}) : failure(new DocumentToolError('UNKNOWN_TOOL', 'Use tools/list to discover supported tools.'));
  });
  server.onclose = () => workspace.dispose();
  return server;
}
