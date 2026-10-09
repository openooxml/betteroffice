import { imageSize } from 'image-size';
import { pptxBatchSchema, type PptxAgentBatch, type PptxAgentEdit } from './pptx-schema';
import { renderPptxSlide } from './pptx-render';
import { initWasm, openPresentation, type PptxEditStep, type PptxStoryText, type PptxStoryTarget, type PresentationHandle, type ShapeSnapshot, type SlideSnapshot, type ShapeRect, type StorySnapshot, type TextStyle } from '@betteroffice/pptx';
import { DocumentToolError, type GrepOptions, type TextEdit, type RenderedPage } from './types';
import { initialize, integer, plainText, PrototypeDocument, textWindow, unwrap, type PrototypeOptions } from './prototype';

export interface SlideTextChange { ref: string; slide: number; story: string; start: number; oldText: string; newText: string }
interface Paragraph { ref: string; slide: number; story: PptxStoryText; start: number; end: number; editable: boolean }
interface HeldMatch { ref: string; slide: number; target: PptxStoryTarget; start: number; paragraphStart: number; text: string; version: string }

export class PptxAgentPresentation extends PrototypeDocument<SlideTextChange, PptxEditStep> {
  readonly format = 'pptx' as const;
  private readonly refs = new Map<string, string>();
  private readonly matches = new Map<string, HeldMatch>();
  private nextMatch = 1;
  private nextRef = 1;
  private cachedLayouts?: ReturnType<PresentationHandle['layouts']>;

  private constructor(private readonly handle: PresentationHandle, readonly name = 'presentation.pptx', private readonly owned = true) { super(); }

  static async open(bytes: Uint8Array, options: PrototypeOptions = {}) {
    await initialize('pptx', initWasm, options.wasm);
    if (bytes.length > 64 * 1024 * 1024) throw new DocumentToolError('FILE_TOO_LARGE', 'Open a PPTX file up to 64 MiB.');
    const source = bytes.slice();
    return new PptxAgentPresentation(openPresentation(source), options.name);
  }

  static attach(handle: PresentationHandle, options: { name?: string } = {}) {
    const required = ['fork', 'layouts', 'snapshot', 'save', 'applyUpdate', 'encodeStateVector', 'version'] as const;
    if (!handle || required.some(method => typeof handle[method] !== 'function')) throw new DocumentToolError('UNSUPPORTED_SESSION', 'attachPptx requires a PresentationHandle from openPresentation. Worker editor access does not expose state forking or update application.');
    return new PptxAgentPresentation(handle, options.name, false);
  }

  overview() {
    const read = this.content();
    return { name: this.name, format: this.format, version: read.version, slides: read.slides.length, paragraphs: this.paragraphs(read).length, capabilities: { outline: true, readSlide: true, atomicEdits: true, render: true, export: true, attachment: 'PresentationHandle', proposals: 'paragraph-local text replacements', prototype: false }, workflow: 'Use pptx_outline for slide/shape/layout IDs, pptx_read_slide for paragraphs, runs and notes, pptx_edit with the current version for atomic edits, pptx_preview to verify visually, then office_export. Coordinates for edits are points. Attached handles remain owned by the editor.' };
  }

  list(options: { story?: string; offset?: number; limit?: number; headingsOnly?: boolean } = {}) {
    if (options.headingsOnly) throw new DocumentToolError('UNSUPPORTED', 'PPTX outline lists slide paragraphs; omit headingsOnly.');
    const read = this.content();
    this.checkStory(options.story, read.stories);
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 30, 1, 100, 'limit');
    const items = this.paragraphs(read).filter(p => !options.story || p.story.storyId === options.story);
    return { version: read.version, items: items.slice(offset, offset + limit).map(p => ({ ...this.reference(p), text: p.story.text.slice(p.start, Math.min(p.end, p.start + 120)), length: p.end - p.start, editable: p.editable })), total: items.length, nextOffset: offset + limit < items.length ? offset + limit : null };
  }

  grep(options: GrepOptions) {
    this.checkQuery(options);
    const read = this.content();
    this.checkStory(options.story, read.stories);
    const pattern = new RegExp(options.query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), options.caseSensitive ? 'gu' : 'giu');
    const hits: Array<{ paragraph: Paragraph; start: number; text: string }> = [];
    outer: for (const p of this.paragraphs(read)) {
      if (options.story && options.story !== p.story.storyId) continue;
      const text = p.story.text.slice(p.start, p.end);
      pattern.lastIndex = 0;
      for (let match = pattern.exec(text); match; match = pattern.exec(text)) {
        hits.push({ paragraph: p, start: p.start + match.index, text: match[0] });
        if (hits.length > 10000) break outer;
      }
    }
    const clipped = hits.length > 10000;
    const page = this.searchPage(options, hits.slice(0, 10000).map(hit => ({ ...this.reference(hit.paragraph), start: hit.start - hit.paragraph.start, text: hit.text, context: hit.paragraph.story.text.slice(Math.max(hit.paragraph.start, hit.start - 80), Math.min(hit.paragraph.end, hit.start + hit.text.length + 80)) })), clipped);
    return { ...page, matches: page.matches.map(hit => {
      const p = hits.find(item => item.paragraph.ref === hit.ref && item.start - item.paragraph.start === hit.start)!.paragraph;
      const match = `${this.prefix}:match${this.nextMatch++}`;
      if (this.matches.size >= 1024) this.matches.delete(this.matches.keys().next().value!);
      this.matches.set(match, { ref: p.ref, slide: p.slide, target: { slideId: p.story.slideId, shapeId: p.story.shapeId, storyId: p.story.storyId }, start: p.start + hit.start, paragraphStart: p.start, text: hit.text, version: read.version });
      return { ...hit, match };
    }) };
  }

  read(ref: string, options: { start?: number; length?: number } = {}) {
    const read = this.content();
    const p = this.paragraphs(read).find(p => p.ref === ref);
    if (!p) throw new DocumentToolError('UNKNOWN_REF', 'Use a ref from outline or grep.');
    return { ...this.reference(p), version: read.version, ...textWindow(p.story.text.slice(p.start, p.end), options), editable: p.editable };
  }

  propose(input: { author: string; note?: string; edits: TextEdit[] }) {
    this.alive();
    if (!Array.isArray(input.edits) || input.edits.length < 1 || input.edits.length > 32) throw new DocumentToolError('INVALID_ARGUMENT', 'Supply 1–32 replacements.');
    const version = this.currentVersion();
    const changes: SlideTextChange[] = [];
    const steps: PptxEditStep[] = [];
    for (const edit of input.edits) {
      if (!('match' in edit)) throw new DocumentToolError('UNSUPPORTED', 'PPTX replacements require a match ID from grep.');
      const hit = this.matches.get(edit.match);
      if (!hit) throw new DocumentToolError('UNKNOWN_MATCH', 'Match expired or belongs to another deck. Grep again.');
      if (hit.version !== version) throw new DocumentToolError('STALE_TARGET', 'Presentation changed. Grep for fresh matches.');
      plainText(edit.newText);
      if (/[\r\n\ufffc]/u.test(edit.newText) || edit.newText.length > 16000) throw new DocumentToolError('INVALID_ARGUMENT', 'Replacement must be plain text in one paragraph, up to 16000 UTF-16 units.');
      changes.push({ ref: hit.ref, slide: hit.slide, story: hit.target.storyId, start: hit.start - hit.paragraphStart, oldText: hit.text, newText: edit.newText });
      steps.push({ op: 'replaceText', target: { kind: 'range', ...hit.target, start: hit.start, end: hit.start + hit.text.length }, text: edit.newText, expect: { text: hit.text } });
    }
    return this.stage(input, changes, steps, version);
  }

  outline(options: { offset?: number; limit?: number; shapeOffset?: number; shapeLimit?: number; layoutOffset?: number } = {}) {
    const deck = this.snapshot();
    const offset = integer(options.offset ?? 0, 0, deck.slides.length, 'offset');
    const limit = integer(options.limit ?? 10, 1, 100, 'limit');
    const shapeOffset = integer(options.shapeOffset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'shapeOffset');
    const shapeLimit = integer(options.shapeLimit ?? 10, 1, 100, 'shapeLimit');
    const layoutOffset = integer(options.layoutOffset ?? 0, 0, this.layouts.length, 'layoutOffset');
    const layouts = this.bounded(this.layouts.slice(layoutOffset, layoutOffset + 30), 8000);
    const items = [];
    let size = 0;
    for (const slide of deck.slides.slice(offset, offset + limit)) {
      const shapes = this.flatten(slide.shapes);
      const selected = this.bounded(shapes.slice(shapeOffset, shapeOffset + shapeLimit).map(node => this.shapeSummary(node.shape, node.parentId, node.depth)), 12000);
      const item = { ...this.slideSummary(slide, deck.slides.indexOf(slide)), shapes: selected, shapeCount: shapes.length, nextShapeOffset: shapeOffset + selected.length < shapes.length ? shapeOffset + selected.length : null };
      const cost = JSON.stringify(item).length;
      if (items.length && size + cost > 24000) break;
      items.push(item); size += cost;
    }
    return { version: this.handle.version(), size: this.rect({ x: 0, y: 0, width: deck.widthEmu, height: deck.heightEmu }), slides: items, total: deck.slides.length, nextOffset: offset + items.length < deck.slides.length ? offset + items.length : null, layouts, nextLayoutOffset: layoutOffset + layouts.length < this.layouts.length ? layoutOffset + layouts.length : null, ids: 'IDs remain stable during this session. Opened exports receive fresh IDs.' };
  }

  readSlide(options: { slide: string; shape?: string; story?: string; offset?: number; limit?: number; storyOffset?: number; paragraphOffset?: number; runOffset?: number; textStart?: number; textLength?: number; notesStart?: number } ) {
    const deck = this.snapshot();
    const slide = this.slide(this.handle, options.slide);
    const all = this.flatten(slide.shapes);
    const shape = options.shape ? this.shape(slide, options.shape) : undefined;
    const selectedIds = shape ? new Set(this.flatten([shape]).map(node => node.shape.id)) : undefined;
    const nodes = selectedIds ? all.filter(node => selectedIds.has(node.shape.id)) : all;
    if (options.story && !nodes.some(node => node.shape.textStories.some(story => story.id === options.story))) throw new DocumentToolError('UNKNOWN_STORY', 'Copy a story ID from pptx_read_slide.', { validStories: nodes.flatMap(node => node.shape.textStories.map(story => story.id)).slice(0, 20) });
    const offset = integer(options.offset ?? 0, 0, nodes.length, 'offset');
    const limit = integer(options.limit ?? 10, 1, 100, 'limit');
    const storyOffset = integer(options.storyOffset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'storyOffset');
    const paragraphOffset = integer(options.paragraphOffset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'paragraphOffset');
    const runOffset = integer(options.runOffset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'runOffset');
    const textLength = integer(options.textLength ?? 1000, 1, 4000, 'textLength');
    const textStart = integer(options.textStart ?? 0, 0, Number.MAX_SAFE_INTEGER, 'textStart');
    const shapes = [];
    let size = 0;
    for (const node of nodes.slice(offset, offset + limit)) {
      const stories = node.shape.textStories.filter(story => !options.story || story.id === options.story);
      const selected = stories.slice(storyOffset, storyOffset + 1).map(story => {
        const fullText = story.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n');
        const window = textWindow(fullText, { start: Math.min(textStart, fullText.length), length: textLength });
        let position = 0;
        const paragraphs = story.paragraphs.map(paragraph => {
          const start = position;
          const runs = paragraph.runs.map(run => {
            const start = position;
            position += run.text.length;
            return { start, end: position, text: run.text.slice(Math.max(0, window.start - start), Math.max(0, Math.min(run.text.length, window.end - start))), formatting: { ...run.style, fontFamily: run.style.fontFamily?.slice(0, 200) ?? null, underline: run.style.underline?.slice(0, 32) ?? null } };
          });
          const end = position++;
          const visible = runs.filter(run => run.end > window.start && run.start < window.end);
          const page = visible.slice(runOffset, runOffset + 8);
          return { id: paragraph.id, start, end, alignment: paragraph.alignment, level: paragraph.level, bullet: paragraph.bulletJson?.slice(0, 256) ?? null, runs: page, runCount: paragraph.runs.length, nextRunOffset: runOffset + page.length < visible.length ? runOffset + page.length : null };
        }).filter(p => p.end >= window.start && p.start <= window.end);
        const selected = paragraphs.slice(paragraphOffset, paragraphOffset + 4);
        return { id: story.id, ...window, paragraphs: selected, paragraphCount: paragraphs.length, nextParagraphOffset: paragraphOffset + selected.length < paragraphs.length ? paragraphOffset + selected.length : null };
      });
      const item = { ...this.shapeSummary(node.shape, node.parentId, node.depth), rotationDeg: node.shape.rotationDeg, flipH: node.shape.flipH, flipV: node.shape.flipV, hidden: node.shape.hidden ?? false, fill: node.shape.resolvedFillColor, outline: node.shape.resolvedOutlineColor, mediaPartPath: node.shape.mediaPartPath, stories: selected, storyCount: stories.length, nextStoryOffset: storyOffset + selected.length < stories.length ? storyOffset + selected.length : null };
      const cost = JSON.stringify(item).length;
      if (shapes.length && size + cost > 24000) break;
      shapes.push(item); size += cost;
    }
    return { version: this.handle.version(), ...this.slideSummary(slide, deck.slides.findIndex(s => s.id === slide.id)), shapes, shapeCount: nodes.length, nextOffset: offset + shapes.length < nodes.length ? offset + shapes.length : null, notes: textWindow(slide.notes ?? '', { start: options.notesStart, length: 4000 }) };
  }

  edit(input: PptxAgentBatch) {
    this.alive();
    const parsed = pptxBatchSchema.safeParse(input);
    if (!parsed.success) throw new DocumentToolError('INVALID_ARGUMENT', 'Use the pptx_edit schema: version and 1–32 edits.', { issues: parsed.error.issues.map(issue => ({ path: issue.path, message: issue.message })) });
    const version = this.handle.version();
    if (input.version !== version) throw new DocumentToolError('STALE_TARGET', 'Deck changed. Read pptx_outline or pptx_read_slide again and copy its version.');
    const imageBytes = parsed.data.edits.reduce((sum, edit) => sum + (edit.op === 'add_image' ? edit.base64.length : 0), 0);
    const textBytes = JSON.stringify(parsed.data.edits.map(edit => edit.op === 'add_image' ? { ...edit, base64: '' } : edit)).length;
    if (imageBytes > 11200000 || textBytes > 64000) throw new DocumentToolError('EDIT_LIMIT', 'Batch allows at most 8 MiB of images and 64000 characters of other input. Split the batch.');
    const draft = this.handle.fork();
    const vector = this.handle.encodeStateVector();
    const results: Array<Record<string, unknown>> = [];
    let stepIndex = 0;
    try {
      for (const edit of parsed.data.edits) {
        results.push({ index: stepIndex, op: edit.op, ...this.applyOperation(draft, edit, Math.floor(8000 / parsed.data.edits.length)) });
        stepIndex++;
      }
      draft.save();
      if (this.handle.version() !== version) throw new DocumentToolError('STALE_TARGET', 'Deck changed during the batch. Read fresh IDs and version, then retry.');
      this.handle.applyUpdate(draft.encodeStateAsUpdate(vector));
      return { applied: true, version: this.handle.version(), results };
    } catch (error) {
      const failure = error instanceof DocumentToolError ? error : new DocumentToolError('INVALID_EDIT', error instanceof Error ? error.message : String(error));
      throw new DocumentToolError(failure.code, `${failure.message} No batch edits were applied.`, { ...failure.details, failedIndex: Math.min(stepIndex, input.edits.length - 1), results: input.edits.map((edit, index) => ({ index, op: edit.op, applied: false, status: index < stepIndex ? 'rolled_back' : index === stepIndex ? 'failed' : 'not_applied' })) });
    } finally { draft.dispose(); }
  }

  async preview(slideId: string, options: { scale?: number } = {}) {
    const deck = this.snapshot();
    const slide = this.slide(this.handle, slideId);
    return renderPptxSlide(this.handle, deck.slides.findIndex(s => s.id === slide.id) + 1, options.scale);
  }

  override async render(page?: number, proposal?: string): Promise<RenderedPage> {
    this.alive();
    if (page === undefined) throw new DocumentToolError('UNSUPPORTED', 'Use pptx_preview with a slide ID, or render(1) with a one-based slide index.');
    if (!proposal) return renderPptxSlide(this.handle, page);
    const draft = openPresentation(await this.export(proposal));
    try { return await renderPptxSlide(draft, page); }
    finally { draft.dispose(); }
  }

  private applyOperation(handle: PresentationHandle, edit: PptxAgentEdit, changeBudget: number): Record<string, unknown> {
    if (edit.op === 'add_slide') {
      const deck = handle.snapshot();
      const index = integer(edit.index ?? deck.slides.length + 1, 1, deck.slides.length + 1, 'index');
      const layout = edit.layout ?? this.layouts[0]?.id;
      if (layout && !this.layouts.some(l => l.id === layout)) throw new DocumentToolError('UNKNOWN_LAYOUT', 'Copy a layout ID from pptx_outline.', { validLayouts: this.layouts.slice(0, 20) });
      const receipt = handle.insertSlide(index - 1, layout);
      return { changed: true, slideId: receipt.slideId, index, layout: layout ?? null };
    }
    if (edit.op === 'find_replace') {
      plainText(edit.query); plainText(edit.replacement);
      if (/[\r\n]/u.test(edit.query + edit.replacement)) throw new DocumentToolError('INVALID_ARGUMENT', 'find_replace stays within paragraphs. Use set_paragraphs to replace multiple paragraphs.');
      const scope = edit.slide ? this.slide(handle, edit.slide) : undefined;
      const read = unwrap(handle.readContent(scope ? { slideIds: [scope.id] } : undefined));
      const pattern = new RegExp(edit.query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), edit.caseSensitive ? 'gu' : 'giu');
      const steps: PptxEditStep[] = [];
      const changes = new Map<string, { slideId: string; shapeId: string; replacements: number }>();
      for (const story of read.stories) for (const paragraph of story.paragraphs) {
        pattern.lastIndex = 0;
        const text = story.text.slice(paragraph.start, paragraph.end);
        for (let match = pattern.exec(text); match; match = pattern.exec(text)) {
          if (steps.length === 128) throw new DocumentToolError('EDIT_LIMIT', 'More than 128 occurrences. Narrow query or supply a slide ID, then retry.');
          steps.push({ op: 'replaceText', target: { kind: 'range', slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId, start: paragraph.start + match.index, end: paragraph.start + match.index + match[0].length }, text: edit.replacement, expect: { text: match[0] } });
          const change = changes.get(story.shapeId) ?? { slideId: story.slideId, shapeId: story.shapeId, replacements: 0 };
          change.replacements++; changes.set(story.shapeId, change);
        }
      }
      if (!steps.length) return { changed: false, replacements: 0, changes: [] };
      const receipt = unwrap(handle.applyEdits({ expectVersion: read.version, steps, source: 'agent' }));
      const changed = this.bounded([...changes.values()], changeBudget);
      return { changed: receipt.applied, replacements: steps.length, changedShapes: changes.size, changes: changed, changesTruncated: changed.length < changes.size };
    }
    const slide = this.slide(handle, edit.slide);
    if (edit.op === 'replace_text' || edit.op === 'set_paragraphs') {
      const shape = this.shape(slide, edit.shape);
      const story = edit.story ? shape.textStories.find(story => story.id === edit.story) : shape.textStories.length === 1 ? shape.textStories[0] : undefined;
      if (!story) throw new DocumentToolError('UNKNOWN_STORY', 'Shape must have one text story, or supply a story ID from pptx_read_slide.', { validStories: shape.textStories.map(story => story.id).slice(0, 20) });
      const content = unwrap(handle.readContent({ slideIds: [slide.id] })).stories.find(item => item.storyId === story.id)!;
      if (content.paragraphs.some(p => p.fields.length || !p.editable)) throw new DocumentToolError('UNSUPPORTED_TARGET', 'Shape contains text fields. Use find_replace for plain text outside the fields.');
      const paragraphs = edit.op === 'replace_text' ? edit.text.replace(/\r\n?/g, '\n').split('\n').map(text => ({ runs: [{ text }], alignment: undefined, bullet: undefined })) : edit.paragraphs;
      const text = paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n');
      plainText(text);
      if (text.length > 16000) throw new DocumentToolError('EDIT_LIMIT', 'Shape text must fit in 16000 characters.');
      if (edit.op === 'replace_text' && text === content.text) return { changed: false, slideId: slide.id, shapeId: shape.id, storyId: story.id };
      const formatted = paragraphs.map((paragraph, index) => ({
        bullet: paragraph.bullet,
        alignment: paragraph.alignment ?? null,
        runs: paragraph.runs.flatMap(run => {
          const { text, ...style } = run;
          if ('fontFamily' in style && style.fontFamily) plainText(style.fontFamily);
          if (/[\r\n]/u.test(text)) throw new DocumentToolError('INVALID_ARGUMENT', 'Each run must stay in one paragraph. Add another paragraphs item for a line break.');
          return edit.op === 'replace_text' ? this.replacementRuns(story, index, text) : [{ text, style }];
        }),
      }));
      handle.setStoryParagraphs(story.id, formatted);
      return { changed: true, slideId: slide.id, shapeId: shape.id, storyId: story.id, before: this.summary(content.text), after: this.summary(text), paragraphs: paragraphs.length };
    }
    if (edit.op === 'duplicate_slide') {
      const count = handle.snapshot().slides.length;
      const index = integer(edit.index ?? handle.snapshot().slides.findIndex(s => s.id === slide.id) + 2, 1, count + 1, 'index');
      const receipt = handle.duplicateSlide(slide.id, index - 1);
      return { changed: true, sourceSlideId: slide.id, slideId: receipt.slideId, index };
    }
    if (edit.op === 'delete_slide') {
      if (handle.snapshot().slides.length === 1) throw new DocumentToolError('LAST_SLIDE', 'Add a slide before deleting the last slide.');
      const receipt = handle.deleteSlide(slide.id);
      return { changed: true, slideId: slide.id, oldIndex: receipt.fromIndex! + 1 };
    }
    if (edit.op === 'move_slide') {
      const index = integer(edit.index, 1, handle.snapshot().slides.length, 'index');
      const receipt = handle.moveSlide(slide.id, index - 1);
      return { changed: receipt.fromIndex !== receipt.toIndex, slideId: slide.id, oldIndex: receipt.fromIndex! + 1, index };
    }
    if (edit.op === 'set_notes') {
      plainText(edit.text);
      const before = slide.notes ?? '';
      if (before !== edit.text) handle.setSlideNotes(slide.id, edit.text);
      return { changed: before !== edit.text, slideId: slide.id, before: this.summary(before), after: this.summary(edit.text) };
    }
    plainText(edit.name);
    const rect = Object.fromEntries(Object.entries(edit.rect).map(([key, value]) => [key, Math.round(value * 12700)])) as unknown as ShapeRect;
    let receipt;
    if (edit.op === 'add_text_box') {
      plainText(edit.text);
      if (edit.style?.fontFamily) plainText(edit.style.fontFamily);
      if (/[\r\n]/u.test(edit.text)) throw new DocumentToolError('INVALID_ARGUMENT', 'add_text_box takes one paragraph. Use set_paragraphs on its returned shape ID for multiple paragraphs.');
      receipt = handle.addTextBox(slide.id, { name: edit.name, rect, text: edit.text, style: edit.style ?? {} });
    } else if (edit.op === 'add_shape') {
      receipt = handle.addShape(slide.id, { name: edit.name, rect, geometry: edit.geometry, fill: edit.fill });
    } else {
      const bytes = this.imageBytes(edit.base64, edit.contentType);
      receipt = handle.addPicture(slide.id, { name: edit.name, rect, contentType: edit.contentType, mediaBase64: Buffer.from(bytes).toString('base64') });
    }
    return { changed: true, slideId: slide.id, shapeId: receipt.shapeId, index: receipt.index + 1, rect: this.rect(rect) };
  }

  private replacementRuns(story: StorySnapshot, index: number, text: string) {
    const runs = (story.paragraphs[index] ?? story.paragraphs.at(-1))?.runs ?? [];
    if (!runs.length) return [{ text, style: {} }];
    const characters = Array.from(text);
    let offset = 0;
    return runs.map((run, index) => {
      const end = index + 1 === runs.length ? characters.length : Math.min(characters.length, offset + Array.from(run.text).length);
      const text = characters.slice(offset, end).join('');
      offset = end;
      const style = Object.fromEntries(Object.entries(run.style).filter(([, value]) => value !== null)) as TextStyle;
      return { text, style };
    });
  }
  private imageBytes(base64: string, contentType: string) {
    if (base64.length % 4 !== 0 || /[^A-Za-z0-9+/=]/u.test(base64)) throw new DocumentToolError('INVALID_IMAGE', 'Supply canonical base64 image bytes, without a data URL prefix.');
    const bytes = Buffer.from(base64, 'base64');
    if (!bytes.length || bytes.length > 8 * 1024 * 1024 || bytes.toString('base64') !== base64) throw new DocumentToolError('INVALID_IMAGE', 'Supply a canonical base64 image up to 8 MiB.');
    let info;
    try { info = imageSize(bytes); }
    catch { throw new DocumentToolError('INVALID_IMAGE', 'Image cannot be read. Supply PNG, JPEG, GIF, or WebP bytes.'); }
    const mime = { png: 'image/png', jpg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' };
    if (mime[info.type as keyof typeof mime] !== contentType) throw new DocumentToolError('INVALID_IMAGE', 'contentType must match the encoded image format.');
    if (![info.width, info.height].every(value => Number.isSafeInteger(value) && value > 0) || !Number.isSafeInteger(info.width * info.height) || info.width * info.height > 16000000) throw new DocumentToolError('INVALID_IMAGE', 'Image must fit in 16 megapixels. Resize it before adding.');
    return bytes;
  }
  private get layouts() { return this.cachedLayouts ??= this.handle.layouts().map(layout => ({ ...layout, name: layout.name?.slice(0, 200) ?? null, type: layout.type?.slice(0, 64) ?? null })); }
  private snapshot() { this.alive(); return this.handle.snapshot(); }
  private bounded<T>(items: T[], budget: number) {
    const result: T[] = [];
    let size = 0;
    for (const item of items) {
      const cost = JSON.stringify(item).length;
      if (result.length && size + cost > budget) break;
      result.push(item); size += cost;
    }
    return result;
  }
  private summary(text: string) { return { text: text ? textWindow(text, { length: 160 }).text : '', length: text.length, truncated: text.length > 160 }; }
  private rect(rect: ShapeRect) { return { emu: { x: rect.x, y: rect.y, width: rect.width, height: rect.height }, points: { x: rect.x / 12700, y: rect.y / 12700, width: rect.width / 12700, height: rect.height / 12700 } }; }
  private flatten(shapes: ShapeSnapshot[], parentId: string | null = null, depth = 0): Array<{ shape: ShapeSnapshot; parentId: string | null; depth: number }> {
    return shapes.flatMap(shape => [{ shape, parentId, depth }, ...this.flatten(shape.children, shape.id, depth + 1)]);
  }
  private shapeSummary(shape: ShapeSnapshot, parentId: string | null, depth: number) {
    const text = shape.textStories.map(story => story.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n')).join('\n');
    return { id: shape.id, type: shape.kind, name: shape.name.slice(0, 200), nameTruncated: shape.name.length > 200, geometry: shape.geometry.slice(0, 100), parentId, depth, childCount: shape.children.length, placeholder: shape.placeholder, position: this.rect(shape.inherited ?? shape), coordinateSpace: parentId ? 'group' : 'slide', text: this.summary(text) };
  }
  private slideSummary(slide: SlideSnapshot, index: number) {
    const shapes = this.flatten(slide.shapes).map(node => node.shape);
    const titleShape = shapes.find(shape => ['title', 'ctrTitle'].includes((shape.placeholder as { placeholderType?: string } | null)?.placeholderType ?? '')) ?? shapes.find(shape => shape.textStories.some(s => s.paragraphs.some(p => p.runs.some(r => r.text.trim()))));
    const title = titleShape?.textStories[0]?.paragraphs[0]?.runs.map(run => run.text).join('') ?? '';
    return { index: index + 1, id: slide.id, layoutId: slide.layoutPartPath, layoutName: this.layouts.find(layout => layout.id === slide.layoutPartPath)?.name ?? null, title: title.slice(0, 160), titleTruncated: title.length > 160 };
  }
  private slide(handle: PresentationHandle, id: string) {
    const slides = handle.snapshot().slides;
    const found = slides.find(slide => slide.id === id);
    if (!found) throw new DocumentToolError('UNKNOWN_SLIDE', 'Copy a slide ID from pptx_outline.', { validSlides: slides.slice(0, 20).map((slide, index) => ({ id: slide.id, index: index + 1 })) });
    return found;
  }
  private shape(slide: SlideSnapshot, id: string) {
    const shapes = this.flatten(slide.shapes).map(node => node.shape);
    const found = shapes.find(shape => shape.id === id);
    if (!found) throw new DocumentToolError('UNKNOWN_SHAPE', 'Copy a shape ID from pptx_outline or pptx_read_slide for this slide.', { slideId: slide.id, nearestShapeIds: shapes.sort((a, b) => this.distance(a.id, id) - this.distance(b.id, id)).slice(0, 5).map(shape => ({ id: shape.id, name: shape.name.slice(0, 100) })) });
    return found;
  }
  private distance(a: string, b: string) {
    let prefix = 0;
    while (prefix < Math.min(a.length, b.length) && a[prefix] === b[prefix]) prefix++;
    return a.length + b.length - 2 * prefix;
  }

  protected currentVersion() { return this.handle.version(); }
  protected validate(steps: PptxEditStep[], version: string) { unwrap(this.handle.validateEdits({ expectVersion: version, steps, source: 'agent' })); }
  protected apply(steps: PptxEditStep[], version: string) {
    const result = unwrap(this.handle.applyEdits({ expectVersion: version, steps, source: 'agent' }));
    return { applied: result.applied, version: result.version, changedSlides: result.changedSlides.length, changedStories: result.changedStories.length };
  }
  protected save(steps?: PptxEditStep[]) {
    if (!steps) return this.handle.save();
    const draft = this.handle.fork();
    try { unwrap(draft.applyEdits({ expectVersion: draft.version(), steps, source: 'agent' })); return draft.save(); }
    finally { draft.dispose(); }
  }
  protected reopen(bytes: Uint8Array) { const reopened = openPresentation(bytes); reopened.dispose(); }
  protected dispose() { this.refs.clear(); this.matches.clear(); if (this.owned) this.handle.dispose(); }

  private content() { this.alive(); return unwrap(this.handle.readContent()); }
  private reference(p: Paragraph) { return { ref: p.ref, slide: p.slide, slideId: p.story.slideId, shapeId: p.story.shapeId, story: p.story.storyId }; }
  private checkStory(story: string | undefined, stories: PptxStoryText[]) {
    if (story && !stories.some(s => s.storyId === story)) throw new DocumentToolError('UNKNOWN_STORY', 'Use a story from outline.');
  }
  private paragraphs(read: ReturnType<PptxAgentPresentation['content']>) {
    const active = new Set<string>();
    const result: Paragraph[] = [];
    for (const story of read.stories) {
      for (const paragraph of story.paragraphs) {
        const key = JSON.stringify([story.storyId, paragraph.paragraphId]);
        active.add(key);
        let ref = this.refs.get(key);
        if (!ref) { ref = `p${this.nextRef++}`; this.refs.set(key, ref); }
        result.push({ ref, slide: read.slides.findIndex(slide => slide.id === story.slideId) + 1, story, start: paragraph.start, end: paragraph.end, editable: paragraph.editable });
      }
    }
    for (const key of this.refs.keys()) if (!active.has(key)) this.refs.delete(key);
    return result;
  }
}

export async function openPptx(bytes: Uint8Array, options: PrototypeOptions = {}) {
  return PptxAgentPresentation.open(bytes, options);
}

export function attachPptx(handle: PresentationHandle, options: { name?: string } = {}) {
  return PptxAgentPresentation.attach(handle, options);
}
