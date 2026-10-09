import { initWasm, openPresentation, type PptxEditStep, type PptxStoryText, type PptxStoryTarget, type PresentationHandle } from '@betteroffice/pptx';
import { DocumentToolError, type GrepOptions, type TextEdit } from './types';
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

  private constructor(private readonly handle: PresentationHandle, private readonly source: Uint8Array, readonly name = 'presentation.pptx') { super(); }

  static async open(bytes: Uint8Array, options: PrototypeOptions = {}) {
    await initialize('pptx', initWasm, options.wasm);
    const source = bytes.slice();
    return new PptxAgentPresentation(openPresentation(source), source, options.name);
  }

  overview() {
    const read = this.content();
    return { name: this.name, format: this.format, version: read.version, slides: read.slides.length, paragraphs: this.paragraphs(read).length, capabilities: { grep: 'literal', proposals: 'paragraph-local text replacements', render: false, export: true, prototype: true }, workflow: 'outline lists slide text; grep for exact text; read its ref; propose with match and newText; review; verify; export. This prototype edits slide text, not notes, layout, or media.' };
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

  protected currentVersion() { return this.handle.version(); }
  protected validate(steps: PptxEditStep[], version: string) { unwrap(this.handle.validateEdits({ expectVersion: version, steps, source: 'agent' })); }
  protected apply(steps: PptxEditStep[], version: string) {
    const result = unwrap(this.handle.applyEdits({ expectVersion: version, steps, source: 'agent' }));
    return { applied: result.applied, version: result.version, changedSlides: result.changedSlides.length, changedStories: result.changedStories.length };
  }
  protected save(steps?: PptxEditStep[]) {
    if (!steps) return this.handle.save();
    const draft = openPresentation(this.source, { initialUpdate: this.handle.encodeStateAsUpdate() });
    try { unwrap(draft.applyEdits({ expectVersion: draft.version(), steps, source: 'agent' })); return draft.save(); }
    finally { draft.dispose(); }
  }
  protected reopen(bytes: Uint8Array) { const reopened = openPresentation(bytes); reopened.dispose(); }
  protected dispose() { this.refs.clear(); this.matches.clear(); this.handle.dispose(); }

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
