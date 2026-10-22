import { createYrsSession, saveYrsDocx, type DocxParagraphTarget, type YrsSession, type YrsStorySegment } from '@betteroffice/docx/yrs';
import { readCommentPart, verifyCommentOutput, type CommentSource } from './comment-output';
import { cellPosition, commentBody, commentRanges, commentRoot, commentText, continueList, deleteParagraph, deleteTableRow, hostEdit, insertParagraphsWithAnchorFormat, insertTableRow, lineText, listLevel, notInTable, requireParagraph, resolveStyle, tableCells, type ParagraphStyle } from './docx-edits';
import { docxBatchSchema, type DocxAgentBatch, type DocxAgentEdit, type DocxEditOutcome } from './docx-schema';
import {
  DocumentToolError,
  type DocumentOptions,
  type ExactTextEdit,
  type GrepMatch,
  type GrepOptions,
  type ParagraphRead,
  type ParagraphRef,
  type Proposal,
  type ProposalChange,
  type TableCellPosition,
  type TextEdit,
  type TextRun,
} from './types';

interface Paragraph extends ParagraphRef {
  paraId: string;
  text: string;
  style: string | null;
  headingLevel: number | null;
  runs: TextRun[];
  fingerprint: string;
  table?: TableCellPosition;
  list?: { level: number };
}

interface OutlinePage<T> {
  version: number;
  items: T[];
  nextOffset: number | null;
  total: number;
}

interface CommentSummary {
  id: string;
  parentId: string | null;
  done: boolean;
  author: string;
  date: string;
  text: string;
  length: number;
  anchor: string | null;
}

type ParagraphSummary = ParagraphRef & {
  style: string | null;
  headingLevel: number | null;
  text: string;
  length: number;
  table?: TableCellPosition;
  list?: { level: number };
};

type EditStep = (draft: YrsSession) => { result: Omit<DocxEditOutcome, 'index' | 'op' | 'refs'>; created?: DocxParagraphTarget[] };

interface HeldProposal {
  value: Proposal;
  targets: Map<string, { story: string; paraId: string; fingerprint: string }>;
}

const MAX_TEXT = 16_000;

function integer(value: number, min: number, max: number, name: string): number {
  if (!Number.isSafeInteger(value) || value < min || value > max) {
    throw new DocumentToolError('INVALID_ARGUMENT', `${name} must be an integer from ${min} to ${max}.`);
  }
  return value;
}

function boundary(text: string, offset: number): boolean {
  return offset === 0 || offset === text.length ||
    !(text.charCodeAt(offset - 1) >= 0xd800 && text.charCodeAt(offset - 1) <= 0xdbff &&
      text.charCodeAt(offset) >= 0xdc00 && text.charCodeAt(offset) <= 0xdfff);
}

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

function graphemeCuts(text: string): Set<number> {
  const cuts = new Set([text.length]);
  for (const { index } of graphemes.segment(text)) cuts.add(index);
  return cuts;
}

function trimChange(change: ProposalChange): ProposalChange {
  const { oldText, newText } = change;
  let left = 0;
  while (left < oldText.length && left < newText.length && oldText[left] === newText[left]) left++;
  const oldCuts = graphemeCuts(oldText), newCuts = graphemeCuts(newText);
  while (left > 0 && !(oldCuts.has(left) && newCuts.has(left))) left--;
  let right = 0;
  while (right < oldText.length - left && right < newText.length - left && oldText[oldText.length - right - 1] === newText[newText.length - right - 1]) right++;
  while (right > 0 && !(oldCuts.has(oldText.length - right) && newCuts.has(newText.length - right))) right--;
  return { ...change, start: change.start + left, oldText: oldText.slice(left, oldText.length - right), newText: newText.slice(left, newText.length - right) };
}

function visibleFormatting(attributes: Record<string, unknown>): Record<string, unknown> {
  const keys = new Set(['bold', 'italic', 'underline', 'strike', 'textColor', 'color', 'fontFamily', 'fontFamilyAscii', 'fontFamilyHAnsi', 'highlight', 'verticalAlign', 'hyperlink', 'ins', 'del']);
  const formatting = Object.fromEntries(Object.entries(attributes).filter(([key]) => keys.has(key)).map(([key, value]) =>
    [key, JSON.stringify(value).length <= 500 ? value : '(formatting value exceeds preview limit)']));
  const fontSize = attributes.fontSize;
  if (fontSize && typeof fontSize === 'object') {
    const size = fontSize as { size?: unknown; sizeCs?: unknown };
    if (typeof size.size === 'number') formatting.fontSizePt = size.size / 2;
    if (typeof size.sizeCs === 'number') formatting.complexScriptFontSizePt = size.sizeCs / 2;
  }
  return formatting;
}

function literal(query: string, caseSensitive: boolean): RegExp {
  return new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseSensitive ? 'gu' : 'giu');
}

export class DocxAgentDocument {
  readonly name: string;
  private epoch = 0;
  private indexedEpoch = -1;
  private paragraphs: Paragraph[] = [];
  private readonly identities = new Map<string, Paragraph>();
  private readonly proposals = new Map<string, HeldProposal>();
  private readonly matches = new Map<string, Omit<ExactTextEdit, 'newText'>>();
  private readonly cursors = new Map<string, { epoch: number; signature: string; paragraph: number; offset: number }>();
  private readonly unsubscribe: () => void;
  private nextRef = 1;
  private nextProposal = 1;
  private nextCursor = 1;
  private nextMatch = 1;
  private closed = false;
  private paragraphStyles?: ParagraphStyle[];
  private commentPartCheck?: { source: ArrayBuffer; check: Promise<CommentSource> };
  private resolvedComments = new Map<string, boolean>();

  constructor(
    readonly session: YrsSession,
    private readonly options: DocumentOptions = {},
    private readonly ownsSession = false,
  ) {
    this.name = options.name ?? 'document.docx';
    this.unsubscribe = session.onUpdate(() => { this.epoch++; });
  }

  overview(options: { offset?: number; limit?: number } = {}) {
    this.index();
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 30, 1, 100, 'limit');
    const stories = new Map<string, number>();
    for (const p of this.paragraphs) stories.set(p.story, (stories.get(p.story) ?? 0) + 1);
    const entries = [...stories].slice(offset, offset + limit).map(([story, paragraphs]) => ({ story, paragraphs }));
    return {
      name: this.name,
      format: 'docx' as const,
      version: this.epoch,
      paragraphs: this.paragraphs.length,
      stories: entries,
      nextOffset: offset + entries.length < stories.size ? offset + entries.length : null,
      capabilities: { grep: 'literal', proposals: true, atomicEdits: true, render: !!this.options.renderer, export: !!this.session.materializeDocx()?.originalBuffer },
      workflow: 'grep for the exact text to replace; read its ref for context; propose with match and newText; review; render; verify; export. For styles, bold/italic/underline, paragraphs, table rows and comments, use docx_edit with the outline version.',
    };
  }

  list(options: { story: 'comments'; offset?: number; limit?: number; headingsOnly?: boolean }): OutlinePage<CommentSummary>;
  list(options?: { story?: string; offset?: number; limit?: number; headingsOnly?: boolean }): OutlinePage<ParagraphSummary>;
  list(options: { story?: string; offset?: number; limit?: number; headingsOnly?: boolean } = {}) {
    return options.story === 'comments' ? this.commentOutline(options) : this.paragraphOutline(options);
  }

  private commentOutline(options: { offset?: number; limit?: number }) {
    this.index();
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 30, 1, 100, 'limit');
    const comments = this.session.listComments();
    return {
      version: this.epoch,
      items: comments.slice(offset, offset + limit).map(comment => {
        const text = commentText(comment.body).replace(/\n$/, '');
        let ref: string | undefined;
        try {
          const [anchor] = this.session.resolveComment(comment.id);
          if (anchor) {
            const paragraph = this.paragraphs.find(p => {
              if (p.story !== anchor.story) return false;
              const span = this.session.locateParagraph(p.story, p.paraId);
              return span.start <= anchor.start && anchor.start <= span.end;
            });
            ref = paragraph?.ref;
          }
        } catch {}
        return { id: comment.id, parentId: comment.parentId, done: comment.done,
          author: comment.author, date: comment.date, text: text.slice(0, 120), length: text.length, anchor: ref ?? null };
      }),
      nextOffset: offset + limit < comments.length ? offset + limit : null,
      total: comments.length,
    };
  }

  private paragraphOutline(options: { story?: string; offset?: number; limit?: number; headingsOnly?: boolean } = {}) {
    this.index();
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 30, 1, 100, 'limit');
    this.checkStory(options.story);
    const entries = this.paragraphs.filter(p => (!options.story || p.story === options.story) &&
      (!options.headingsOnly || p.headingLevel !== null));
    return {
      version: this.epoch,
      items: entries.slice(offset, offset + limit).map(p => ({ ...this.reference(p), style: p.style, headingLevel: p.headingLevel, text: p.text.slice(0, 120), length: p.text.length, ...(p.table && { table: { ...p.table } }), ...(p.list && { list: { ...p.list } }) })),
      nextOffset: offset + limit < entries.length ? offset + limit : null,
      total: entries.length,
    };
  }

  grep(options: GrepOptions) {
    this.index();
    if (typeof options.query !== 'string' || !options.query || options.query.length > 1000 || /[\r\n\ufffc]/u.test(options.query)) {
      throw new DocumentToolError('INVALID_ARGUMENT', 'query must be 1–1000 characters of literal text within a paragraph.');
    }
    const limit = integer(options.limit ?? 20, 1, 100, 'limit');
    this.checkStory(options.story);
    const signature = JSON.stringify([options.query, options.caseSensitive ?? false, options.story ?? null]);
    const cursor = options.cursor ? this.cursors.get(options.cursor) : undefined;
    if (options.cursor && (!cursor || cursor.epoch !== this.epoch || cursor.signature !== signature)) {
      throw new DocumentToolError('STALE_CURSOR', 'Search changed or the document was edited. Repeat grep without cursor.');
    }
    const matches: GrepMatch[] = [];
    const pattern = literal(options.query, options.caseSensitive ?? false);
    let size = 0;
    for (let i = cursor?.paragraph ?? 0; i < this.paragraphs.length; i++) {
      const p = this.paragraphs[i];
      if (options.story && p.story !== options.story) continue;
      pattern.lastIndex = i === cursor?.paragraph ? cursor.offset : 0;
      for (let match = pattern.exec(p.text); match; match = pattern.exec(p.text)) {
        const start = match.index;
        const end = start + match[0].length;
        if (p.runs.some(run => run.protected && run.start < end && run.end > start)) continue;
        if (matches.length === limit || (matches.length > 0 && size + match[0].length + 160 > MAX_TEXT)) {
          const nextCursor = `c${this.nextCursor++}`;
          if (this.cursors.size >= 64) this.cursors.delete(this.cursors.keys().next().value!);
          this.cursors.set(nextCursor, { epoch: this.epoch, signature, paragraph: i, offset: start });
          return { matches, nextCursor, version: this.epoch };
        }
        const contextStart = Math.max(0, start - 80);
        const context = p.text.slice(contextStart, end + 80);
        size += context.length;
        const id = `m${this.nextMatch++}`;
        if (this.matches.size >= 1024) this.matches.delete(this.matches.keys().next().value!);
        this.matches.set(id, { ref: p.ref, revision: p.revision, oldText: match[0], start });
        matches.push({ ...this.reference(p), match: id, start, end, text: match[0], context, contextStart });
      }
    }
    return { matches, nextCursor: null, version: this.epoch };
  }

  read(ref: string, options: { start?: number; length?: number } = {}): ParagraphRead {
    const p = this.get(ref);
    let start = integer(options.start ?? 0, 0, p.text.length, 'start');
    const length = integer(options.length ?? 4000, 1, MAX_TEXT, 'length');
    if (!boundary(p.text, start)) start--;
    let end = Math.min(p.text.length, start + length);
    const visibleRuns = p.runs.filter(run => run.start < end && run.end > start);
    if (visibleRuns.length > 100) end = Math.min(end, visibleRuns[99].end);
    if (!boundary(p.text, end)) end--;
    if (end === start && end < p.text.length) end += 2;
    return {
      ...this.reference(p), text: p.text.slice(start, end), start, end,
      length: p.text.length, nextStart: end < p.text.length ? end : null, style: p.style,
      runs: p.runs.filter(run => run.start < end && run.end > start).map(run => structuredClone({ ...run, start: Math.max(start, run.start), end: Math.min(end, run.end) })),
      ...(p.table && { table: { ...p.table } }),
    };
  }

  /** Applies 1–32 direct edits on a fork, then commits them as one undo step; any failure applies none. */
  async edit(input: DocxAgentBatch) {
    this.index();
    const parsed = docxBatchSchema.safeParse(input);
    if (!parsed.success) throw new DocumentToolError('INVALID_ARGUMENT', 'Use the docx_edit schema: version and 1–32 edits.', { issues: parsed.error.issues.map(issue => ({ path: issue.path, message: issue.message })) });
    const { version, edits } = parsed.data;
    if (version !== this.epoch) throw new DocumentToolError('STALE_VERSION', 'The document changed. Read office_outline or office_grep again and copy its version.');
    if (JSON.stringify(edits).length > 64_000) throw new DocumentToolError('EDIT_LIMIT', 'A batch holds at most 64000 characters of input. Split it.');
    const outcomes = await this.commit(version, edits);
    this.index();
    return {
      applied: true as const, version: this.epoch,
      results: outcomes.map(({ result, created }, index): DocxEditOutcome => ({
        index, op: edits[index].op, ...result,
        ...(created && { refs: created.map(({ story, paraId }) => this.identities.get(JSON.stringify([story, paraId]))?.ref ?? null) }),
      })),
    };
  }

  propose(input: { author: string; note?: string; edits: TextEdit[] }): Proposal {
    this.index();
    if (!input.author || input.author.length > 200 || (input.note?.length ?? 0) > 2000 || !Array.isArray(input.edits) || input.edits.length < 1 || input.edits.length > 32) {
      throw new DocumentToolError('INVALID_ARGUMENT', 'Supply an author (1–200 characters), note (up to 2000), and 1–32 edits.');
    }
    if (this.proposals.size >= 64) {
      const closed = [...this.proposals].find(([, held]) => held.value.status !== 'pending');
      if (closed) this.proposals.delete(closed[0]);
      else throw new DocumentToolError('PROPOSAL_LIMIT', 'Reject or discard a pending proposal before staging more than 64.');
    }
    const changes: ProposalChange[] = [];
    const targets: HeldProposal['targets'] = new Map();
    let size = 0;
    for (const inputEdit of input.edits) {
      let edit: ExactTextEdit;
      if ('oldText' in inputEdit) edit = inputEdit;
      else {
        const target = this.matches.get(inputEdit.match);
        if (!target) throw new DocumentToolError('UNKNOWN_MATCH', 'Match expired or belongs to another document. Run grep again and copy its match ID.');
        edit = { ...target, newText: inputEdit.newText };
      }
      const p = this.get(edit.ref);
      if (p.revision !== edit.revision) this.stale([p.ref]);
      if (typeof edit.oldText !== 'string' || typeof edit.newText !== 'string' || /[\r\n\ufffc\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(edit.newText)) {
        throw new DocumentToolError('INVALID_ARGUMENT', 'Edits replace plain text inside one paragraph; paragraph breaks, embeds, and control characters are unsupported.');
      }
      if (/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(edit.newText)) {
        throw new DocumentToolError('INVALID_ARGUMENT', 'newText must contain complete Unicode characters.');
      }
      size += edit.oldText.length + edit.newText.length;
      if (size > MAX_TEXT) throw new DocumentToolError('EDIT_LIMIT', `Combined oldText and newText must fit in ${MAX_TEXT} UTF-16 units.`);
      let start = edit.start;
      if (start === undefined) {
        if (!edit.oldText) throw new DocumentToolError('INVALID_ARGUMENT', 'Insertion requires start from read or grep.');
        start = p.text.indexOf(edit.oldText);
        if (start >= 0 && p.text.indexOf(edit.oldText, start + 1) >= 0) {
          throw new DocumentToolError('AMBIGUOUS_TEXT', 'oldText appears more than once. Supply start from grep.', { ref: p.ref });
        }
      }
      if (start < 0 || !Number.isSafeInteger(start) || start + edit.oldText.length > p.text.length || p.text.slice(start, start + edit.oldText.length) !== edit.oldText) {
        throw new DocumentToolError('TEXT_MISMATCH', 'oldText does not match at start. Run grep and use its match ID instead of calculating offsets.', { ref: p.ref });
      }
      const end = start + edit.oldText.length;
      if (!boundary(p.text, start) || !boundary(p.text, end)) throw new DocumentToolError('INVALID_RANGE', 'Text ranges must not split a Unicode character.');
      if (edit.oldText === edit.newText) throw new DocumentToolError('NO_CHANGE', 'oldText and newText are identical.');
      if (p.runs.some(run => run.protected && (run.start < end && run.end > start || start === end && run.start <= start && start < run.end))) {
        throw new DocumentToolError('PROTECTED_CONTENT', 'This range contains an embed or an existing tracked change. Choose an ordinary text range.');
      }
      if (changes.some(change => change.ref === p.ref && start <= change.start + change.oldText.length && end >= change.start)) {
        throw new DocumentToolError('OVERLAPPING_EDITS', 'Edits within one proposal must have separate, non-touching ranges.');
      }
      changes.push({ ...this.reference(p), start, oldText: edit.oldText, newText: edit.newText });
      targets.set(p.ref, { story: p.story, paraId: p.paraId, fingerprint: p.fingerprint });
    }
    const value: Proposal = { id: `proposal${this.nextProposal++}`, author: input.author, note: input.note ?? '', status: 'pending', changes, staleRefs: [] };
    this.proposals.set(value.id, { value, targets });
    return structuredClone(value);
  }

  review(id: string): Proposal {
    const held = this.proposal(id);
    return structuredClone({ ...held.value, staleRefs: held.value.status === 'pending' ? this.staleRefs(held) : [] });
  }

  listProposals() {
    this.index();
    return [...this.proposals.values()].map(({ value }) => ({ id: value.id, author: value.author, note: value.note.slice(0, 120), status: value.status, changes: value.changes.length }));
  }

  async accept(id: string, options: { tracked?: boolean; date?: string } = {}) {
    const held = this.pending(id);
    if (options.tracked && (!options.date || !Number.isFinite(Date.parse(options.date)))) {
      throw new DocumentToolError('INVALID_ARGUMENT', 'Tracked acceptance requires an ISO date.');
    }
    const draft = await this.fork();
    try {
      this.checkPending(held);
      const vector = draft.encodeStateVector();
      this.apply(draft, held, options.tracked ? { name: held.value.author, date: options.date! } : undefined);
      this.checkPending(held);
      this.session.addUndoBoundary();
      try { this.session.applyLocalUpdate(draft.encodeStateAsUpdate(vector)); }
      finally { this.session.addUndoBoundary(); }
      held.value.status = 'accepted';
      return { id, status: 'accepted' as const, tracked: options.tracked ?? false, changes: held.value.changes.length };
    } finally { draft.destroy(); }
  }

  reject(id: string) {
    const held = this.proposal(id);
    if (held.value.status !== 'pending') throw new DocumentToolError('PROPOSAL_CLOSED', `Proposal is ${held.value.status}.`);
    held.value.status = 'rejected';
    return { id, status: 'rejected' as const };
  }

  discard(id: string) {
    this.proposal(id);
    this.proposals.delete(id);
  }

  async render(page = 1, proposalId?: string) {
    integer(page, 1, 100_000, 'page');
    if (!this.options.renderer) throw new DocumentToolError('UNSUPPORTED', 'No renderer configured. Supply a DocumentRenderer or use the MCP server.');
    const held = proposalId ? this.pending(proposalId) : undefined;
    const draft = await this.fork();
    try {
      if (held) { this.checkPending(held); this.apply(draft, held); }
      return await this.options.renderer(draft, page);
    } finally { draft.destroy(); }
  }

  async export(proposalId?: string): Promise<Uint8Array> {
    const held = proposalId ? this.pending(proposalId) : undefined;
    const draft = await this.fork();
    try {
      if (held) { this.checkPending(held); this.apply(draft, held); }
      if (!draft.materializeDocx()?.originalBuffer) throw new DocumentToolError('UNSUPPORTED', 'Export requires a session opened from DOCX bytes.');
      return (await saveYrsDocx(draft)).bytes;
    } finally { draft.destroy(); }
  }

  async verify(proposalId?: string) {
    const bytes = await this.export(proposalId);
    const reopened = await openDocx(bytes);
    try {
      return { reopened: true, bytes: bytes.length, paragraphs: reopened.overview().paragraphs, checks: ['save', 'reopen'], unchecked: ['visual fidelity', 'semantic correctness'] };
    } finally { reopened.close(); }
  }

  close() {
    if (this.closed) return;
    this.closed = true;
    this.unsubscribe();
    this.proposals.clear();
    this.cursors.clear();
    this.matches.clear();
    this.identities.clear();
    this.paragraphs = [];
    if (this.ownsSession) this.session.destroy();
  }

  private index() {
    if (this.closed) throw new DocumentToolError('CLOSED', 'This document is closed. Open it again.');
    if (this.indexedEpoch === this.epoch) return;
    const paragraphs: Paragraph[] = [];
    const active = new Set<string>();
    const cells = new Map<string, TableCellPosition>();
    const stories = this.session.storyIds();
    const segmentsByStory = new Map(stories.map(story => [story, this.session.storySegments(story)]));
    for (const [story, segments] of segmentsByStory) {
      for (const [cell, position] of tableCells(story, segments)) cells.set(cell, cellPosition(position));
    }
    const position = (story: string): { root: string; path: number[] } => {
      const cell = cells.get(story);
      if (!cell) return { root: story, path: [] };
      const parent = position(cell.story);
      return { root: parent.root, path: [...parent.path, cell.index, cell.row, cell.column] };
    };
    const positions = new Map(stories.map(story => [story, position(story)]));
    stories.sort((a, b) => {
      const left = positions.get(a)!;
      const right = positions.get(b)!;
      if (left.root !== right.root) return left.root < right.root ? -1 : 1;
      for (let i = 0; i < Math.min(left.path.length, right.path.length); i++) {
        if (left.path[i] !== right.path[i]) return left.path[i] - right.path[i];
      }
      return left.path.length - right.path.length;
    });
    for (const story of stories) {
      const headings = new Map(this.session.headings(story).map(entry => [entry.paraId, entry.heading.outlineLevel + 1]));
      let text = '';
      let runs: TextRun[] = [];
      let segments: YrsStorySegment[] = [];
      let ordinal = 0;
      const storySegments = segmentsByStory.get(story)!;
      for (const segment of storySegments) {
        segments.push(segment);
        if (segment.kind === 'pilcrow') {
          const key = JSON.stringify([story, segment.paraId]);
          active.add(key);
          const previous = this.identities.get(key);
          const fingerprint = JSON.stringify(segments);
          const level = listLevel(segment.properties);
          const p: Paragraph = {
            ref: previous?.ref ?? `p${this.nextRef++}`,
            revision: previous ? previous.revision + Number(previous.fingerprint !== fingerprint) : 1,
            story, paraId: segment.paraId, paragraph: ++ordinal, text, runs, fingerprint,
            style: typeof segment.properties.pStyle === 'string' ? segment.properties.pStyle : null,
            headingLevel: headings.get(segment.paraId) ?? null,
            ...(level !== undefined && { list: { level } }),
          };
          this.identities.set(key, p);
          paragraphs.push(p);
          text = ''; runs = []; segments = [];
        } else {
          const value = segment.kind === 'text' ? segment.text : '\ufffc';
          runs.push({ start: text.length, end: text.length + value.length, formatting: visibleFormatting(segment.attributes), protected: segment.kind === 'embed' || !!segment.attributes.ins || !!segment.attributes.del });
          text += value;
        }
      }
    }
    for (const key of this.identities.keys()) if (!active.has(key)) this.identities.delete(key);
    for (const p of paragraphs) p.table = cells.get(p.story);
    this.paragraphs = paragraphs;
    this.indexedEpoch = this.epoch;
  }

  private reference(p: Paragraph): ParagraphRef {
    return { ref: p.ref, revision: p.revision, story: p.story, paragraph: p.paragraph };
  }

  private get(ref: string): Paragraph {
    this.index();
    const p = this.paragraphs.find(p => p.ref === ref);
    if (!p) throw new DocumentToolError('UNKNOWN_REF', 'Paragraph missing. Use list or grep to find a current ref.', { ref });
    return p;
  }

  private checkStory(story?: string) {
    if (story && !this.paragraphs.some(p => p.story === story)) throw new DocumentToolError('UNKNOWN_STORY', 'Story missing. Use overview to list stories.', { story });
  }

  private proposal(id: string): HeldProposal {
    this.index();
    const held = this.proposals.get(id);
    if (!held) throw new DocumentToolError('UNKNOWN_PROPOSAL', 'Proposal missing. Stage a proposal first.', { id });
    return held;
  }

  private pending(id: string): HeldProposal {
    const held = this.proposal(id);
    this.checkPending(held);
    return held;
  }

  private checkPending(held: HeldProposal) {
    if (held.value.status !== 'pending') throw new DocumentToolError('PROPOSAL_CLOSED', `Proposal is ${held.value.status}. Stage a new proposal.`);
    const stale = this.staleRefs(held);
    if (stale.length) this.stale(stale);
  }

  private staleRefs(held: HeldProposal): string[] {
    this.index();
    return [...held.targets].filter(([ref, target]) => this.paragraphs.find(p => p.ref === ref)?.fingerprint !== target.fingerprint).map(([ref]) => ref);
  }

  private stale(refs: string[]): never {
    throw new DocumentToolError('STALE_TARGET', 'Target text or formatting changed. Read these refs and propose again.', { refs });
  }

  private async fork(): Promise<YrsSession> {
    this.index();
    const draft = await createYrsSession();
    try {
      const source = this.session.materializeDocx()?.originalBuffer;
      if (source) draft.openDocx(new Uint8Array(source), false);
      draft.loadState(this.session.encodeState());
      return draft;
    } catch (error) { draft.destroy(); throw error; }
  }

  private apply(draft: YrsSession, held: HeldProposal, author?: { name: string; date: string }) {
    const changes = held.value.changes.map(trimChange).sort((a, b) => a.ref.localeCompare(b.ref) || b.start - a.start);
    for (const change of changes) {
      if (!change.oldText && !change.newText) continue;
      const { story, paraId } = held.targets.get(change.ref)!;
      draft.replaceRange({ story, start: { paraId, offset: change.start }, end: { paraId, offset: change.start + change.oldText.length } }, change.newText, author);
    }
  }

  private async commit(version: number, edits: DocxAgentEdit[]) {
    let failed: number | undefined;
    try {
      const steps: EditStep[] = [];
      for (const [index, edit] of edits.entries()) { failed = index; steps.push(await this.plan(edit)); }
      failed = undefined;
      const draft = await this.fork();
      try {
        if (this.epoch !== version) throw new DocumentToolError('STALE_VERSION', 'The document changed while the batch started. Read the version again.');
        const vector = draft.encodeStateVector();
        const outcomes: Array<ReturnType<EditStep>> = [];
        let resolved: Map<string, boolean> | undefined;
        const doneBefore = edits.some(edit => edit.op === 'resolve_comment') ? new Map(draft.listComments().map(comment => [comment.id, comment.done])) : undefined;
        for (const [index, step] of steps.entries()) { failed = index; outcomes.push(step(draft)); }
        failed = undefined;
        if (edits.some(edit => edit.op === 'add_comment' || edit.op === 'reply_comment' || edit.op === 'resolve_comment')) {
          failed = edits.findIndex(edit => edit.op === 'add_comment' || edit.op === 'reply_comment' || edit.op === 'resolve_comment');
          try {
            const source = await this.checkCommentPart();
            if (!source) throw new DocumentToolError('UNSUPPORTED_TARGET', 'Comment export requires a session opened from DOCX bytes.');
            const seeded = new Set((draft.materializeDocx()?.package.document.comments ?? []).map(comment => String(comment.id)));
            const comments = draft.listComments();
            const batch = new Set(edits.flatMap((edit, index) => edit.op === 'resolve_comment' ? [outcomes[index].result.comment!] : []));
            resolved = new Map(this.resolvedComments);
            for (const id of batch) if (!resolved.has(id)) resolved.set(id, doneBefore?.get(id) ?? false);
            const earlier = this.resolvedComments;
            await verifyCommentOutput(source, (await saveYrsDocx(draft)).bytes, comments.filter(comment => !seeded.has(comment.id)).length,
              comments.map(comment => ({ id: comment.id, parentId: comment.parentId, ...(batch.has(comment.id) || (earlier.get(comment.id) ?? comment.done) !== comment.done ? { done: comment.done } : {}) })));
          } catch (error) {
            if (error instanceof DocumentToolError) throw error;
            throw new DocumentToolError('UNSUPPORTED_TARGET', 'The saved comment package could not be verified.');
          }
          if (this.epoch !== version) throw new DocumentToolError('STALE_VERSION', 'The document changed while the comment export was verified. Read the version again.');
        }
        failed = undefined;
        this.session.addUndoBoundary();
        try { this.session.applyLocalUpdate(draft.encodeStateAsUpdate(vector)); }
        finally { this.session.addUndoBoundary(); }
        if (resolved) this.resolvedComments = resolved;
        return outcomes;
      } finally { draft.destroy(); }
    } catch (error) {
      const failure = error instanceof DocumentToolError ? error : new DocumentToolError('INVALID_EDIT', error instanceof Error ? error.message : String(error));
      throw new DocumentToolError(failure.code, `${failure.message} No batch edits were applied.`, { ...failure.details, ...(failed === undefined ? {} : { failedIndex: failed }) });
    }
  }

  private async plan(edit: DocxAgentEdit): Promise<EditStep> {
    if (edit.op === 'reply_comment' || edit.op === 'resolve_comment') {
      const body = edit.op === 'reply_comment' ? commentBody(edit.author, edit.text) : undefined;
      const date = edit.op === 'reply_comment' ? edit.date ?? new Date().toISOString().replace(/\.\d+Z$/, 'Z') : undefined;
      await this.checkCommentPart();
      return draft => {
        const root = commentRoot(draft, edit.comment);
        if (edit.op === 'resolve_comment') {
          const done = edit.done ?? true;
          draft.updateComment(root.id, { done });
          return { result: { comment: root.id, parentId: null, done } };
        }
        const { commentId } = draft.addComment(commentRanges(draft, root.id), edit.author, date!, body);
        draft.updateComment(commentId, { parentId: root.id });
        return { result: { comment: commentId, parentId: root.id, done: false } };
      };
    }
    if (edit.op === 'format_text' || edit.op === 'add_comment') {
      const held = this.matches.get(edit.match);
      if (!held) throw new DocumentToolError('UNKNOWN_MATCH', 'Match expired or belongs to another document. Run grep again and copy its match ID.');
      const p = this.get(held.ref);
      if (p.revision !== held.revision) throw new DocumentToolError('STALE_TARGET', 'The matched paragraph changed. Grep again and use a fresh match ID.', { refs: [p.ref] });
      const start = held.start ?? 0;
      const end = start + held.oldText.length;
      const range = { story: p.story, start: { paraId: p.paraId, offset: start }, end: { paraId: p.paraId, offset: end } };
      if (edit.op === 'add_comment') {
        const stories = new Map(this.session.listStories().stories.map(story => [story.story, story]));
        let story = stories.get(p.story);
        while (story?.kind === 'tableCell' && story.parent) story = stories.get(story.parent);
        if (story?.story !== 'body' || story.kind !== 'body') throw new DocumentToolError('UNSUPPORTED_TARGET', 'Comments anchor in the main body text and its tables, not inside content controls, headers, footers, notes or text boxes.');
        const body = commentBody(edit.author, edit.text);
        await this.checkCommentPart();
        const date = edit.date ?? new Date().toISOString().replace(/\.\d+Z$/, 'Z');
        return draft => {
          requireParagraph(draft, p);
          return { result: { ref: p.ref, start, end, comment: draft.addComment([range], edit.author, date, body).commentId } };
        };
      }
      const delta = { bold: edit.bold, italic: edit.italic, underline: edit.underline };
      if (Object.values(delta).every(value => value === undefined)) throw new DocumentToolError('INVALID_ARGUMENT', 'format_text needs bold, italic or underline.');
      return draft => {
        requireParagraph(draft, p);
        draft.formatRange(range, delta);
        return { result: { ref: p.ref, start, end } };
      };
    }
    const p = this.get(edit.op === 'insert_paragraphs' ? edit.after : edit.op === 'set_paragraph_style' || edit.op === 'delete_paragraph' ? edit.paragraph : edit.row);
    const target = { story: p.story, paraId: p.paraId };
    switch (edit.op) {
      case 'set_paragraph_style': {
        const styleId = resolveStyle(this.styles(), edit.style);
        return draft => {
          requireParagraph(draft, target);
          hostEdit(draft, { op: 'setParagraphStyle', target, styleId });
          return { result: { ref: p.ref, style: styleId } };
        };
      }
      case 'insert_paragraphs': {
        const fallback = this.styles().find(style => style.isDefault)?.id;
        if (edit.list && edit.runFormat !== undefined) {
          throw new DocumentToolError('INVALID_ARGUMENT', 'runFormat cannot be combined with list: continue. Omit runFormat.');
        }
        if (edit.list && edit.paragraphs.some(({ style }) => style)) {
          throw new DocumentToolError('INVALID_ARGUMENT', 'list: continue takes the style of the anchor list item. Omit style.');
        }
        const paragraphs = edit.paragraphs.map(({ text, style }) => {
          lineText(text);
          return { text, styleId: style ? resolveStyle(this.styles(), style) : fallback };
        });
        if (edit.list) {
          const stories = new Map(this.session.listStories().stories.map(story => [story.story, story]));
          for (let story = stories.get(p.story); story; story = story.parent ? stories.get(story.parent) : undefined) {
            if (story.kind === 'contentControl') throw new DocumentToolError('UNSUPPORTED_TARGET', 'Lists inside content controls cannot be continued yet. Use a list item outside the content control.');
          }
        }
        return draft => {
          requireParagraph(draft, target);
          if (edit.list) return { result: {}, created: continueList(draft, target, paragraphs) };
          if (edit.runFormat === 'anchor') return { result: {}, created: insertParagraphsWithAnchorFormat(draft, target, paragraphs) };
          return { result: {}, created: hostEdit(draft, { op: 'insertParagraphs', target, at: 'end', paragraphs }).newParagraphs };
        };
      }
      case 'delete_paragraph':
        return draft => {
          requireParagraph(draft, target);
          deleteParagraph(draft, target);
          return { result: { ref: p.ref } };
        };
      case 'insert_table_row': {
        const table = p.table ?? notInTable();
        if (edit.cells.length !== table.columns) {
          throw new DocumentToolError('CELL_COUNT', `This table has ${table.columns} columns; supply exactly ${table.columns} cell texts.`, { columns: table.columns });
        }
        edit.cells.forEach(text => lineText(text));
        return draft => {
          requireParagraph(draft, target);
          return { result: {}, created: insertTableRow(draft, target, table.story, edit.position, edit.cells) };
        };
      }
      case 'delete_table_row': {
        const table = p.table ?? notInTable();
        return draft => {
          requireParagraph(draft, target);
          return { result: { ref: p.ref, deletedTable: deleteTableRow(draft, target, table.story) } };
        };
      }
    }
  }

  private async checkCommentPart(): Promise<CommentSource | undefined> {
    const source = this.session.materializeDocx()?.originalBuffer;
    if (!source) return;
    if (this.commentPartCheck?.source !== source) this.commentPartCheck = { source, check: readCommentPart(source) };
    return this.commentPartCheck.check;
  }

  private styles(): ParagraphStyle[] {
    return this.paragraphStyles ??= (this.session.materializeDocx()?.package.styles?.styles ?? [])
      .filter(style => style.type === 'paragraph')
      .map(style => ({ id: style.styleId, name: style.name, isDefault: !!style.default }));
  }
}

export async function openDocx(bytes: Uint8Array, options: DocumentOptions = {}): Promise<DocxAgentDocument> {
  const session = await createYrsSession();
  try {
    session.seedFromDocx(bytes);
    return new DocxAgentDocument(session, options, true);
  } catch (error) { session.destroy(); throw error; }
}

export function attachDocx(session: YrsSession, options: DocumentOptions = {}): DocxAgentDocument {
  return new DocxAgentDocument(session, options);
}
