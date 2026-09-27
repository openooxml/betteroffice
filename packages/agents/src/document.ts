import { createYrsSession, saveYrsDocx, type YrsSession, type YrsStorySegment } from '@betteroffice/docx/yrs';
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
}

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

function visibleFormatting(attributes: Record<string, unknown>): Record<string, unknown> {
  const keys = new Set(['bold', 'italic', 'underline', 'strike', 'textColor', 'color', 'fontSize', 'fontFamily', 'fontFamilyAscii', 'fontFamilyHAnsi', 'highlight', 'verticalAlign', 'hyperlink', 'ins', 'del']);
  return Object.fromEntries(Object.entries(attributes).filter(([key]) => keys.has(key)).map(([key, value]) =>
    [key, JSON.stringify(value).length <= 500 ? value : '(formatting value exceeds preview limit)']));
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
      capabilities: { grep: 'literal', proposals: true, render: !!this.options.renderer, export: !!this.session.materializeDocx()?.originalBuffer },
      workflow: 'grep for the exact text to replace; read its ref for context; propose with match and newText; review; render; verify; export.',
    };
  }

  list(options: { story?: string; offset?: number; limit?: number; headingsOnly?: boolean } = {}) {
    this.index();
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 30, 1, 100, 'limit');
    this.checkStory(options.story);
    const entries = this.paragraphs.filter(p => (!options.story || p.story === options.story) &&
      (!options.headingsOnly || p.headingLevel !== null));
    return {
      version: this.epoch,
      items: entries.slice(offset, offset + limit).map(p => ({ ...this.reference(p), style: p.style, headingLevel: p.headingLevel, text: p.text.slice(0, 120), length: p.text.length })),
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
      if (p.runs.some(run => run.protected && (run.start < end && run.end > start || start === end && run.start <= start && run.end >= start))) {
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
    for (const story of this.session.storyIds().sort()) {
      const headings = new Map(this.session.headings(story).map(entry => [entry.paraId, entry.heading.outlineLevel + 1]));
      let text = '';
      let runs: TextRun[] = [];
      let segments: YrsStorySegment[] = [];
      let ordinal = 0;
      for (const segment of this.session.storySegments(story)) {
        segments.push(segment);
        if (segment.kind === 'pilcrow') {
          const key = JSON.stringify([story, segment.paraId]);
          const previous = this.identities.get(key);
          const fingerprint = JSON.stringify(segments);
          const p: Paragraph = {
            ref: previous?.ref ?? `p${this.nextRef++}`,
            revision: previous ? previous.revision + Number(previous.fingerprint !== fingerprint) : 1,
            story, paraId: segment.paraId, paragraph: ++ordinal, text, runs, fingerprint,
            style: typeof segment.properties.pStyle === 'string' ? segment.properties.pStyle : null,
            headingLevel: headings.get(segment.paraId) ?? null,
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
    const changes = [...held.value.changes].sort((a, b) => a.ref.localeCompare(b.ref) || b.start - a.start);
    for (const change of changes) {
      const { story, paraId } = held.targets.get(change.ref)!;
      draft.replaceRange({ story, start: { paraId, offset: change.start }, end: { paraId, offset: change.start + change.oldText.length } }, change.newText, author);
    }
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
