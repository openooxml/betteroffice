import { engines } from '#agents-engines';
import type { DocxDroppedStructure, DocxKeptParagraph, DocxCommentSaveBaseline, DocxCommentSaveIdentity, DocxParagraphTarget, DocxProposalInput, DocxProposalRecord, DocxProposalSnapshot, DocxProposalResult, DocxProposalStateRequest, YrsInlineFormatDelta, YrsSession, YrsStorySegment } from '@betteroffice/docx/yrs';
import { readCommentPart, readSavedCommentIdentities, verifyCommentOutput, type CommentSource } from './comment-output';
import { cellPosition, commentBody, commentRanges, commentRoot, commentText, continueList, deleteParagraph, deleteTableRow, formatParagraph, hostEdit, insertParagraphsBefore, insertParagraphsWithAnchorFormat, insertSectionBreak, replaceCommentText, setSection, insertTableRow, lineText, listLevel, mergeCells, notInTable, requireParagraph, resolveStyle, tableCells, type ParagraphStyle } from './docx-edits';
import { createDocxBatchSchema, type DocxAgentBatch, type DocxAgentEdit, type DocxEditOutcome } from './docx-schema';
import { compareDocxStructure, docxModelCounts, type DocxModelCounts } from './docx-structure';
import { editLimit } from './edit-limits';
import { cleanupGuard } from './cleanup';
import {
  DocumentToolError,
  TrackedRevisionConflictError,
  AcceptDateUnsupportedError,
  ProposalDecidedInEditorError,
  unknownMatchError,
  type DocumentOptions,
  type DocxRunPage,
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

interface SectionSummary {
  number: number;
  firstParagraph: string | null;
  lastParagraph: string | null;
  orientation: 'portrait' | 'landscape';
  pageWidth: number;
  pageHeight: number;
  margins: { [name: string]: number };
}

interface SectionPage {
  items: SectionSummary[];
  nextOffset: number | null;
  total: number;
}

const BLOCK_EMBEDS = new Set(['table', 'blockSdt', 'pageBreak', 'columnBreak']);

type EditStep = (draft: YrsSession) => { result: Omit<DocxEditOutcome, 'index' | 'op' | 'refs'>; created?: DocxParagraphTarget[] };

interface HeldProposal {
  value: Proposal;
  targets: Map<string, { story: string; paraId: string; fingerprint: string }>;
  engine: DocxProposalRecord[];
  registered: Set<string>;
  editorDecisions: Map<string, 'accepted' | 'rejected'>;
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

function fixedRunKind(run: TextRun): string | undefined {
  return run.kind ?? (run.protected || run.formatting.hyperlink ? 'protected' : undefined);
}

function literal(query: string, caseSensitive: boolean): RegExp {
  return new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseSensitive ? 'gu' : 'giu');
}

export class DocxAgentDocument {
  readonly name: string;
  private get attached() { return !this.ownsSession; }
  private get proposalPeer() { return this.attached ? this.options.editPeer : undefined; }
  private epoch = 0;
  private indexedEpoch = -1;
  private proposalOperation?: { request?: DocxProposalStateRequest; withdrawal?: readonly string[]; withdrawalSnapshot?: DocxProposalSnapshot; snapshots: DocxProposalSnapshot[] };
  private proposalSnapshot?: DocxProposalSnapshot;
  private readonly ownStateVersions = new Map<string, Set<number>>();
  private readonly ownWithdrawals = new Set<string>();
  private readonly editorDecisionRevisions = new Map<string, number>();
  private paragraphs: Paragraph[] = [];
  private readonly identities = new Map<string, Paragraph>();
  private readonly runRevisions = new Map<string, { fingerprint: string; revision: number }>();
  private readonly proposals = new Map<string, HeldProposal>();
  private readonly matches = new Map<string, Omit<ExactTextEdit, 'newText'>>();
  private readonly cursors = new Map<string, { epoch: number; signature: string; paragraph: number; offset: number }>();
  private readonly unsubscribe: () => void;
  private readonly unsubscribeProposals: () => void;
  private readonly proposalPrefix = crypto.randomUUID();
  private nextRef = 1;
  private nextProposal = 1;
  private nextCursor = 1;
  private nextMatch = 1;
  private closed = false;
  private paragraphStyles?: ParagraphStyle[];
  private commentPartCheck?: { source: ArrayBuffer; allowLegacy: boolean; check: Promise<CommentSource> };
  private commentPackageBaseline?: { version: number; deletion: boolean; bytes: Promise<Uint8Array> };
  private resolvedComments = new Map<string, boolean>();
  private deletedComments = new Set<string>();
  private commentSaveIdentities = new Map<string, DocxCommentSaveIdentity>();
  private commentSaveBaseline?: DocxCommentSaveBaseline;
  private commentSaveBaselineVersion?: number;
  private readonly textEditedComments = new Set<string>();
  private readonly maxEdits: number;
  private sourceCounts?: DocxModelCounts;

  constructor(
    readonly session: YrsSession,
    private readonly options: DocumentOptions = {},
    private readonly ownsSession = false,
  ) {
    this.maxEdits = editLimit(options.maxEdits);
    this.name = options.name ?? 'document.docx';
    this.unsubscribe = session.onUpdate(() => { this.epoch++; });
    this.unsubscribeProposals = this.attached ? session.onProposalChange(snapshot => this.observeProposals(snapshot)) : () => {};
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
      workflow: 'grep for the exact text to replace; read its ref for context; propose with match and newText; review; render; verify; export. For styles, bold/italic/underline, paragraphs, table rows, cell merges, comments and page setup, use docx_edit with the outline version.',
    };
  }

  list(options: { story: 'comments'; offset?: number; limit?: number; headingsOnly?: boolean }): OutlinePage<CommentSummary>;
  list(options?: { story?: string; offset?: number; limit?: number; headingsOnly?: boolean; sectionOffset?: number; sectionLimit?: number }): OutlinePage<ParagraphSummary> & { sections: SectionPage };
  list(options: { story?: string; offset?: number; limit?: number; headingsOnly?: boolean; sectionOffset?: number; sectionLimit?: number } = {}) {
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

  private paragraphOutline(options: { story?: string; offset?: number; limit?: number; headingsOnly?: boolean; sectionOffset?: number; sectionLimit?: number } = {}): OutlinePage<ParagraphSummary> & { sections: SectionPage } {
    this.index();
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 30, 1, 100, 'limit');
    this.checkStory(options.story);
    const sectionOffset = integer(options.sectionOffset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'sectionOffset');
    const sectionLimit = integer(options.sectionLimit ?? 20, 1, 100, 'sectionLimit');
    const body = this.paragraphs.filter(paragraph => paragraph.story === 'body');
    const sections = this.session.listSections();
    const marginFields = { top: 'marginTop', bottom: 'marginBottom', left: 'marginLeft', right: 'marginRight', header: 'headerDistance', footer: 'footerDistance', gutter: 'gutter' } as const;
    const entries = this.paragraphs.filter(p => (!options.story || p.story === options.story) &&
      (!options.headingsOnly || p.headingLevel !== null));
    return {
      version: this.epoch,
      items: entries.slice(offset, offset + limit).map(p => ({ ...this.reference(p), style: p.style, headingLevel: p.headingLevel, text: p.text.slice(0, 120), length: p.text.length, ...(p.table && { table: { ...p.table } }), ...(p.list && { list: { ...p.list } }) })),
      nextOffset: offset + limit < entries.length ? offset + limit : null,
      total: entries.length,
      sections: {
        items: sections.slice(sectionOffset, sectionOffset + sectionLimit).map((section, index) => ({
          number: sectionOffset + index + 1,
          firstParagraph: body[section.firstParagraph - 1]?.ref ?? null,
          lastParagraph: body[section.lastParagraph - 1]?.ref ?? null,
          orientation: section.resolved.orientation ?? 'portrait',
          pageWidth: section.resolved.pageWidth ?? 12240,
          pageHeight: section.resolved.pageHeight ?? 15840,
          margins: Object.fromEntries(Object.entries(marginFields).map(([name, field]) => [name, section.resolved[field] ?? (name === 'gutter' ? 0 : name === 'header' || name === 'footer' ? 720 : 1440)])),
        })),
        total: sections.length,
        nextOffset: sectionOffset + sectionLimit < sections.length ? sectionOffset + sectionLimit : null,
      },
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

  readRuns(options: { story?: string; offset?: number; limit?: number; maxChars?: number } = {}): DocxRunPage {
    this.index();
    this.checkStory(options.story);
    const offset = integer(options.offset ?? 0, 0, Number.MAX_SAFE_INTEGER, 'offset');
    const limit = integer(options.limit ?? 50, 1, 200, 'limit');
    const maxChars = integer(options.maxChars ?? 32000, 1, Number.MAX_SAFE_INTEGER, 'maxChars');
    const entries = this.paragraphs.filter(p => !options.story || p.story === options.story);
    const paragraphs: DocxRunPage['paragraphs'] = [];
    const formats: DocxRunPage['formats'] = {};
    const keys = new Map<string, string>();
    let size = 0;
    for (const p of entries.slice(offset, offset + limit)) {
      if (paragraphs.length && size + p.text.length > maxChars) break;
      const runs = p.runs.map((run, index) => {
        const id = `${p.ref}.${index + 1}`;
        const kind = fixedRunKind(run);
        if (kind) return { id, kind };
        const text = p.text.slice(run.start, run.end);
        if (!Object.keys(run.formatting).length) return { id, text };
        const signature = JSON.stringify(Object.entries(run.formatting).sort(([a], [b]) => a.localeCompare(b)));
        let f = keys.get(signature);
        if (!f) {
          f = `f${keys.size + 1}`;
          keys.set(signature, f);
          formats[f] = structuredClone(run.formatting);
        }
        return { id, text, f };
      });
      paragraphs.push({ ref: p.ref, revision: this.runRevision(p), text: p.text, runs });
      size += p.text.length;
    }
    const next = offset + paragraphs.length;
    return { version: this.epoch, paragraphs, formats, nextOffset: next < entries.length ? next : null, total: entries.length };
  }

  async edit(input: DocxAgentBatch) {
    this.index();
    if (Array.isArray(input?.edits) && input.edits.some(edit => edit?.op === 'delete_comment')
      && input.edits.some(edit => ['add_comment', 'reply_comment', 'resolve_comment', 'replace_comment_text'].includes(edit?.op))) {
      throw new DocumentToolError('INVALID_ARGUMENT', 'delete_comment cannot be combined with other comment operations in one batch; run it in its own docx_edit call.');
    }
    const parsed = createDocxBatchSchema(this.maxEdits).safeParse(input);
    if (!parsed.success) throw new DocumentToolError('INVALID_ARGUMENT', `Use the docx_edit schema: version and 1-${this.maxEdits} edits.`, { issues: parsed.error.issues.map(issue => ({ path: issue.path, message: issue.message })) });
    const { version, edits } = parsed.data;
    if (version !== this.epoch) throw new DocumentToolError('STALE_VERSION', 'The document changed. Read office_outline or office_grep again and copy its version.');
    const maxChars = Math.max(64_000, 1000 * this.maxEdits);
    if (JSON.stringify(edits).length > maxChars) throw new DocumentToolError('EDIT_LIMIT', `A batch holds at most ${maxChars} characters of input. Split it.`);
    const paragraphs = new Set<string>();
    for (const edit of edits) if (edit.op === 'set_runs') {
      if (paragraphs.has(edit.paragraph)) throw new DocumentToolError('INVALID_ARGUMENT', 'Use one set_runs operation per paragraph in a batch.');
      paragraphs.add(edit.paragraph);
    }
    for (const edit of edits) if ('match' in edit && edit.match !== undefined && paragraphs.has(this.matches.get(edit.match)?.ref ?? '')) {
      throw new DocumentToolError('INVALID_ARGUMENT', 'set_runs moves text offsets, so a match in the same paragraph needs its own batch. Run grep again after set_runs.');
    }
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

  private prepareChanges(input: { author: string; note?: string; edits: TextEdit[] }): HeldProposal {
    this.index();
    if (!input.author || input.author.length > 200 || (input.note?.length ?? 0) > 2000 || !Array.isArray(input.edits) || input.edits.length < 1 || input.edits.length > this.maxEdits) {
      throw new DocumentToolError('INVALID_ARGUMENT', `Supply an author (1-200 characters), note (up to 2000), and 1-${this.maxEdits} edits.`);
    }
    if (this.proposals.size >= 64) {
      const closed = [...this.proposals].find(([, held]) => held.value.status !== 'pending');
      if (closed) {
        this.checkEditorDecision(closed[1]);
        this.forgetProposalOwnership(closed[1]);
        this.proposals.delete(closed[0]);
      } else throw new DocumentToolError('PROPOSAL_LIMIT', 'Reject or discard a pending proposal before staging more than 64.');
    }
    const changes: ProposalChange[] = [];
    const targets: HeldProposal['targets'] = new Map();
    let size = 0;
    for (const inputEdit of input.edits) {
      let edit: ExactTextEdit;
      if ('oldText' in inputEdit) edit = inputEdit;
      else {
        const target = this.matches.get(inputEdit.match);
        if (!target) throw unknownMatchError(inputEdit.match, 'm1', /^m\d+$/u, 'Match expired or belongs to another document. Run grep again and copy its match ID.');
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
        if (this.attached) {
          const conflicts = this.revisionConflicts([{ ...this.reference(p), start, oldText: edit.oldText, newText: edit.newText }]);
          if (conflicts.length) throw new TrackedRevisionConflictError(conflicts);
        }
        throw new DocumentToolError('PROTECTED_CONTENT', 'This range contains an embed or an existing tracked change. Choose an ordinary text range.');
      }
      if (changes.some(change => change.ref === p.ref && start <= change.start + change.oldText.length && end >= change.start)) {
        throw new DocumentToolError('OVERLAPPING_EDITS', 'Edits within one proposal must have separate, non-touching ranges.');
      }
      changes.push({ ...this.reference(p), start, oldText: edit.oldText, newText: edit.newText });
      targets.set(p.ref, { story: p.story, paraId: p.paraId, fingerprint: p.fingerprint });
    }
    const value: Proposal = { id: `proposal${this.nextProposal++}`, author: input.author, note: input.note ?? '', status: 'pending', changes, staleRefs: [] };
    return { value, targets, engine: [], registered: new Set(), editorDecisions: new Map() };
  }

  propose(input: { author: string; note?: string; edits: TextEdit[] }): Proposal {
    if (!this.attached) {
      const held = this.prepareChanges(input);
      this.proposals.set(held.value.id, held);
      return structuredClone(held.value);
    }
    if (this.proposalPeer) throw new DocumentToolError('UNSUPPORTED', 'Use proposeAsync with a DOCX editor edit peer.');
    const plan = this.prepareProposal(input);
    try {
      const snapshot = this.proposalResult(this.session.proposeChanges({ expectVersion: plan.version, expectPreviewVersion: plan.previewVersion, proposals: plan.proposals }), plan.held.value.changes);
      return this.recordProposal(plan, snapshot);
    } catch (error) {
      this.forgetProposalOwnership(plan.held);
      this.proposals.delete(plan.held.value.id);
      throw error;
    }
  }

  async proposeAsync(input: { author: string; note?: string; edits: TextEdit[] }): Promise<Proposal> {
    if (!this.proposalPeer) return this.propose(input);
    const plan = this.prepareProposal(input);
    try {
      const result = await this.proposalPeer.proposeChanges({ expectVersion: plan.version, expectPreviewVersion: plan.previewVersion, proposals: plan.proposals });
      this.synchronizeProposals(this.session.getProposals());
      const snapshot = this.proposalResult(result, plan.held.value.changes);
      const value = this.recordProposal(plan, snapshot);
      if (!plan.held.editorDecisions.size) this.checkDecisionReferences(plan.references, plan.held);
      return value;
    } catch (error) {
      if (!plan.held.editorDecisions.size && plan.held.registered.size) {
        const current = this.proposalSnapshot ?? this.session.getProposals();
        const ids = [...plan.held.registered];
        this.proposalOperation = { withdrawal: ids, withdrawalSnapshot: current, snapshots: [] };
        try {
          const result = await this.proposalPeer.withdrawProposals({ expectVersion: current.version, expectPreviewVersion: current.previewVersion, ids });
          this.finishProposalOperation(result);
        } finally { if (this.proposalOperation) this.finishProposalOperation(); }
        this.checkEditorDecision(plan.held);
      }
      if (!plan.held.editorDecisions.size && !plan.held.registered.size) {
        this.forgetProposalOwnership(plan.held);
        this.proposals.delete(plan.held.value.id);
      }
      throw error;
    }
  }

  private prepareProposal(input: { author: string; note?: string; edits: TextEdit[] }) {
    const held = this.prepareChanges(input);
    const { value, targets } = held;
    const changes = value.changes;
    const references = this.captureReferences();
    const sessionId = this.session.paragraphIdentities().sessionId;
    const date = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const proposals = changes.map(trimChange).map((change, index): DocxProposalInput => {
      const { story, paraId } = targets.get(change.ref)!;
      const paragraph = { kind: 'session' as const, sessionId, story, paraId };
      const id = `${this.proposalPrefix}:${value.id}:${index}`;
      const suggest = { author: input.author, date };
      const offset = this.acceptedOffset(story, paraId, change.start);
      if (!change.oldText) return { id, paragraph, suggest, op: 'insertText', at: { offset }, text: change.newText };
      return { id, paragraph, suggest, op: 'replaceText', at: { offset }, search: change.oldText, replaceWith: change.newText };
    });
    held.engine = proposals.map((proposal, index): DocxProposalRecord => {
      const target = targets.get(changes[index].ref)!;
      return { id: proposal.id, state: 'proposed', paragraph: { kind: 'session', sessionId, story: target.story, paraId: target.paraId }, revisionIds: [], changed: false };
    });
    this.proposals.set(value.id, held);
    const snapshot = this.proposalSnapshot ?? this.session.getProposals();
    return { held, proposals, references, version: snapshot.version, previewVersion: snapshot.previewVersion };
  }

  private recordProposal(plan: { held: HeldProposal; proposals: DocxProposalInput[]; references: ReturnType<DocxAgentDocument['captureReferences']> }, snapshot: DocxProposalSnapshot): Proposal {
    this.synchronizeProposals(snapshot);
    this.synchronizeProposals(this.session.getProposals());
    this.epoch++;
    this.preserveReferences(plan.references, snapshot.version);
    return structuredClone(plan.held.value);
  }

  review(id: string): Proposal {
    const held = this.proposal(id);
    return structuredClone({ ...held.value, staleRefs: held.value.status === 'pending' ? this.staleRefs(held) : [] });
  }

  listProposals() {
    this.index();
    return [...this.proposals.values()].map(({ value }) => ({ id: value.id, author: value.author, note: value.note.slice(0, 120), status: value.status, changes: value.changes.length, ...(value.records && { records: structuredClone(value.records) }) }));
  }

  private async acceptHeadless(id: string, options: { tracked?: boolean; date?: string } = {}) {
    const held = this.pending(id);
    if (options.tracked && options.date !== undefined && !Number.isFinite(Date.parse(options.date))) {
      throw new DocumentToolError('INVALID_ARGUMENT', 'Tracked acceptance requires an ISO date.');
    }
    const draft = await this.fork();
    const release = cleanupGuard(() => draft.destroy());
    try {
      this.checkPending(held);
      const vector = draft.encodeStateVector();
      this.apply(draft, held, options.tracked ? { name: held.value.author, date: options.date ?? new Date().toISOString().replace(/\.\d+Z$/, 'Z') } : undefined);
      this.checkPending(held);
      this.session.addUndoBoundary();
      try { this.session.applyLocalUpdate(draft.encodeStateAsUpdate(vector)); }
      finally { this.session.addUndoBoundary(); }
      held.value.status = 'accepted';
      return { id, status: 'accepted' as const, tracked: options.tracked ?? false, changes: held.value.changes.length };
    } catch (error) { throw release.failed(error); } finally { release.run(); }
  }

  async accept(id: string, options: { tracked?: boolean; date?: string } = {}) {
    if (!this.attached) return this.acceptHeadless(id, options);
    const held = this.proposal(id);
    this.checkEditorDecision(held);
    if (options.date !== undefined) throw new AcceptDateUnsupportedError();
    this.checkPending(held);
    await Promise.resolve();
    this.synchronizeProposals(this.session.getProposals());
    this.checkEditorDecision(held);
    this.checkPending(held);
    this.session.addUndoBoundary();
    try { await this.decideProposalAsync(held, 'accepted', options.tracked ?? false); }
    finally { this.session.addUndoBoundary(); }
    return { id, status: this.settledStatus(held), tracked: options.tracked ?? false, changes: held.value.changes.length };
  }

  reject(id: string) {
    const held = this.proposal(id);
    this.checkEditorDecision(held);
    if (held.value.status !== 'pending') throw new DocumentToolError('PROPOSAL_CLOSED', `Proposal is ${held.value.status}.`);
    if (this.attached) this.decideProposal(held, 'rejected');
    else held.value.status = 'rejected';
    return { id, status: this.attached ? this.settledStatus(held) : 'rejected' as const };
  }

  async rejectAsync(id: string) {
    if (!this.proposalPeer) return this.reject(id);
    const held = this.proposal(id);
    this.checkEditorDecision(held);
    if (held.value.status !== 'pending') throw new DocumentToolError('PROPOSAL_CLOSED', `Proposal is ${held.value.status}.`);
    await this.decideProposalAsync(held, 'rejected');
    return { id, status: this.settledStatus(held) };
  }

  async discardAsync(id: string) {
    if (!this.proposalPeer) return this.discard(id);
    const held = this.proposal(id);
    this.checkEditorDecision(held);
    if (held.value.status === 'pending') await this.decideProposalAsync(held, 'rejected');
    this.forgetProposalOwnership(held);
    this.proposals.delete(id);
  }

  discard(id: string) {
    const held = this.proposal(id);
    this.checkEditorDecision(held);
    if (this.attached && held.value.status === 'pending') this.decideProposal(held, 'rejected');
    this.forgetProposalOwnership(held);
    this.proposals.delete(id);
  }

  async render(page = 1, proposalId?: string) {
    const held = proposalId ? this.pending(proposalId) : undefined;
    integer(page, 1, 100_000, 'page');
    if (!this.options.renderer) throw new DocumentToolError('UNSUPPORTED', 'No renderer configured. Supply a DocumentRenderer or use the MCP server.');
    const draft = await this.fork();
    const release = cleanupGuard(() => draft.destroy());
    try {
      if (held) this.checkPending(held);
      if (this.attached) this.projectProposals(draft, held);
      else if (held) this.apply(draft, held);
      return await this.options.renderer(draft, page);
    } catch (error) { throw release.failed(error); } finally { release.run(); }
  }

  async export(proposalId?: string): Promise<Uint8Array> {
    return this.exportBytes(proposalId);
  }

  /** Paragraphs the last export wrote as their source bytes, see `DocxSavedDocument.keptUnchanged`. */
  exportKeptUnchanged() {
    return this.lastKeptUnchanged;
  }

  private lastKeptUnchanged: { structure: number; paragraphs: Array<DocxKeptParagraph & { ref?: string }> } | undefined;

  /** Story parts the last export wrote whole with less structure, see `DocxSavedDocument.structureDropped`. */
  exportStructureDropped() {
    return this.lastStructureDropped;
  }

  private lastStructureDropped: Array<Omit<DocxDroppedStructure, 'paragraphs'> & { paragraphs: Array<DocxDroppedStructure['paragraphs'][number] & { ref?: string }> }> | undefined;

  private async exportBytes(proposalId?: string, updateModifiedDate = true, inspect?: (draft: YrsSession) => void): Promise<Uint8Array> {
    const held = proposalId ? this.pending(proposalId) : undefined;
    const draft = await this.fork();
    const release = cleanupGuard(() => draft.destroy());
    try {
      if (held) this.checkPending(held);
      if (this.attached) this.projectProposals(draft, held);
      else if (held) this.apply(draft, held);
      inspect?.(draft);
      if (!draft.materializeDocx()?.originalBuffer) throw new DocumentToolError('UNSUPPORTED', 'Export requires a session opened from DOCX bytes.');
      const version = this.epoch;
      const comments = draft.listComments();
      const baseline = this.commentSaveBaseline && JSON.stringify(comments) === JSON.stringify(this.commentSaveBaseline.comments)
        ? this.commentSaveBaseline : undefined;
      const saved = await engines.get('docx').module.yrs.saveYrsDocx(draft, { updateModifiedDate,
        commentIdentities: this.deletedComments.size ? this.commentSaveIdentities : undefined,
        commentBaseline: baseline,
        commentDeletion: !!baseline && !held && this.commentSaveBaselineVersion === version,
      });
      const bytes = saved.bytes;
      const refs = new Map([...this.identities.values()].map(paragraph => [paragraph.paraId, paragraph.ref]));
      this.lastKeptUnchanged = saved.keptUnchanged && {
        ...saved.keptUnchanged,
        paragraphs: saved.keptUnchanged.paragraphs.map(paragraph => ({
          ...paragraph,
          ...(paragraph.paraId && refs.has(paragraph.paraId) ? { ref: refs.get(paragraph.paraId) } : {}),
        })),
      };
      this.lastStructureDropped = saved.structureDropped?.map(part => ({
        ...part,
        paragraphs: part.paragraphs.map(paragraph => ({
          ...paragraph,
          ...(refs.has(paragraph.paraId) ? { ref: refs.get(paragraph.paraId) } : {}),
        })),
      }));
      const seeded = new Set((draft.materializeDocx()?.package.document.comments ?? []).map(comment => String(comment.id)));
      if (this.deletedComments.size && comments.some(comment => !seeded.has(comment.id))) {
        const source = await this.checkCommentPart(true);
        if (source) {
          const identities = await readSavedCommentIdentities(source, bytes, comments, this.commentSaveIdentities);
          if (this.epoch === version) this.commentSaveIdentities = identities;
        }
      }
      if (baseline && !held && this.epoch === version) {
        this.commentSaveBaseline = { bytes, comments };
        this.commentSaveBaselineVersion = version;
      }
      return bytes;
    } catch (error) {
      if (error instanceof engines.get('docx').module.yrs.DocxSectionError) throw release.failed(new DocumentToolError('UNSUPPORTED', error.message));
      throw release.failed(error);
    } finally { release.run(); }
  }

  async verify(proposalId?: string, options: { structure?: boolean } = {}) {
    const structure = options.structure ?? true;
    let after: DocxModelCounts | undefined;
    const bytes = await this.exportBytes(proposalId, true, structure ? draft => { after = docxModelCounts(draft); } : undefined);
    const reopened = await openDocx(bytes);
    const release = cleanupGuard(() => reopened.close());
    try {
      const kept = this.lastKeptUnchanged;
      const dropped = this.lastStructureDropped;
      const paragraphs = reopened.overview().paragraphs;
      const source = this.session.materializeDocx()?.originalBuffer;
      if (!structure || !source || !after) return { reopened: true, bytes: bytes.length, paragraphs, checks: ['save', 'reopen'], unchecked: ['visual fidelity', 'semantic correctness', ...(structure ? [] : ['structure'])], structureDropped: dropped, keptUnchanged: kept };
      const original = new Uint8Array(source);
      this.sourceCounts ??= await sourceModelCounts(original);
      const report = await compareDocxStructure(original, bytes, this.sourceCounts, after);
      return { reopened: true, bytes: bytes.length, paragraphs, checks: ['save', 'reopen', 'structure'], structure: report, unchecked: ['visual fidelity', 'semantic correctness'], structureDropped: dropped, keptUnchanged: kept };
    } catch (error) { throw release.failed(error); } finally { release.run(); }
  }

  close() {
    if (this.closed) return;
    this.closed = true;
    this.unsubscribe();
    this.unsubscribeProposals();
    this.proposals.clear();
    this.ownStateVersions.clear();
    this.ownWithdrawals.clear();
    this.cursors.clear();
    this.matches.clear();
    this.identities.clear();
    this.runRevisions.clear();
    this.paragraphs = [];
    if (this.ownsSession) this.session.destroy();
  }

  private index() {
    if (this.closed) throw new DocumentToolError('CLOSED', 'This document is closed. Open it again.');
    if (this.indexedEpoch === this.epoch) return;
    let snapshot = this.attached ? this.session.getProposals() : undefined;
    if (snapshot) {
      this.synchronizeProposals(snapshot);
      snapshot = this.proposalSnapshot ?? snapshot;
    }
    const paragraphs: Paragraph[] = [];
    const active = new Set<string>();
    const cells = new Map<string, TableCellPosition>();
    const stories = this.session.storyIds();
    const segmentsByStory = new Map(stories.map(story => [story, this.attached ? this.projectSegments(this.session.storySegments(story), snapshot!) : this.session.storySegments(story)]));
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
          const decisionRevision = this.editorDecisionRevisions.get(key);
          const fingerprint = JSON.stringify(segments) + (decisionRevision === undefined ? '' : `:${decisionRevision}`);
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
          const values = segment.kind === 'text' ? segment.text.split(/([\t\r\n\v\f])/u).filter(Boolean) : ['\ufffc'];
          for (const value of values) {
            const kind = segment.kind === 'embed' ? segment.embedKind : value === '\t' ? 'tab' : /^[\r\n\v\f]$/u.test(value) ? 'break' : undefined;
            runs.push({ start: text.length, end: text.length + value.length, formatting: visibleFormatting(segment.attributes), protected: segment.kind === 'embed' || !!segment.attributes.ins || !!segment.attributes.del,
              ...(kind && { kind }) });
            text += value;
          }
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
    if (this.attached) this.synchronizeProposals(this.session.getProposals());
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
    this.checkEditorDecision(held);
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

  private async fork(clientId?: number): Promise<YrsSession> {
    this.index();
    const { module } = await engines.load('docx');
    const draft = await module.yrs.createYrsSession({ clientId });
    try {
      const source = this.session.materializeDocx()?.originalBuffer;
      const state = this.session.encodeState();
      if (source) {
        const bytes = new Uint8Array(source);
        const host = draft.openDocx(bytes, false);
        draft.bootstrapPeer(state, this.session.encodePeerMetadata(), bytes, host);
      } else {
        draft.loadState(state);
      }
      return draft;
    } catch (error) { throw cleanupGuard(() => draft.destroy()).failed(error); }
  }

  private runRevision(p: Paragraph) {
    const key = JSON.stringify([p.story, p.paraId]);
    const held = this.runRevisions.get(key);
    if (!held) { this.runRevisions.set(key, { fingerprint: p.fingerprint, revision: p.revision }); return p.revision; }
    if (held.fingerprint !== p.fingerprint) { held.fingerprint = p.fingerprint; held.revision++; }
    return held.revision;
  }

  private captureReferences() {
    this.index();
    return { epoch: this.indexedEpoch, paragraphs: this.paragraphs.map(p => ({ ...p })), decisions: new Map(this.editorDecisionRevisions) };
  }

  private preserveReferences(before: ReturnType<DocxAgentDocument['captureReferences']>, version: string, changes: ProposalChange[] = []) {
    if (this.session.version() !== version) return;
    this.index();
    const touched = new Set(before.paragraphs.filter(previous => {
      const key = JSON.stringify([previous.story, previous.paraId]);
      return before.decisions.get(key) !== this.editorDecisionRevisions.get(key);
    }).map(previous => previous.ref));
    if (before.paragraphs.length !== this.paragraphs.length || before.paragraphs.some(previous => {
      if (touched.has(previous.ref)) return false;
      const current = this.paragraphs.find(p => p.ref === previous.ref);
      const edits = changes.filter(change => change.ref === previous.ref).sort((a, b) => b.start - a.start);
      let expected = previous.text;
      for (const edit of edits) expected = expected.slice(0, edit.start) + edit.newText + expected.slice(edit.start + edit.oldText.length);
      return !current || current.text !== expected || (!edits.length && current.fingerprint !== previous.fingerprint);
    })) return;
    const shift = (ref: string, offset: number, right = false, afterInsertion = right) => {
      let delta = 0;
      for (const change of changes.filter(change => change.ref === ref).sort((a, b) => a.start - b.start)) {
        if (offset < change.start) break;
        if (offset === change.start) return offset + delta + (!change.oldText && afterInsertion ? change.newText.length : 0);
        if (offset < change.start + change.oldText.length) return change.start + delta + (right ? change.newText.length : 0);
        delta += change.newText.length - change.oldText.length;
      }
      return offset + delta;
    };
    for (const p of this.paragraphs) {
      const previous = before.paragraphs.find(entry => entry.ref === p.ref);
      if (!previous || touched.has(p.ref)) continue;
      p.revision = previous.revision;
      for (const match of this.matches.values()) {
        if (match.ref !== p.ref || match.revision !== previous.revision || match.start === undefined) continue;
        if (previous.text.slice(match.start, match.start + match.oldText.length) !== match.oldText) continue;
        const start = shift(p.ref, match.start, true);
        const end = shift(p.ref, match.start + match.oldText.length, true, false);
        match.start = start;
        match.oldText = p.text.slice(start, end);
      }
      for (const held of this.proposals.values()) {
        const target = held.targets.get(p.ref);
        if (target?.fingerprint !== previous.fingerprint) continue;
        target.fingerprint = p.fingerprint;
        if (held.value.status === 'pending') for (const change of held.value.changes) {
          if (change.ref === p.ref) change.start = shift(p.ref, change.start);
        }
      }
    }
    if (touched.size) return;
    for (const cursor of this.cursors.values()) {
      if (cursor.epoch !== before.epoch) continue;
      const paragraph = before.paragraphs[cursor.paragraph];
      if (paragraph) cursor.offset = shift(paragraph.ref, cursor.offset);
      cursor.epoch = this.epoch;
    }
  }

  private revisionConflicts(changes: ProposalChange[]) {
    const snapshot = this.proposalSnapshot ?? this.session.getProposals();
    return this.session.listRevisions().flatMap(revision => {
      const paragraphs = this.paragraphs.filter(p => p.story === revision.story);
      const first = paragraphs.findIndex(p => p.paraId === revision.range.start.paraId);
      const last = paragraphs.findIndex(p => p.paraId === revision.range.end.paraId);
      if (first < 0 || last < 0) return [];
      const project = (paragraph: Paragraph, offset: number) => {
        let segments: YrsStorySegment[] = [];
        for (const segment of this.session.storySegments(revision.story)) {
          if (segment.kind === 'pilcrow') {
            if (segment.paraId === paragraph.paraId) break;
            segments = [];
          } else segments.push(segment);
        }
        let raw = 0, visible = 0;
        for (const segment of segments) {
          const length = segment.kind === 'text' ? segment.text.length : 1;
          const shown = this.projectSegments([segment], snapshot).length > 0;
          if (offset <= raw + length) return visible + (shown ? Math.max(0, offset - raw) : 0);
          raw += length;
          if (shown) visible += length;
        }
        return visible;
      };
      const conflicting = changes.some(change => {
        const index = paragraphs.findIndex(p => p.ref === change.ref);
        if (revision.kind === 'pPrChange' && index === last + 1 && change.start === 0) return true;
        if (index < first || index > last) return false;
        if (revision.kind.startsWith('pPr')) return true;
        const paragraph = paragraphs[index]!;
        const start = index === first ? project(paragraph, revision.range.start.offset) : 0;
        const end = index === last ? project(paragraph, revision.range.end.offset) : paragraph.text.length;
        return change.start <= end && change.start + change.oldText.length >= start;
      });
      if (!conflicting) return [];
      return [{ revisionId: revision.revisionId, author: revision.author, kind: revision.kind, range: revision.range, ref: paragraphs[first]!.ref }];
    });
  }

  private proposalResult(result: DocxProposalResult, changes?: ProposalChange[]): DocxProposalSnapshot {
    if (!result.ok && result.failure.code === 'tracked-revision-conflict') {
      const affected = changes ?? [...this.proposals.values()].filter(held => held.value.status === 'pending').flatMap(held => held.value.changes);
      const conflicts = this.revisionConflicts(affected);
      if (conflicts.length) throw new TrackedRevisionConflictError(conflicts);
      throw new DocumentToolError('PROTECTED_CONTENT', result.failure.message, result.failure);
    }
    if (!result.ok) throw new DocumentToolError((result.failure.code === 'stale-version' || result.failure.code === 'stale-preview') ? 'STALE_TARGET' : result.failure.code, result.failure.message, result.failure);
    return result.snapshot;
  }

  private synchronizeProposals(snapshot: DocxProposalSnapshot) {
    if (this.proposalOperation) {
      this.proposalOperation.snapshots.push(snapshot);
      return;
    }
    const previous = this.proposalSnapshot;
    if (previous && (snapshot.previewVersion < previous.previewVersion
      || (snapshot.revision ?? 0) < (previous.revision ?? 0))) return;
    this.proposalSnapshot = snapshot;
    const records = new Map(snapshot.proposals.map(record => [record.id, record]));
    const withdrawn = new Map((snapshot.withdrawn ?? []).map(record => [record.id, record]));
    const touched = new Set<string>();
    for (const held of this.proposals.values()) {
      held.engine = held.engine.map(record => {
        const next = records.get(record.id);
        const removed = withdrawn.get(record.id);
        if (!next && !removed && !held.registered.has(record.id)) return record;
        const current = next ?? removed ?? { ...record, state: record.state === 'proposed' ? 'rejected' as const : record.state };
        const stateVersion = current.stateVersion ?? snapshot.previewVersion;
        const changed = current.state !== record.state || (current.stateVersion ?? 0) > (record.stateVersion ?? 0)
          || (!next && held.registered.has(record.id));
        const own = next ? this.ownStateVersions.get(record.id)?.has(stateVersion) : this.ownWithdrawals.has(record.id);
        if (changed && !own) {
          const state = current.state === 'proposed' ? current.decisionState : current.state;
          if (state) held.editorDecisions.set(record.id, state);
          touched.add(JSON.stringify([record.paragraph.story, record.paragraph.paraId]));
        }
        if (next) held.registered.add(record.id);
        else held.registered.delete(record.id);
        return current;
      });
      held.value.records = held.engine.map(({ id, state }) => ({ id, state }));
      held.value.status = held.engine.some(record => record.state === 'proposed') ? 'pending' : this.settledStatus(held);
      if (!held.registered.size) this.forgetProposalOwnership(held);
    }
    for (const key of touched) this.editorDecisionRevisions.set(key, (this.editorDecisionRevisions.get(key) ?? 0) + 1);
    if (touched.size) this.indexedEpoch = -1;
  }

  private finishProposalOperation(result?: DocxProposalResult) {
    const operation = this.proposalOperation!;
    this.proposalOperation = undefined;
    if (result?.ok) {
      for (const change of operation.request?.changes ?? []) {
        const versions = this.ownStateVersions.get(change.id) ?? new Set<number>();
        versions.add(operation.request!.expectPreviewVersion + 1);
        this.ownStateVersions.set(change.id, versions);
      }
      const before = operation.withdrawalSnapshot;
      const withdrawn = result.withdrawnIds ?? (before?.revision !== undefined && result.snapshot.revision === before.revision + 1
        ? before.proposals.filter(record => !result.snapshot.proposals.some(next => next.id === record.id)).map(record => record.id) : []);
      for (const id of withdrawn) {
        if (operation.withdrawal?.includes(id)) this.ownWithdrawals.add(id);
      }
    }
    for (const snapshot of operation.snapshots) this.synchronizeProposals(snapshot);
    if (result?.ok) this.synchronizeProposals(result.snapshot);
    this.synchronizeProposals(this.session.getProposals());
    if (result?.ok) for (const change of operation.request?.changes ?? []) {
      if (change.state === 'proposed') this.ownStateVersions.delete(change.id);
    }
    for (const versions of this.ownStateVersions.values()) {
      for (const version of versions) {
        if (versions.size <= 256) break;
        versions.delete(version);
      }
    }
    for (const entries of [this.ownStateVersions, this.ownWithdrawals]) {
      for (const id of entries.keys()) {
        if (entries.size <= 256) break;
        entries.delete(id);
      }
    }
  }

  private forgetProposalOwnership(held: HeldProposal) {
    for (const { id } of held.engine) {
      this.ownStateVersions.delete(id);
      this.ownWithdrawals.delete(id);
    }
  }

  private checkEditorDecision(held: HeldProposal) {
    if (this.attached) this.synchronizeProposals(this.session.getProposals());
    if (held.editorDecisions.size) throw new ProposalDecidedInEditorError([...held.editorDecisions].map(([id, state]) => ({ id, state })));
  }

  private checkDecisionReferences(before: ReturnType<DocxAgentDocument['captureReferences']>, held: HeldProposal) {
    this.checkEditorDecision(held);
    const stale = [...held.targets].filter(([, target]) => {
      const key = JSON.stringify([target.story, target.paraId]);
      return before.decisions.get(key) !== this.editorDecisionRevisions.get(key);
    }).map(([ref]) => ref);
    if (stale.length) this.stale(stale);
  }

  private settledStatus(held: HeldProposal) {
    return held.engine.some(record => record.state === 'accepted') ? 'accepted' as const : 'rejected' as const;
  }

  private observeProposals(snapshot: DocxProposalSnapshot) {
    this.synchronizeProposals(snapshot);
    this.epoch++;
  }

  private proposalDecision(held: HeldProposal, state: 'accepted' | 'rejected') {
    this.synchronizeProposals(this.session.getProposals());
    this.checkEditorDecision(held);
    const snapshot = this.proposalSnapshot ?? this.session.getProposals();
    const ids = new Set(held.engine.map(record => record.id));
    if ([...ids].some(id => !snapshot.proposals.some(record => record.id === id))) {
      throw new DocumentToolError('STALE_TARGET', 'Proposal records are missing from the current snapshot. Read fresh targets.', { id: held.value.id });
    }
    const changes = snapshot.proposals.filter(record => ids.has(record.id) && record.state === 'proposed').map(record => ({ id: record.id, state }));
    return { snapshot, changes, request: { expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion, changes } };
  }

  private proposalRollback(held: HeldProposal, decision: ReturnType<DocxAgentDocument['proposalDecision']>, decided: DocxProposalSnapshot) {
    this.synchronizeProposals(this.session.getProposals());
    const current = this.proposalSnapshot ?? this.session.getProposals();
    return { expectVersion: current.version, expectPreviewVersion: current.previewVersion,
      changes: decision.snapshot.proposals.filter(record => !held.editorDecisions.has(record.id)
        && decision.changes.some(change => change.id === record.id && current.proposals.some(existing => existing.id === record.id
          && existing.state === change.state && existing.stateVersion === decided.proposals.find(value => value.id === record.id)?.stateVersion)))
        .map(({ id, state }) => ({ id, state })),
    };
  }

  private decideProposal(held: HeldProposal, state: 'accepted' | 'rejected', keepRevisions = false) {
    this.checkEditorDecision(held);
    if (this.proposalPeer) throw new DocumentToolError('UNSUPPORTED', 'Use rejectAsync or discardAsync with a DOCX editor edit peer.');
    const before = this.captureReferences();
    const decision = this.proposalDecision(held, state);
    let decided: DocxProposalSnapshot | undefined;
    let settled: DocxProposalSnapshot | undefined;
    try {
      this.proposalOperation = { request: decision.request, snapshots: [] };
      const result = this.session.setProposalStates(decision.request);
      this.finishProposalOperation(result);
      decided = result.ok ? result.snapshot : undefined;
      this.checkEditorDecision(held);
      decided = this.proposalResult(result, held.value.changes);
      this.checkDecisionReferences(before, held);
      const ids = held.engine.map(record => record.id);
      this.proposalOperation = { withdrawal: ids, withdrawalSnapshot: decided, snapshots: [] };
      const withdrawal = this.session.withdrawProposals({ ids, expectVersion: decided.version, expectPreviewVersion: decided.previewVersion, keepRevisions });
      this.finishProposalOperation(withdrawal);
      settled = this.proposalResult(withdrawal, held.value.changes);
      this.preserveReferences(before, settled.version, this.agentDecisionChanges(held, state, keepRevisions));
      this.checkDecisionReferences(before, held);
    } catch (error) {
      if (this.proposalOperation) this.finishProposalOperation();
      if (settled) {
        if (error instanceof DocumentToolError) throw new DocumentToolError(error.code, `The agent decision applied. ${error.message}`, { ...error.details, applied: true, id: held.value.id, status: this.settledStatus(held) });
        throw error;
      }
      if (decided) {
        const rollback = this.proposalRollback(held, decision, decided);
        if (rollback.changes.length) {
          this.proposalOperation = { request: rollback, snapshots: [] };
          try { this.finishProposalOperation(this.session.setProposalStates(rollback)); }
          catch { if (this.proposalOperation) this.finishProposalOperation(); }
          finally { if (this.proposalOperation) this.finishProposalOperation(); }
        }
      }
      this.checkEditorDecision(held);
      throw error;
    }
  }

  private async decideProposalAsync(held: HeldProposal, state: 'accepted' | 'rejected', keepRevisions = false) {
    this.checkEditorDecision(held);
    const peer = this.proposalPeer;
    if (!peer) return this.decideProposal(held, state, keepRevisions);
    const before = this.captureReferences();
    const decision = this.proposalDecision(held, state);
    let decided: DocxProposalSnapshot | undefined;
    let settled: DocxProposalSnapshot | undefined;
    try {
      this.proposalOperation = { request: decision.request, snapshots: [] };
      const result = await peer.setProposalStates(decision.request);
      this.finishProposalOperation(result);
      decided = result.ok ? result.snapshot : undefined;
      this.checkEditorDecision(held);
      decided = this.proposalResult(result, held.value.changes);
      this.checkDecisionReferences(before, held);
      const ids = held.engine.map(record => record.id);
      this.proposalOperation = { withdrawal: ids, withdrawalSnapshot: decided, snapshots: [] };
      const withdrawal = await peer.withdrawProposals({ ids, expectVersion: decided.version, expectPreviewVersion: decided.previewVersion, keepRevisions });
      this.finishProposalOperation(withdrawal);
      settled = this.proposalResult(withdrawal, held.value.changes);
      this.preserveReferences(before, settled.version, this.agentDecisionChanges(held, state, keepRevisions));
      this.checkDecisionReferences(before, held);
    } catch (error) {
      if (this.proposalOperation) this.finishProposalOperation();
      if (settled) {
        if (error instanceof DocumentToolError) throw new DocumentToolError(error.code, `The agent decision applied. ${error.message}`, { ...error.details, applied: true, id: held.value.id, status: this.settledStatus(held) });
        throw error;
      }
      if (decided) {
        const rollback = this.proposalRollback(held, decision, decided);
        if (rollback.changes.length) {
          this.proposalOperation = { request: rollback, snapshots: [] };
          try { this.finishProposalOperation(await peer.setProposalStates(rollback)); }
          catch { if (this.proposalOperation) this.finishProposalOperation(); }
          finally { if (this.proposalOperation) this.finishProposalOperation(); }
        }
      }
      this.checkEditorDecision(held);
      throw error;
    }
  }

  private agentDecisionChanges(held: HeldProposal, state: 'accepted' | 'rejected', keepRevisions: boolean) {
    return state === 'rejected' ? [] : held.value.changes.map(trimChange).map(change => ({ ...change, newText: keepRevisions ? change.oldText + change.newText : change.newText }));
  }

  private previewDecisions(snapshot: DocxProposalSnapshot) {
    const owned = new Set([...this.proposals.values()].flatMap(held => held.engine.map(record => record.id)));
    const decisions = new Map<string, 'accepted' | 'rejected'>();
    for (const record of snapshot.proposals) {
      const state = record.state === 'proposed' ? owned.has(record.id) ? 'rejected' : undefined : record.state;
      if (state) for (const id of record.revisionIds) decisions.set(id, state);
    }
    return decisions;
  }

  private projectSegments(segments: YrsStorySegment[], snapshot: DocxProposalSnapshot): YrsStorySegment[] {
    const decisions = this.previewDecisions(snapshot);
    const projected: YrsStorySegment[] = [];
    for (const segment of segments) {
      const attributes = { ...segment.attributes };
      let hidden = false;
      for (const kind of ['ins', 'del'] as const) {
        const stamp = attributes[kind] as { revisionId?: string; id?: string } | undefined;
        const decision = stamp && decisions.get(stamp.revisionId ?? stamp.id ?? '');
        if (!decision) continue;
        hidden ||= kind === 'ins' ? decision === 'rejected' : decision === 'accepted';
        delete attributes[kind];
      }
      if (hidden) continue;
      const previous = projected.at(-1);
      if (previous?.kind === 'text' && segment.kind === 'text' && JSON.stringify(previous.attributes) === JSON.stringify(attributes)) previous.text += segment.text;
      else projected.push({ ...segment, attributes });
    }
    return projected;
  }

  private acceptedOffset(story: string, paraId: string, offset: number) {
    const decisions = this.previewDecisions(this.proposalSnapshot ?? this.session.getProposals());
    let paragraph: YrsStorySegment[] = [];
    for (const segment of this.session.storySegments(story)) {
      if (segment.kind !== 'pilcrow') { paragraph.push(segment); continue; }
      if (segment.paraId !== paraId) { paragraph = []; continue; }
      let visible = 0;
      let accepted = 0;
      let leading = true;
      for (const run of paragraph) {
        const length = run.kind === 'text' ? run.text.length : 1;
        const stamp = (run.attributes.ins ?? run.attributes.del) as { revisionId?: string; id?: string } | undefined;
        const decision = stamp && decisions.get(stamp.revisionId ?? stamp.id ?? '');
        const shown = decision === undefined || (run.attributes.ins ? decision === 'accepted' : decision === 'rejected');
        const block = leading && run.kind === 'embed' && BLOCK_EMBEDS.has(run.embedKind);
        const inAccepted = !block && !run.attributes.del;
        if (!block) leading = false;
        if (shown && visible + length >= offset) return accepted + (inAccepted ? offset - visible : 0);
        if (shown) visible += length;
        if (inAccepted) accepted += length;
      }
      return accepted;
    }
    throw new DocumentToolError('STALE_TARGET', 'Target paragraph changed. Read fresh targets.');
  }

  private projectProposals(draft: YrsSession, selected?: HeldProposal) {
    const snapshot = this.proposalSnapshot ?? this.session.getProposals();
    const decisions = this.previewDecisions(snapshot);
    if (selected) for (const record of selected.engine.filter(record => record.state === 'proposed')) for (const id of record.revisionIds) decisions.set(id, 'accepted');
    for (const [revisionId, state] of decisions) {
      if (state === 'accepted') draft.acceptChange({ revisionId });
      else draft.rejectChange({ revisionId });
    }
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
      const release = cleanupGuard(() => draft.destroy());
      try {
        if (this.epoch !== version) throw new DocumentToolError('STALE_VERSION', 'The document changed while the batch started. Read the version again.');
        const vector = draft.encodeStateVector();
        const outcomes: Array<ReturnType<EditStep>> = [];
        let resolved: Map<string, boolean> | undefined;
        let deleted: Set<string> | undefined;
        let commentIdentities: Map<string, DocxCommentSaveIdentity> | undefined;
        let commentBaseline: DocxCommentSaveBaseline | undefined;
        let resetCommentBaseline = false;
        const doneBefore = edits.some(edit => edit.op === 'resolve_comment') ? new Map(draft.listComments().map(comment => [comment.id, comment.done])) : undefined;
        for (const [index, step] of steps.entries()) { failed = index; outcomes.push(step(draft)); }
        failed = undefined;
        if (edits.some(edit => edit.op === 'add_comment' || edit.op === 'reply_comment' || edit.op === 'resolve_comment' || edit.op === 'replace_comment_text' || edit.op === 'delete_comment')) {
          failed = edits.findIndex(edit => edit.op === 'add_comment' || edit.op === 'reply_comment' || edit.op === 'resolve_comment' || edit.op === 'replace_comment_text' || edit.op === 'delete_comment');
          try {
            const deletingComments = edits.some(edit => edit.op === 'delete_comment');
            const source = await this.checkCommentPart(deletingComments);
            if (!source) throw new DocumentToolError('UNSUPPORTED_TARGET', 'Comment export requires a session opened from DOCX bytes.');
            const seeded = new Set((draft.materializeDocx()?.package.document.comments ?? []).map(comment => String(comment.id)));
            const comments = draft.listComments();
            const batch = new Set(edits.flatMap((edit, index) => edit.op === 'resolve_comment' ? [outcomes[index].result.comment!] : []));
            resolved = new Map(this.resolvedComments);
            for (const id of batch) if (!resolved.has(id)) resolved.set(id, doneBefore?.get(id) ?? false);
            const earlier = this.resolvedComments;
            const remaining = new Set(comments.map(comment => comment.id));
            deleted = new Set([...this.deletedComments, ...outcomes.flatMap(outcome => outcome.result.deleted ?? [])].filter(id => !remaining.has(id)));
            const previous = this.session.listComments();
            let baseline = await this.commentBaselineBytes(version, deletingComments);
            const identities = deletingComments || this.deletedComments.size
              ? await readSavedCommentIdentities(source, baseline, previous, this.commentSaveIdentities) : undefined;
            let saved: Uint8Array;
            if (deletingComments) {
              if (edits.some(edit => edit.op !== 'delete_comment')) {
                const kept = await this.fork(draft.clientId);
                try {
                  for (const [index, step] of steps.entries()) if (edits[index].op !== 'delete_comment') step(kept);
                  baseline = (await engines.get('docx').module.yrs.saveYrsDocx(kept, { updateModifiedDate: false, commentIdentities: identities,
                    commentBaseline: { bytes: baseline, comments: previous },
                  })).bytes;
                } finally { kept.destroy(); }
              }
              saved = (await engines.get('docx').module.yrs.saveYrsDocx(draft, { updateModifiedDate: false, commentIdentities: identities,
                commentBaseline: { bytes: baseline, comments: previous }, commentDeletion: true,
              })).bytes;
              commentBaseline = { bytes: saved, comments };
            } else {
              saved = (await engines.get('docx').module.yrs.saveYrsDocx(draft, identities ? { commentIdentities: identities } : undefined)).bytes;
              resetCommentBaseline = true;
            }
            await verifyCommentOutput(source, saved, comments.filter(comment => !seeded.has(comment.id)).length,
              comments.map(comment => ({ id: comment.id, parentId: comment.parentId, ...(batch.has(comment.id) || (earlier.get(comment.id) ?? comment.done) !== comment.done ? { done: comment.done } : {}) })),
              new Set([...this.textEditedComments, ...edits.flatMap(edit => edit.op === 'replace_comment_text' ? [edit.comment] : [])]), baseline, deleted, deletingComments ? previous : undefined, identities);
            if (identities) commentIdentities = await readSavedCommentIdentities(source, saved, comments, identities);
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
        for (const edit of edits) if (edit.op === 'replace_comment_text') this.textEditedComments.add(edit.comment);
        if (deleted) this.deletedComments = deleted;
        if (commentIdentities) this.commentSaveIdentities = commentIdentities;
        if (commentBaseline) {
          this.commentSaveBaseline = commentBaseline;
          this.commentSaveBaselineVersion = this.epoch;
        } else if (resetCommentBaseline) {
          this.commentSaveBaseline = undefined;
          this.commentSaveBaselineVersion = undefined;
        }
        return outcomes;
      } catch (error) { throw release.failed(error); } finally { release.run(); }
    } catch (error) {
      const failure = error instanceof DocumentToolError ? error
        : error instanceof Error && (error as Error & { code?: unknown }).code === 'table-unavailable'
          ? new DocumentToolError('TABLE_UNAVAILABLE', error.message)
          : new DocumentToolError('INVALID_EDIT', error instanceof Error ? error.message : String(error));
      throw new DocumentToolError(failure.code, `${failure.message} No batch edits were applied.`, { ...failure.details, ...(failed === undefined ? {} : { failedIndex: failed }) });
    }
  }

  private async plan(edit: DocxAgentEdit): Promise<EditStep> {
    if (edit.op === 'set_runs') {
      const p = this.get(edit.paragraph);
      if (this.runRevision(p) !== edit.revision) throw new DocumentToolError('STALE_TARGET', 'The paragraph changed. Read its runs again.', { refs: [p.ref] });
      if (new Set(edit.runs.map(run => run.run)).size !== edit.runs.length) throw new DocumentToolError('INVALID_ARGUMENT', 'Run IDs must be unique within set_runs.');
      const ids = new Map(p.runs.map((run, index) => [`${p.ref}.${index + 1}`, run]));
      const replacements = edit.runs.map(({ run: id, text }) => {
        const run = ids.get(id);
        if (!run) throw new DocumentToolError('UNKNOWN_REF', 'Unknown run ID. Copy IDs from docx_read_runs for this paragraph.', { run: id });
        if (fixedRunKind(run)) throw new DocumentToolError('PROTECTED_RUN', 'Fixed tokens cannot be changed.', { run: id });
        lineText(text);
        return { start: this.acceptedOffset(p.story, p.paraId, run.start), end: this.acceptedOffset(p.story, p.paraId, run.end), oldText: p.text.slice(run.start, run.end), text };
      }).sort((a, b) => b.start - a.start);
      return draft => {
        requireParagraph(draft, p);
        for (const run of replacements) {
          if (run.oldText === run.text) continue;
          hostEdit(draft, { op: 'replaceText', target: { kind: 'range', story: p.story, start: { paraId: p.paraId, offset: run.start }, end: { paraId: p.paraId, offset: run.end }, view: 'accepted' }, text: run.text, expect: { text: run.oldText } });
        }
        return { result: { ref: p.ref } };
      };
    }
    if (edit.op === 'delete_comment') {
      await this.checkCommentPart(true);
      return draft => {
        const comments = draft.listComments();
        const target = comments.find(comment => comment.id === edit.comment);
        if (!target) throw new DocumentToolError('UNKNOWN_REF', "Unknown comment ID. List office_outline with story 'comments'.");
        const deleted = new Set([target.id]);
        if (target.parentId === null) {
          let size: number;
          do {
            size = deleted.size;
            for (const comment of comments) if (comment.parentId !== null && deleted.has(comment.parentId)) deleted.add(comment.id);
          } while (deleted.size !== size);
        }
        draft.applyRawOps('body', [...deleted].map(id => ({ op: 'removeComment', id })));
        return { result: { comment: target.id, deleted: [...deleted] } };
      };
    }
    if (edit.op === 'set_section') {
      const section = this.session.listSections()[edit.section - 1];
      if (!section) throw new DocumentToolError('INVALID_ARGUMENT', 'Section missing. Read office_outline and use its section numbers.');
      const { orientation, pageWidth, pageHeight, margins } = edit;
      return draft => {
        setSection(draft, section.closingParagraphId, { orientation, pageWidth, pageHeight, margins });
        return { result: {} };
      };
    }
    if (edit.op === 'insert_section_break') {
      const anchor = this.get((edit.before ?? edit.after)!);
      if (anchor.story !== 'body' || anchor.table) throw new DocumentToolError('UNSUPPORTED_TARGET', 'Section breaks require a body paragraph outside a table.');
      const body = this.paragraphs.filter(paragraph => paragraph.story === 'body');
      const target = edit.before ? body[body.findIndex(paragraph => paragraph.paraId === anchor.paraId) - 1] : anchor;
      if (!target) throw new DocumentToolError('UNSUPPORTED_TARGET', 'Cannot insert a section break before the first body paragraph.');
      if (edit.before) {
        const segments = this.session.storySegments('body');
        const leading = segments[segments.findIndex(segment => segment.kind === 'pilcrow' && segment.paraId === target.paraId) + 1];
        if (leading?.kind === 'embed' && BLOCK_EMBEDS.has(leading.embedKind)) throw new DocumentToolError('UNSUPPORTED_TARGET', 'A table or other block precedes this paragraph, so a break before it would land ahead of that block. Use after with the paragraph that should close the section.');
      }
      return draft => {
        insertSectionBreak(draft, target, edit.kind);
        return { result: { ref: target.ref } };
      };
    }
    if (edit.op === 'replace_comment_text') {
      await this.checkCommentPart();
      return draft => ({ result: { comment: edit.comment, replacements: replaceCommentText(draft, edit.comment, edit.find, edit.replace) } });
    }
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
    if (edit.op === 'set_language' && edit.match === undefined) {
      const slots = { lang: edit.lang, eastAsia: edit.eastAsia, bidi: edit.bidi };
      if (edit.style !== undefined || edit.defaults) {
        if (edit.defaults && !this.session.materializeDocx()?.package.styles) {
          throw new DocumentToolError('UNSUPPORTED', 'The styles part word/styles.xml is missing; document-default language cannot be changed.');
        }
        const styleId = edit.style !== undefined ? resolveStyle(this.languageStyles(), edit.style) : null;
        return draft => {
          draft.setStyleLanguage(styleId, slots);
          return { result: styleId === null ? {} : { style: styleId } };
        };
      }
      const p = this.get(edit.paragraph!);
      return draft => {
        requireParagraph(draft, p);
        const keptUnchanged = formatParagraph(draft, p, { language: slots });
        return { result: { ref: p.ref, ...(Object.keys(keptUnchanged).length > 0 ? { keptUnchanged } : {}) } };
      };
    }
    if (edit.op === 'format_text' || edit.op === 'add_comment' || edit.op === 'set_language') {
      const held = this.matches.get(edit.match!);
      if (!held) throw unknownMatchError(edit.match!, 'm1', /^m\d+$/u, 'Match expired or belongs to another document. Run grep again and copy its match ID.');
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
      const delta = edit.op === 'set_language'
        ? { language: { lang: edit.lang, eastAsia: edit.eastAsia, bidi: edit.bidi } }
        : formatTextDelta(edit);
      if (Object.values(delta).every(value => value === undefined)) throw new DocumentToolError('INVALID_ARGUMENT', 'format_text needs a mark.');
      return draft => {
        requireParagraph(draft, p);
        const { keptUnchanged } = draft.formatRange(range, delta);
        return { result: { ref: p.ref, start, end, ...(Object.keys(keptUnchanged).length > 0 ? { keptUnchanged } : {}) } };
      };
    }
    const p = this.get(edit.op === 'insert_paragraphs' ? (edit.before ?? edit.after)! : edit.op === 'set_paragraph_style' || edit.op === 'delete_paragraph' ? edit.paragraph : edit.op === 'merge_cells' ? edit.from : edit.row);
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
        if (edit.before !== undefined && edit.list) {
          throw new DocumentToolError('INVALID_ARGUMENT', 'list: continue needs after.');
        }
        if (edit.list && edit.runFormat !== undefined) {
          throw new DocumentToolError('INVALID_ARGUMENT', 'runFormat cannot be combined with list: continue. Omit runFormat.');
        }
        if (edit.list && edit.paragraphs.some(({ style }) => style)) {
          throw new DocumentToolError('INVALID_ARGUMENT', 'list: continue takes the style of the anchor list item. Omit style.');
        }
        const paragraphs = edit.paragraphs.map(({ text, style }) => {
          lineText(text);
          const resolved = style ? resolveStyle(this.styles(), style) : undefined;
          return edit.before !== undefined ? (resolved ? { text, styleId: resolved } : { text }) : { text, styleId: resolved ?? fallback };
        });
        if (edit.list) {
          const stories = new Map(this.session.listStories().stories.map(story => [story.story, story]));
          for (let story = stories.get(p.story); story; story = story.parent ? stories.get(story.parent) : undefined) {
            if (story.kind === 'contentControl') throw new DocumentToolError('UNSUPPORTED_TARGET', 'Lists inside content controls cannot be continued yet. Use a list item outside the content control.');
          }
        }
        return draft => {
          requireParagraph(draft, target);
          if (edit.before !== undefined) return { result: {}, created: insertParagraphsBefore(draft, target, paragraphs, edit.runFormat) };
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
      case 'merge_cells': {
        const table = p.table ?? notInTable();
        const corner = this.get(edit.to);
        const other = corner.table ?? notInTable();
        if (other.story !== table.story || other.index !== table.index) throw new DocumentToolError('UNSUPPORTED_TARGET', 'from and to are in different tables. Merge cells of one table.');
        const end = { story: corner.story, paraId: corner.paraId };
        return draft => {
          requireParagraph(draft, target);
          requireParagraph(draft, end);
          const { rowSpan, colSpan, paragraphs } = mergeCells(draft, target, end, table.story);
          return { result: { rowSpan, colSpan }, created: paragraphs };
        };
      }
    }
  }

  private async checkCommentPart(allowLegacy = false): Promise<CommentSource | undefined> {
    const source = this.session.materializeDocx()?.originalBuffer;
    if (!source) return;
    if (!this.commentPartCheck || this.commentPartCheck.source !== source || this.commentPartCheck.allowLegacy !== allowLegacy) this.commentPartCheck = { source, allowLegacy, check: readCommentPart(source, allowLegacy) };
    return this.commentPartCheck.check;
  }

  private commentBaselineBytes(version: number, deletion = false): Promise<Uint8Array> {
    if (this.commentPackageBaseline?.version !== version || this.commentPackageBaseline.deletion !== deletion) {
      this.commentPackageBaseline = { version, deletion, bytes: deletion ? this.exportBytes(undefined, false) : this.export() };
    }
    return this.commentPackageBaseline.bytes;
  }

  private languageStyles(): ParagraphStyle[] {
    return (this.session.materializeDocx()?.package.styles?.styles ?? [])
      .filter(style => style.type === 'paragraph' || style.type === 'character')
      .map(style => ({ id: style.styleId, name: style.name, isDefault: !!style.default }));
  }

  private styles(): ParagraphStyle[] {
    return this.paragraphStyles ??= (this.session.materializeDocx()?.package.styles?.styles ?? [])
      .filter(style => style.type === 'paragraph')
      .map(style => ({ id: style.styleId, name: style.name, isDefault: !!style.default }));
  }
}

export async function openDocx(bytes: Uint8Array, options: DocumentOptions = {}): Promise<DocxAgentDocument> {
  const { module } = await engines.load('docx');
  const session = await module.yrs.createYrsSession();
  try {
    session.seedFromDocx(bytes);
    return new DocxAgentDocument(session, options, true);
  } catch (error) { throw cleanupGuard(() => session.destroy()).failed(error); }
}

async function sourceModelCounts(bytes: Uint8Array): Promise<DocxModelCounts> {
  const { module } = await engines.load('docx');
  const session = await module.yrs.createYrsSession();
  const release = cleanupGuard(() => session.destroy());
  try {
    session.seedFromDocx(bytes);
    return docxModelCounts(session);
  } catch (error) { throw release.failed(error); } finally { release.run(); }
}

export function attachDocx(session: YrsSession, options: DocumentOptions = {}): DocxAgentDocument {
  return new DocxAgentDocument(session, options);
}

/** The session formatting delta a `format_text` edit writes. */
function formatTextDelta(edit: Extract<DocxAgentEdit, { op: 'format_text' }>): YrsInlineFormatDelta {
  const strike = edit.strike === undefined ? undefined
    : edit.strike === 'dblStrike' ? { double: true }
      : edit.strike === true || edit.strike === 'sngStrike';
  return {
    bold: edit.bold, italic: edit.italic, underline: edit.underline, strike,
    color: typeof edit.color === 'string' ? { rgb: edit.color } : edit.color,
    highlight: edit.highlight,
    fontSize: edit.fontSizePt,
    fontFamily: edit.fontFamily === undefined ? undefined : { ascii: edit.fontFamily, hAnsi: edit.fontFamily },
    other: edit.verticalAlign === undefined ? undefined
      : { superscript: edit.verticalAlign === 'superscript' ? true : null, subscript: edit.verticalAlign === 'subscript' ? true : null },
  };
}
