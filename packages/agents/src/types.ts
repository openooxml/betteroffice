import type { DocxProposalRecord, DocxProposalRequest, DocxProposalResult, DocxProposalStateRequest, DocxProposalWithdrawRequest, YrsRevisionInfo, YrsSession } from '@betteroffice/docx/yrs';

export interface ParagraphRef {
  ref: string;
  revision: number;
  story: string;
  paragraph: number;
}

export interface TextRun {
  start: number;
  end: number;
  formatting: Record<string, unknown>;
  protected: boolean;
  kind?: string;
}

export interface DocxRunPage {
  version: number;
  paragraphs: Array<{ ref: string; revision: number; text: string; runs: Array<{ id: string; text: string; f?: string } | { id: string; kind: string }> }>;
  formats: Record<string, Record<string, unknown>>;
  nextOffset: number | null;
  total: number;
}

/** A table cell paragraph's place: parent story, one-based table index, row and column, the table's size, and the cell's spans when merged. */
export interface TableCellPosition {
  story: string;
  index: number;
  row: number;
  column: number;
  rows: number;
  columns: number;
  rowSpan?: number;
  colSpan?: number;
}

export interface ParagraphRead extends ParagraphRef {
  text: string;
  start: number;
  end: number;
  length: number;
  nextStart: number | null;
  style: string | null;
  runs: TextRun[];
  table?: TableCellPosition;
}

export interface GrepOptions {
  query: string;
  caseSensitive?: boolean;
  story?: string;
  limit?: number;
  cursor?: string;
}

export interface GrepMatch extends ParagraphRef {
  match: string;
  start: number;
  end: number;
  text: string;
  context: string;
  contextStart: number;
}

export interface ExactTextEdit {
  ref: string;
  revision: number;
  oldText: string;
  newText: string;
  start?: number;
}

export type TextEdit = ExactTextEdit | { match: string; newText: string };

export interface ProposalChange extends ParagraphRef {
  start: number;
  oldText: string;
  newText: string;
}

export interface Proposal {
  id: string;
  author: string;
  note: string;
  status: 'pending' | 'accepted' | 'rejected';
  changes: ProposalChange[];
  staleRefs: string[];
  records?: DocxProposalRecordStatus[];
}

export interface RenderedPage {
  png: Uint8Array;
  page: number;
  pageCount: number;
  width: number;
  height: number;
  warnings: string[];
}

export type DocumentRenderer = (session: YrsSession, page: number) => Promise<RenderedPage>;

export interface DocxProposalPeer {
  proposeChanges(request: DocxProposalRequest): Promise<DocxProposalResult>;
  setProposalStates(request: DocxProposalStateRequest): Promise<DocxProposalResult>;
  withdrawProposals(request: DocxProposalWithdrawRequest): Promise<DocxProposalResult>;
}

export interface DocumentOptions {
  name?: string;
  maxEdits?: number;
  renderer?: DocumentRenderer;
  editPeer?: DocxProposalPeer;
}

export function unknownMatchError(value: string, example: string, shape: RegExp, expiredMessage: string): DocumentToolError {
  return new DocumentToolError('UNKNOWN_MATCH', typeof value === 'string' && shape.test(value)
    ? expiredMessage
    : `match takes a match ID from office_grep (such as ${example}), not the text to replace. Grep for the text, then copy the match field of the occurrence.`);
}

export class DocumentToolError extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly details?: Record<string, unknown>,
  ) {
    super(message);
    this.name = 'DocumentToolError';
  }
}

export type DocxRevisionConflict = Pick<YrsRevisionInfo, 'revisionId' | 'author' | 'kind' | 'range'> & { ref: string };

export class TrackedRevisionConflictError extends DocumentToolError {
  constructor(readonly conflicts: DocxRevisionConflict[]) {
    super('TRACKED_REVISION_CONFLICT', `Proposal conflicts with tracked revisions: ${conflicts.map(revision => `${revision.revisionId} by ${revision.author} (${revision.kind})`).join(', ')}.`, { conflicts });
    this.name = 'TrackedRevisionConflictError';
  }
}

export class AcceptDateUnsupportedError extends DocumentToolError {
  constructor() {
    super('ACCEPT_DATE_UNSUPPORTED', 'Attached DOCX acceptance cannot restamp revisions. Omit date.');
    this.name = 'AcceptDateUnsupportedError';
  }
}

export type DocxProposalRecordStatus = Pick<DocxProposalRecord, 'id' | 'state'>;

export class ProposalDecidedInEditorError extends DocumentToolError {
  constructor(readonly records: DocxProposalRecordStatus[]) {
    super('PROPOSAL_DECIDED_IN_EDITOR', 'The editor has decided records in this proposal group. Finish the group in the editor.', { records });
    this.name = 'ProposalDecidedInEditorError';
  }
}
