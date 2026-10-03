import type { YrsSession } from '@betteroffice/docx/yrs';

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
}

export interface ParagraphRead extends ParagraphRef {
  text: string;
  start: number;
  end: number;
  length: number;
  nextStart: number | null;
  style: string | null;
  runs: TextRun[];
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

export interface DocumentOptions {
  name?: string;
  renderer?: DocumentRenderer;
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
