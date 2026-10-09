import type { DocxExportStory } from './structuredExport';

/**
 * What a session story holds. Comment bodies live outside the session's stories, so there is no
 * comment kind; `other` is a story the session cannot place, such as a header or footer story
 * whose part it does not know.
 */
export type DocxStoryInfoKind =
  | 'body'
  | 'table-cell'
  | 'content-control'
  | 'header'
  | 'footer'
  | 'footnote'
  | 'endnote'
  | 'other';

/** A section that references a header or footer part, as the structured export reports it. */
export type DocxStoryUse = DocxExportStory['uses'][number];

export interface DocxStoryInfo {
  story: string;
  kind: DocxStoryInfoKind;
  /** The story a table cell or block content control sits in. */
  parent?: string;
  /** The top-level story this one belongs to; itself for a top-level story. */
  root: string;
  /** A header's or footer's package part, when the session was opened from DOCX bytes. */
  part?: string;
  /**
   * A header's or footer's sections whose properties reference the part, inheritance applied,
   * whether or not a page shows it; plugin anchor geometry tells where it is painted. A
   * first-page preview reads only the sections it holds, so its uses may be incomplete.
   */
  uses?: DocxStoryUse[];
}

/** Every story of the session at one version, sorted by id. */
export interface DocxListStoriesResult {
  ok: true;
  version: string;
  stories: DocxStoryInfo[];
}
