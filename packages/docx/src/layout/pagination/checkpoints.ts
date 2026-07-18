/**
 * Layout checkpoints — page-start bookmarks carried on `Layout.checkpoints`.
 *
 * A checkpoint captures everything a paginator holds when a block starts on a
 * pristine page: the block about to be placed, the section cursor, and the
 * live page-flow geometry (including any continuous-section geometry still
 * queued per ECMA-376 §17.6.22). The field is derived data — fully determined
 * by the same inputs as the layout itself — and remains an optional part of
 * the `Layout` contract (omitted from the golden serialization; see
 * `__golden__/serializeLayout.ts`). The Rust engine performs full passes and
 * does not resume from checkpoints.
 */

import type {
  ColumnLayout,
  NoteSettingsContract,
  PageBorderContract,
  PageHeaderFooterRefs,
  PageMargins,
  PageNumberingContract,
  SectionBreakBlock,
} from './types';

/** Stable section/page state copied onto every page created in a section. */
export type PageSectionState = {
  emitPageMetadata?: boolean;
  sectionIndex?: number;
  sectionId?: string;
  orientation?: 'portrait' | 'landscape';
  headerFooterRefs?: PageHeaderFooterRefs;
  pageNumbering?: PageNumberingContract;
  pageBorders?: PageBorderContract;
  watermark?: SectionBreakBlock['watermark'];
  verticalAlign?: SectionBreakBlock['verticalAlign'];
  noteSettings?: NoteSettingsContract;
};

/**
 * The page-flow geometry a layout checkpoint captures: everything a
 * paginator's closure holds that survives across pages, including geometry
 * queued by a continuous section break (ECMA-376 §17.6.22) that has not been
 * applied yet. Plain data — checkpoints ride on `Layout`.
 */
export type PageFlowGeometry = {
  pageSize: { w: number; h: number };
  margins: PageMargins;
  columns: ColumnLayout;
  pendingPageSize?: { w: number; h: number };
  pendingMargins?: PageMargins;
  section?: PageSectionState;
  pendingSection?: PageSectionState;
  sectionPageIndex?: number;
  lastLogicalPageNumber?: number;
  lastCreatedSectionKey?: string;
};

export interface LayoutCheckpoint {
  /** Index of the measured block about to be placed on the pristine page. */
  blockIndex: number;
  /** The place walk's section cursor at that point. */
  sectionIdx: number;
  /** Index into `Layout.pages` of the page being started. */
  pageIndex: number;
  /** 1-based page number of that page. */
  pageNumber: number;
  /** Live paginator geometry at the moment the page was created. */
  flow: PageFlowGeometry;
}
