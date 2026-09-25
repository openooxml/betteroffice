/**
 * Page fragments of a structured export: the physical page, region and story occurrence showing
 * each exported block and inline, read from the layout the export was captured with. Plain JSON
 * produced by the Rust engine; ordinary exports never lay anything out.
 *
 * Pages are counted physically from zero, blank parity fillers included; the displayed number and
 * label follow the section's PAGE numbering and may restart, repeat or use Roman and letter
 * formats. A header or footer part is exported once and has an occurrence on every page that
 * shows it. Text slices use the export's paragraph-local ranges; an atom (field, control, image,
 * note mark, break) is sliced whole, or marked `partial` when only part of its display content is
 * on the page. Geometry, when requested, is in unzoomed CSS pixels (96 per inch) from the
 * physical page's top-left corner.
 */

import type { DocxTextRange } from './edits';
import type { DocxAnchor } from './readTypes';
import type {
  DocxExportOptions,
  DocxMarkdownOptions,
  DocxRevisionView,
  DocxStructuredContent,
} from './structuredExport';

/**
 * A paged export's options: the export's, plus `includeGeometry` (default `false`),
 * `expectLayoutVersion` (refused as `stale-layout` when the retained layout is another),
 * `maxFragments` (default 100,000, at most 1,000,000) and `maxLayoutBytes` (the page map's own
 * budget: default 8,388,608, 1,024 to 67,108,864).
 */
export interface DocxPageExportOptions extends DocxExportOptions {
  includeGeometry?: boolean;
  expectLayoutVersion?: string;
  maxFragments?: number;
  maxLayoutBytes?: number;
}

/** Structured content with the page map captured with it. */
export interface DocxPagedStructuredContent<L> {
  structured: DocxStructuredContent;
  layout: L;
}

/** Markdown options for a paged export; `pageMarkers` defaults to `false`. */
export interface DocxPageMarkdownOptions extends DocxMarkdownOptions {
  pageMarkers?: boolean;
}

/**
 * The page map of a session export. Its references describe the authoritative layout of
 * `documentVersion` computed from the inputs `provenance` fingerprints; a later edit may
 * supersede that layout before it is ever painted.
 */
export interface DocxLayoutMap extends DocxPageMapContent {
  documentVersion: string;
  layoutVersion: string;
  provenance: DocxLayoutProvenance;
}

/** The page map of a bytes export: deterministic, with no session token. */
export interface DocxSnapshotLayoutMap extends DocxPageMapContent {
  snapshotFingerprint: string;
  provenance: DocxSnapshotLayoutProvenance;
}

/** What both page maps hold. `exportFingerprint` identifies the structured content. */
export interface DocxPageMapContent {
  exportFingerprint: string;
  revisionView: DocxRevisionView;
  /** Pages are laid out with revision markup. */
  layoutRevisionView: 'markup';
  pages: DocxExportPage[];
  occurrences: DocxStoryOccurrence[];
  fragments: DocxPageFragment[];
  diagnostics: DocxPageDiagnostic[];
  truncated: boolean;
}

export interface DocxSnapshotLayoutProvenance {
  engineVersion: string;
  /** The fonts, fallback order and measurement defaults, by font content. */
  fontSetFingerprint: string;
  /** Sections, settings, notes, render environment and layout options. */
  optionsFingerprint: string;
}

export interface DocxLayoutProvenance extends DocxSnapshotLayoutProvenance {
  layoutEpoch: string;
}

export interface DocxExportPage {
  pageIndex: number;
  sectionIndex: number;
  sectionPageIndex: number;
  displayedNumber: number;
  displayedLabel: string;
  numberingFormat: string;
  /** `fallback` when the format could not be written and the page is numbered in decimal. */
  numberingStatus: 'resolved' | 'fallback';
  parityFiller: boolean;
  size: { width: number; height: number };
}

export interface DocxStoryOccurrence {
  id: string;
  pageIndex: number;
  story: string;
  part: string | null;
  sectionIndex: number | null;
  region:
    | { kind: 'body' }
    | { kind: 'header' | 'footer'; variant: 'default' | 'first' | 'even' }
    | {
        kind: 'footnote' | 'endnote';
        noteId: string;
        placement: 'pageBottom' | 'beneathText' | 'sectionEnd' | 'documentEnd';
      };
}

/** The part of one exported node (`nodeId`, in block `blockId`) one occurrence shows. */
export interface DocxPageFragment {
  id: string;
  pageIndex: number;
  occurrenceId: string;
  nodeId: string;
  blockId: string;
  anchor: DocxAnchor;
  slice:
    | { kind: 'text'; range: DocxTextRange }
    | { kind: 'atom'; range: DocxTextRange | null; coverage: 'whole' | 'partial' }
    | { kind: 'block' }
    | {
        kind: 'table';
        rows: Array<{
          rowIndex: number;
          continuedFromPrevious: boolean;
          continuedOnNext: boolean;
          repeatedHeader: boolean;
        }>;
      };
  continuedFromPrevious: boolean;
  continuedOnNext: boolean;
  repeatedTableHeader: boolean;
  geometry: DocxFragmentGeometry | null;
}

export interface DocxFragmentGeometry {
  unit: 'cssPx';
  origin: 'pageTopLeft';
  rects: Array<{ x: number; y: number; width: number; height: number }>;
}

export interface DocxPageDiagnostic {
  code:
    | 'unmapped-content'
    | 'anchor-only'
    | 'not-laid-out'
    | 'unsupported-numbering'
    | 'unsupported-note-layout'
    | 'geometry-unavailable'
    | 'truncated';
  nodeId: string | null;
  pageIndex: number | null;
  message: string;
}

/** One font a headless layout may measure with. */
export interface DocxHeadlessFont {
  /** Names the font in `fontChains` and `defaultChain`. */
  key: string;
  /** TrueType or OpenType bytes, base64-encoded. */
  data: string;
}

/**
 * The fonts and options of a headless layout, plain JSON. Only these fonts are measured with;
 * nothing is looked up on the system. A font requirement (`family|bold|italic`, family
 * lowercase) takes the chain named by its key, else by its lowercase family, else
 * `defaultChain`, in fallback order; one left with no font is not laid out, and the export is
 * refused as `layout-unavailable`.
 */
export interface DocxHeadlessLayoutOptions {
  fonts: readonly DocxHeadlessFont[];
  fontChains?: Readonly<Record<string, readonly string[]>>;
  defaultChain: readonly string[];
  /**
   * Text naming no font or size, measured and given its font chain in this family: default
   * `Calibri` at 11 points.
   */
  measurementDefaults?: { fontFamily?: string; fontSize?: number };
  /**
   * How stories render: hidden text laid out (default `false`) and the default tab stop in
   * twips (default the document's).
   */
  renderEnvironment?: { showHiddenText?: boolean; defaultTabStopTwips?: number };
  /** Word layout compatibility options; default the document's settings. */
  compatibility?: { noLeading?: boolean; doNotExpandShiftReturn?: boolean };
}
