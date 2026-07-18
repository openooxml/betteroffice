/**
 * Comprehensive TypeScript types for full DOCX document representation
 *
 * This barrel file re-exports all types from the split modules.
 * Existing imports from './types/document' continue to work unchanged.
 *
 * Module structure:
 * - colors.ts      — Color primitives, borders, shading
 * - formatting.ts  — Text, paragraph, and table formatting properties
 * - lists.ts       — Numbering and list definitions
 * - content.ts     — Content model (runs, images, shapes, tables, paragraphs, sections)
 * - styles.ts      — Styles, theme, fonts, relationships, media
 * @packageDocumentation
 * @public
 */

import type { Chart } from './content/chart';

// Color & Styling Primitives
export type { ThemeColorSlot, ColorValue, BorderSpec, ShadingProperties } from './colors';

// Text & Paragraph Formatting
export type {
  UnderlineStyle,
  TextEffect,
  EmphasisMark,
  RunFontHint,
  RunLanguage,
  TextFormatting,
  TextModernEffects,
  TabStopAlignment,
  TabLeader,
  TabStop,
  LineSpacingRule,
  ParagraphAlignment,
  ParagraphFormatting,
  TableWidthType,
  TableMeasurement,
  TableBorders,
  CellMargins,
  TableLook,
  TableStyleRegion,
  TableStyleCascade,
  FloatingTableProperties,
  TableFormatting,
  TableRowFormatting,
  ConditionalFormatStyle,
  TableCellFormatting,
} from './formatting';

// Additive Word-parity contract types. They are exported directly from their
// split modules so the legacy `content.ts` barrel remains byte-for-byte stable.
export type {
  ContentPosition,
  BookmarkRange,
  HyperlinkContent,
  StructuredFieldContent,
  FieldInlineContent,
  StructuredFieldTree,
} from './content/link';
export type {
  SdtCheckboxGlyph,
  SdtDateState,
  SdtGalleryState,
  SdtControlState,
} from './content/sdt';
export type { FieldFormData, CommentReferenceContent } from './content/run';
export type {
  ImageWrapPoint,
  ImageRotationBounds,
  ImageEffect,
  ImageHyperlink,
} from './content/image';
export type {
  ShapeFillPaint,
  ShapeStrokePaint,
  ShapeEffect,
  ShapeTextBodyProperties,
  DrawingSceneNode,
  DrawingScene,
} from './content/shape';
export type { ChartGrouping, ChartPoint, ChartPlotGroup } from './content/chart';
export type { CommentAuthor, CommentAnchorRange } from './content/comment';
export type { NoteKind, NoteSeparatorReference } from './content/headerFooter';
export type { PageNumberingProperties } from './content/section';

// Lists & Numbering
export type {
  NumberFormat,
  LevelSuffix,
  ListLevel,
  AbstractNumbering,
  NumberingInstance,
  ListRendering,
  NumberingDefinitions,
} from './lists';

// Content Model
export type {
  TextContent,
  TabContent,
  BreakContent,
  SymbolContent,
  NoteReferenceContent,
  NoteRefMarkContent,
  SeparatorContent,
  FieldCharContent,
  InstrTextContent,
  SoftHyphenContent,
  NoBreakHyphenContent,
  DrawingContent,
  ShapeContent,
  ChartContent,
  RunContent,
  Run,
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,
  FieldType,
  SimpleField,
  ComplexField,
  Field,
  ImageSize,
  ImageWrap,
  ImagePosition,
  ImageTransform,
  ImagePadding,
  ImageCrop,
  Image,
  ShapeType,
  ShapeFill,
  ShapeOutline,
  ShapeTextBody,
  Shape,
  TextBox,
  ChartType,
  ChartSeries,
  ChartAxis,
  ChartLegend,
  Chart,
  TableCell,
  TableRow,
  Table,
  Comment,
  CommentRangeStart,
  CommentRangeEnd,
  MathEquation,
  TrackedChangeInfo,
  TrackedRunChange,
  PropertyChangeInfo,
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
  MoveFromRangeStart,
  MoveFromRangeEnd,
  MoveToRangeStart,
  MoveToRangeEnd,
  RunPropertyChange,
  ParagraphPropertyChange,
  TablePropertyChange,
  TableRowPropertyChange,
  TableCellPropertyChange,
  TableStructuralChangeInfo,
  SdtType,
  SdtProperties,
  SdtDataBinding,
  InlineSdt,
  BlockSdt,
  ParagraphContent,
  Paragraph,
  HeaderFooterType,
  HeaderReference,
  FooterReference,
  HeaderFooter,
  TextWatermark,
  PictureWatermark,
  Watermark,
  FootnotePosition,
  EndnotePosition,
  NoteNumberRestart,
  FootnoteProperties,
  EndnoteProperties,
  Footnote,
  Endnote,
  PageOrientation,
  SectionStart,
  VerticalAlign,
  LineNumberRestart,
  Column,
  SectionProperties,
  BlockContent,
  Section,
  DocumentBody,
} from './content';

export { pictureWatermarkDisplayEmu, DEFAULT_WATERMARK_PRESETS } from './content';

// Styles, Theme, Fonts, Relationships & Media
export type {
  StyleType,
  Style,
  DocDefaults,
  StyleDefinitions,
  ThemeColorScheme,
  ThemeFont,
  ThemeFontScheme,
  Theme,
  FontInfo,
  FontTable,
  FontEmbed,
  RelationshipType,
  Relationship,
  RelationshipMap,
  MediaFile,
} from './styles';

// ============================================================================
// DOCX PACKAGE & TOP-LEVEL DOCUMENT
// ============================================================================

import type { DocumentBody } from './content';
import type { StyleDefinitions, Theme, FontTable, RelationshipMap, MediaFile } from './styles';
import type { NumberingDefinitions } from './lists';
import type { Footnote, Endnote, HeaderFooter } from './content';
import type { BookmarkRange } from './content/link';
import type { ColorValue } from './colors';
import type { DocumentSettings } from '../docx/settingsParser';

export type { DocumentSettings, CompatibilityFlags } from '../docx/settingsParser';

/**
 * Complete DOCX package structure
 */
export interface DocxPackage {
  /** Model contract version. Undefined reads as legacy version 0. */
  contractVersion?: number;
  /** Document body */
  document: DocumentBody;
  /** Style definitions */
  styles?: StyleDefinitions;
  /** Theme */
  theme?: Theme;
  /** Numbering definitions */
  numbering?: NumberingDefinitions;
  /** Document-wide settings from `word/settings.xml` */
  settings?: DocumentSettings;
  /** Font table */
  fontTable?: FontTable;
  /** Footnotes (normal notes only — separators live in `footnoteSeparators`) */
  footnotes?: Footnote[];
  /** Endnotes (normal notes only — separators live in `endnoteSeparators`) */
  endnotes?: Endnote[];
  /**
   * Separator footnotes (`w:type="separator"` / `"continuationSeparator"` /
   * `"continuationNotice"`) kept out of `footnotes` so rendering/layout only
   * sees real notes. Retained for round-trip: Word rejects a footnotes part
   * whose separator notes are missing, so the serializer re-emits these ahead
   * of the normal notes.
   */
  footnoteSeparators?: Footnote[];
  /** Separator endnotes — see `footnoteSeparators`. */
  endnoteSeparators?: Endnote[];
  /** Headers by relationship ID */
  headers?: Map<string, HeaderFooter>;
  /** Footers by relationship ID */
  footers?: Map<string, HeaderFooter>;
  /** Document relationships */
  relationships?: RelationshipMap;
  /** Media files */
  media?: Map<string, MediaFile>;
  /** Parsed DrawingML chart parts keyed by normalized package path. */
  charts?: Map<string, Chart>;
  /** Exact document-wide bookmark ranges. Undefined = derive from markers. */
  bookmarkRanges?: BookmarkRange[];
  /** Document page background (`w:document/w:background`). Undefined = transparent/white host default. */
  background?: {
    color?: ColorValue;
    themeTint?: string;
    themeShade?: string;
  };
  /** Document properties */
  properties?: {
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string;
    description?: string;
    lastModifiedBy?: string;
    revision?: number;
    created?: Date;
    modified?: Date;
  };
}

/**
 * Top-level parsed DOCX document — the result of `parseDocx(buffer)`.
 *
 * Wraps the unzipped DOCX package (`document.xml`, `styles.xml`, etc.),
 * the original buffer for round-trip saves, and any template variables /
 * parse warnings detected during ingestion.
 *
 * @example
 * ```ts
 * import { parseDocx } from '@betteroffice/docx/headless';
 * const doc = await parseDocx(buffer);
 * console.log(doc.package.document.content.length);
 * ```
 */
export interface Document {
  /** Top-level model contract version. Undefined reads as legacy version 0. */
  contractVersion?: number;
  /** Parsed DOCX package — body, styles, numbering, theme, media, headers/footers. */
  package: DocxPackage;
  /** Original DOCX buffer. Kept for round-trip saves that preserve untouched parts. */
  originalBuffer?: ArrayBuffer;
  /** Detected docxtemplater variables (e.g. `{name}`, `{address}`). Populated when the document is recognized as a template. */
  templateVariables?: string[];
  /** Non-fatal parser diagnostics — malformed parts, unsupported features, fallbacks. */
  warnings?: string[];
}
