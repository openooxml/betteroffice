/**
 * Document Content Model — barrel.
 *
 * All content-bearing types: runs, hyperlinks, bookmarks, fields,
 * images, shapes, tables, lists, paragraphs, headers/footers,
 * footnotes/endnotes, and sections.
 *
 * The types form a deeply interrelated tree (Paragraph ↔ Table ↔ ShapeTextBody)
 * and are split across `./content/*.ts` by domain. This file re-exports them
 * so existing imports from `@betteroffice/docx/types/content` keep
 * working unchanged.
 * @packageDocumentation
 * @public
 */

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
} from './content/run';

export type {
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,
  FieldType,
  SimpleField,
  ComplexField,
  Field,
} from './content/link';

export type {
  ImageSize,
  ImageWrap,
  ImagePosition,
  ImageTransform,
  ImagePadding,
  ImageCrop,
  Image,
} from './content/image';

export type {
  ShapeType,
  ShapeFill,
  ShapeOutline,
  ShapeTextBody,
  Shape,
  TextBox,
} from './content/shape';

export type { ChartType, ChartSeries, ChartAxis, ChartLegend, Chart } from './content/chart';

export type { TableCell, TableRow, Table } from './content/table';

export type { Comment, CommentRangeStart, CommentRangeEnd } from './content/comment';

export type { MathEquation } from './content/math';

export type {
  TrackedChangeInfo,
  PropertyChangeInfo,
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
  MoveFromRangeStart,
  MoveFromRangeEnd,
  MoveToRangeStart,
  MoveToRangeEnd,
  TrackedRunChange,
  RunPropertyChange,
  ParagraphPropertyChange,
  TablePropertyChange,
  TableRowPropertyChange,
  TableCellPropertyChange,
  TableStructuralChangeInfo,
} from './content/trackedChange';

export type { SdtType, SdtProperties, SdtDataBinding, InlineSdt, BlockSdt } from './content/sdt';

export type { ParagraphContent, Paragraph } from './content/paragraph';

export type {
  HeaderFooterType,
  HeaderReference,
  FooterReference,
  HeaderFooter,
  FootnotePosition,
  EndnotePosition,
  NoteNumberRestart,
  FootnoteProperties,
  EndnoteProperties,
  Footnote,
  Endnote,
} from './content/headerFooter';

export type { TextWatermark, PictureWatermark, Watermark } from './content/watermark';
export { pictureWatermarkDisplayEmu, DEFAULT_WATERMARK_PRESETS } from './content/watermark';

export type {
  PageOrientation,
  SectionStart,
  VerticalAlign,
  LineNumberRestart,
  Column,
  SectionProperties,
  BlockContent,
  Section,
  DocumentBody,
} from './content/section';
