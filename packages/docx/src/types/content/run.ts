/**
 * Run content (`w:r`) and the inline pieces that live inside a run —
 * text, tab, break, symbol, footnote/endnote references, field chars,
 * instruction text, soft/no-break hyphens, drawings, shapes.
 */

import type { TextFormatting } from '../formatting';
import type { Image } from './image';
import type { Shape } from './shape';
import type { Chart } from './chart';
import type { RunPropertyChange } from './trackedChange';

/**
 * Plain text run content (`w:t`). `preserveSpace` mirrors the
 * `xml:space="preserve"` attribute and matters for runs that begin or end
 * with whitespace — without it, Word collapses leading/trailing spaces.
 */
export interface TextContent {
  type: 'text';
  /** The text string */
  text: string;
  /** Preserve whitespace (xml:space="preserve") */
  preserveSpace?: boolean;
}

/**
 * Tab character
 */
export interface TabContent {
  type: 'tab';
}

/**
 * Line break
 */
export interface BreakContent {
  type: 'break';
  /** Break type */
  breakType?: 'page' | 'column' | 'textWrapping';
  /** Clear type for text wrapping break */
  clear?: 'none' | 'left' | 'right' | 'all';
}

/**
 * Symbol character (special font character)
 */
export interface SymbolContent {
  type: 'symbol';
  /** Font name */
  font: string;
  /** Character code */
  char: string;
}

/**
 * Footnote or endnote reference
 */
export interface NoteReferenceContent {
  type: 'footnoteRef' | 'endnoteRef';
  /** Note ID */
  id: number;
  /** `w:customMarkFollows`; undefined = false/automatic number. */
  customMarkFollows?: boolean;
}

/**
 * Footnote/endnote auto-number mark (`w:footnoteRef` / `w:endnoteRef`).
 *
 * Distinct from {@link NoteReferenceContent}: that is the *reference* placed in
 * the document body (`w:footnoteReference`), whereas this is the numbering
 * placeholder that lives *inside* the note body — the run carrying it is what
 * Word renders as the note's leading superscript number. Preserving it keeps
 * the note's own number visible on round-trip.
 */
export interface NoteRefMarkContent {
  type: 'footnoteRefMark' | 'endnoteRefMark';
}

/**
 * Footnote/endnote separator mark (`w:separator` / `w:continuationSeparator`).
 *
 * These appear inside the special separator notes (`w:type="separator"` and
 * `w:type="continuationSeparator"`) and draw the horizontal rule Word places
 * between the body and its notes. They carry no content; Word rejects a notes
 * part whose separator notes have lost these markers, so they must round-trip.
 */
export interface SeparatorContent {
  type: 'separator' | 'continuationSeparator';
}

/** Legacy `w:ffData` projected as inert form-control state. */
export interface FieldFormData {
  name?: string;
  enabled?: boolean;
  helpText?: string;
  statusText?: string;
  entryMacro?: string;
  exitMacro?: string;
  calculateOnExit?: boolean;
  controlType?: 'text' | 'checkbox' | 'dropDown';
  defaultValue?: string;
  textType?: string;
  textFormat?: string;
  checked?: boolean;
  defaultChecked?: boolean;
  checkboxSize?: number;
  listEntries?: string[];
  selectedIndex?: number;
}

/**
 * Field character (begin/separate/end)
 */
export interface FieldCharContent {
  type: 'fieldChar';
  /** Field character type */
  charType: 'begin' | 'separate' | 'end';
  /** Field is locked */
  fldLock?: boolean;
  /** Field is dirty (needs update) */
  dirty?: boolean;
  /** Legacy form data. Undefined = no form control. Macros remain inert. */
  formData?: FieldFormData;
}

/** Zero-width `w:commentReference` anchor for point/malformed comments. */
export interface CommentReferenceContent {
  type: 'commentReference';
  /** Comment id. Undefined = malformed/unresolved reference. */
  id?: number;
}

/**
 * Field instruction text
 */
export interface InstrTextContent {
  type: 'instrText';
  /** Field instruction */
  text: string;
}

/**
 * Soft hyphen
 */
export interface SoftHyphenContent {
  type: 'softHyphen';
}

/**
 * Non-breaking hyphen
 */
export interface NoBreakHyphenContent {
  type: 'noBreakHyphen';
}

/**
 * Drawing/image reference
 */
export interface DrawingContent {
  type: 'drawing';
  /** Image data */
  image: Image;
}

/**
 * Shape reference
 */
export interface ShapeContent {
  type: 'shape';
  /** Shape data */
  shape: Shape;
}

/**
 * DrawingML chart reference.
 */
export interface ChartContent {
  type: 'chart';
  /** Parsed chart data */
  chart: Chart;
}

/**
 * All possible run content types
 */
export type RunContent =
  | TextContent
  | TabContent
  | BreakContent
  | SymbolContent
  | NoteReferenceContent
  | NoteRefMarkContent
  | SeparatorContent
  | CommentReferenceContent
  | FieldCharContent
  | InstrTextContent
  | SoftHyphenContent
  | NoBreakHyphenContent
  | DrawingContent
  | ShapeContent
  | ChartContent;

/**
 * A run (`w:r`) — a contiguous span of inline content sharing one set of
 * character properties (bold, italic, font, color, etc.). Runs are the
 * atomic unit of character formatting; toggling bold on a selection that
 * spans different formatting creates new runs.
 *
 * See ECMA-376 §17.3.2.
 *
 * @example
 * ```ts
 * const run: Run = {
 *   type: 'run',
 *   formatting: { bold: true },
 *   content: [{ type: 'text', text: 'Hello' }],
 * };
 * ```
 */
export interface Run {
  type: 'run';
  /** Text formatting properties */
  formatting?: TextFormatting;
  /** Run-level tracked property changes (w:rPrChange) */
  propertyChanges?: RunPropertyChange[];
  /** Run content (text, tabs, breaks, etc.) */
  content: RunContent[];
}
