/**
 * Content-control (SDT) addressing for the document model.
 *
 * Content controls (`w:sdt`) are the natural anchor for template logic and
 * agent edits: they survive the round trip (see the parser + serializer) and
 * carry a stable `tag`/`alias`/`id`. This module discovers and edits them
 * without a DOM or an editor instance, so server-side pipelines and AI agents
 * can find an anchor by tag and act on it.
 *
 * Both **block-level** (`w:sdt` wrapping paragraphs/tables) and **inline**
 * (`w:sdt` inside a paragraph) controls are addressed, including inline
 * controls inside table cells (and nested tables). With
 * `{ includeHeadersFooters: true }` the walk also covers header/footer parts.
 *
 * Not surfaced (model limitations): a block SDT placed directly inside a table
 * cell (`TableCell.content` cannot hold one), an inline SDT inside a hyperlink
 * (`Hyperlink.children` excludes it), and controls inside tracked-change
 * wrappers.
 */

import type {
  Document,
  DocumentBody,
  BlockContent,
  BlockSdt,
  InlineSdt,
  Paragraph,
  Table,
  HeaderFooter,
  Run,
  TextFormatting,
  SdtType,
  SdtProperties,
  SdtDataBinding,
} from '../types/document';
import { getParagraphText, getTableText, getRunText, getHyperlinkText } from './text-utils';

/** Filter for {@link findContentControls}. All provided fields must match (AND). */
export interface ContentControlFilter {
  /** Developer identifier (`w:tag`), exact match. */
  tag?: string;
  /** Friendly name (`w:alias`), exact match. */
  alias?: string;
  /** Numeric id (`w:id`), exact match. */
  id?: number;
  /** Control type projection (`richText`, `dropDownList`, …). */
  type?: SdtType;
}

/**
 * Where a control lives. `body` = the main document story; `header`/`footer`
 * = a page-furniture part addressed by its relationship id (the key into
 * `package.headers`/`package.footers`).
 */
export type ContentControlLocation = { part: 'body' } | { part: 'header' | 'footer'; rId: string };

/** A discovered content control plus enough context to address and edit it. */
export interface ContentControlInfo {
  /** Developer identifier (`w:tag`). */
  tag?: string;
  /** Friendly name (`w:alias`). */
  alias?: string;
  /** Numeric id (`w:id`). */
  id?: number;
  /** Control type projection. */
  sdtType: SdtType;
  /** Lock setting, if any. A locked control should refuse content edits. */
  lock?: SdtProperties['lock'];
  /** Dropdown/combobox list items, if modeled. */
  listItems?: { displayText: string; value: string }[];
  /** Placeholder docPart reference, if any. */
  placeholder?: string;
  /** Whether the control is currently showing placeholder text (`w:showingPlcHdr`). */
  showingPlaceholder?: boolean;
  /** Checkbox state, for checkbox controls. */
  checked?: boolean;
  /** Date format string, for date controls. */
  dateFormat?: string;
  /** XML data binding (`w:dataBinding`), if the control is bound. */
  dataBinding?: SdtDataBinding;
  /** Plain text of the control's content (paragraphs/tables/nested controls flattened). */
  text: string;
  /**
   * Block-index path to this control: top-level `[i]`, a control nested in the
   * i-th block's content `[i, j]`, and so on. For inline / cell controls it is
   * the block indices of the nearest enclosing blocks.
   */
  path: number[];
  /** Nesting depth = number of enclosing content controls (0 = not inside another control). */
  depth: number;
  /** Block-level (`w:sdt` at block level) or inline (`w:sdt` inside a paragraph). */
  kind: 'block' | 'inline';
  /** Where the control lives (body vs a header/footer part). */
  location: ContentControlLocation;
}

/** Narrow a {@link Document} or {@link DocumentBody} to its block list. */
function bodyOf(input: Document | DocumentBody): DocumentBody {
  return 'package' in input ? input.package.document : input;
}

/** Plain text of a control's content, descending into tables and nested SDTs. */
export function getContentControlText(control: BlockSdt | InlineSdt): string {
  return control.type === 'inlineSdt'
    ? inlineContentText(control.content)
    : blocksText(control.content);
}

function blocksText(blocks: BlockContent[]): string {
  const parts: string[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph') parts.push(getParagraphText(block));
    else if (block.type === 'table') parts.push(getTableText(block));
    else if (block.type === 'blockSdt') parts.push(blocksText(block.content));
  }
  return parts.join('\n');
}

/**
 * Plain text of an inline SDT's content: runs, hyperlink runs, field result
 * text, nested inline SDTs (recursively), and math. The block flattener
 * (`blocksText` → `getParagraphText`) never reaches this content, so without
 * this an inline control's `text` projection would be empty.
 */
function inlineContentText(content: InlineSdt['content']): string {
  const parts: string[] = [];
  for (const node of content) {
    switch (node.type) {
      case 'run':
        parts.push(getRunText(node));
        break;
      case 'hyperlink':
        parts.push(getHyperlinkText(node));
        break;
      case 'simpleField':
        for (const c of node.content) {
          parts.push(c.type === 'run' ? getRunText(c) : getHyperlinkText(c));
        }
        break;
      case 'complexField':
        for (const r of node.fieldResult) parts.push(getRunText(r));
        break;
      case 'inlineSdt':
        parts.push(inlineContentText(node.content));
        break;
      case 'mathEquation':
        if (node.plainText) parts.push(node.plainText);
        break;
    }
  }
  return parts.join('');
}

function matches(props: SdtProperties, filter: ContentControlFilter): boolean {
  if (filter.tag !== undefined && props.tag !== filter.tag) return false;
  if (filter.alias !== undefined && props.alias !== filter.alias) return false;
  if (filter.id !== undefined && props.id !== filter.id) return false;
  if (filter.type !== undefined && props.sdtType !== filter.type) return false;
  return true;
}

function infoOf(
  control: BlockSdt | InlineSdt,
  path: number[],
  depth: number,
  location: ContentControlLocation
): ContentControlInfo {
  const p = control.properties;
  return {
    tag: p.tag,
    alias: p.alias,
    id: p.id,
    sdtType: p.sdtType,
    lock: p.lock,
    listItems: p.listItems,
    placeholder: p.placeholder,
    showingPlaceholder: p.showingPlaceholder,
    checked: p.checked,
    dateFormat: p.dateFormat,
    dataBinding: p.dataBinding,
    text: getContentControlText(control),
    path,
    depth,
    kind: control.type === 'inlineSdt' ? 'inline' : 'block',
    location,
  };
}

/** Options for {@link findContentControls}. */
export interface FindContentControlsOptions {
  /**
   * When `true`, also search header/footer parts — but only when a full
   * {@link Document} is passed (a bare `DocumentBody` carries no parts, so this
   * then searches the body only and never throws). Defaults to `false` (the main
   * document story only).
   */
  includeHeadersFooters?: boolean;
}

/**
 * Find every content control in the document — block-level AND inline —
 * optionally filtered by tag/alias/id/type. Results are in strict document
 * order; nested controls follow their parent.
 *
 * The walk descends body blocks, block SDTs, tables (row-major, including
 * nested tables) into cell content, and paragraph inline content (inline
 * SDTs, recursing into nested inline SDTs). With `{ includeHeadersFooters: true }`
 * and a full {@link Document}, header then footer parts are searched after the
 * body, each sorted by relationship id for deterministic order.
 *
 * Not surfaced (model limitations, documented): a block SDT placed directly
 * inside a table cell (`TableCell.content` is `(Paragraph | Table)[]`), an
 * inline SDT inside a hyperlink (`Hyperlink.children` excludes it), and
 * controls buried inside tracked-change wrappers.
 */
export function findContentControls(
  input: Document | DocumentBody,
  filter: ContentControlFilter = {},
  options: FindContentControlsOptions = {}
): ContentControlInfo[] {
  const out: ContentControlInfo[] = [];

  // Descend a paragraph's (or inline SDT's) inline content for inline SDTs.
  // `path` is the block-index path of the enclosing paragraph; `depth` counts
  // enclosing content controls.
  const walkInline = (
    content: Paragraph['content'],
    location: ContentControlLocation,
    path: number[],
    depth: number
  ): void => {
    for (const node of content) {
      if (node.type === 'inlineSdt') {
        if (matches(node.properties, filter)) out.push(infoOf(node, path, depth, location));
        walkInline(node.content, location, path, depth + 1); // nested inline controls
      }
      // Hyperlink children are runs only (no SDT possible per the model);
      // tracked-change wrappers are not descended in v1.
    }
  };

  // Descend a BlockContent[] (body root, blockSdt content, or a table cell).
  // Tables and cells are not controls, so `depth` is unchanged through them.
  const walkBlocks = (
    blocks: BlockContent[],
    location: ContentControlLocation,
    path: number[],
    depth: number
  ): void => {
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      const blockPath = [...path, i];
      if (block.type === 'blockSdt') {
        if (matches(block.properties, filter)) out.push(infoOf(block, blockPath, depth, location));
        walkBlocks(block.content, location, blockPath, depth + 1); // nested controls
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walkBlocks(cell.content, location, blockPath, depth);
          }
        }
      } else if (block.type === 'paragraph') {
        walkInline(block.content, location, blockPath, depth);
      }
    }
  };

  walkBlocks(bodyOf(input).content, { part: 'body' }, [], 0);

  if (options.includeHeadersFooters && 'package' in input) {
    const walkParts = (
      parts: Map<string, HeaderFooter> | undefined,
      part: 'header' | 'footer'
    ): void => {
      if (!parts) return;
      for (const rId of [...parts.keys()].sort()) {
        walkBlocks(parts.get(rId)!.content, { part, rId }, [], 0);
      }
    };
    walkParts(input.package.headers, 'header');
    walkParts(input.package.footers, 'footer');
  }

  return out;
}

/** Convenience: the first control matching `filter`, or `undefined`. */
export function findContentControl(
  input: Document | DocumentBody,
  filter: ContentControlFilter,
  options: FindContentControlsOptions = {}
): ContentControlInfo | undefined {
  return findContentControls(input, filter, options)[0];
}

// ============================================================================
// MUTATION (edit a control by tag)
// ============================================================================

/** No control matched the filter. */
export class ContentControlNotFoundError extends Error {
  constructor(filter: ContentControlFilter) {
    super(`No content control matched ${JSON.stringify(filter)}`);
    this.name = 'ContentControlNotFoundError';
  }
}

/** The matched control's lock forbids the attempted edit (pass `force` to override). */
export class ContentControlLockedError extends Error {
  constructor(lock: SdtProperties['lock'], op: 'edit' | 'remove') {
    super(`Content control is ${lock}; cannot ${op} it without { force: true }`);
    this.name = 'ContentControlLockedError';
  }
}

/**
 * The control's type doesn't support free text/block replacement (e.g. a
 * dropdown, date, checkbox, or picture control), so writing arbitrary content
 * would desync the type marker from its value. Use a type-specific setter, or
 * pass `{ force: true }` to override.
 */
export class ContentControlTypeError extends Error {
  constructor(sdtType: SdtType) {
    super(
      `Content control is a '${sdtType}' control; replacing its content with free text ` +
        `would desync it. Use a type-specific value setter or pass { force: true }.`
    );
    this.name = 'ContentControlTypeError';
  }
}

/**
 * The control is bound to a Custom XML data store (`w:dataBinding`). Writing its
 * content won't stick — Word re-renders the control from the bound XML node — so
 * the write is refused. Update the data store instead, or pass `{ force: true }`.
 */
export class ContentControlBoundError extends Error {
  constructor() {
    super(
      'Content control is data-bound (w:dataBinding); its content is driven by the ' +
        'Custom XML store and a direct write will not persist. Update the store, or pass { force: true }.'
    );
    this.name = 'ContentControlBoundError';
  }
}

/**
 * Control types whose content is free-form and safe to replace with text/blocks.
 * Typed controls (dropdown, date, checkbox, picture) carry structured state that
 * arbitrary content would contradict, and `group` exists to lock/contain nested
 * structure — all gated unless forced.
 */
const TEXT_REPLACEABLE_TYPES = new Set<SdtType>(['richText', 'plainText', 'unknown']);

/** True if free text/block content can safely replace this control type's content. */
export function isTextReplaceable(sdtType: SdtType): boolean {
  return TEXT_REPLACEABLE_TYPES.has(sdtType);
}

/** `w:lock` values that forbid editing the control's content. */
export function isContentLocked(lock: SdtProperties['lock']): boolean {
  return lock === 'contentLocked' || lock === 'sdtContentLocked';
}

/** `w:lock` values that forbid deleting the control. */
export function isDeletionLocked(lock: SdtProperties['lock']): boolean {
  return lock === 'sdtLocked' || lock === 'sdtContentLocked';
}

/**
 * True if the raw `w:sdtPr` carries a (w15) repeating-section structure. Matches
 * the element name (`<w15:repeatingSection>` / `<w15:repeatingSectionItem>`) so
 * a tag/alias value that merely contains the word doesn't false-match.
 */
export function hasRepeatingSection(props: SdtProperties): boolean {
  return /<w15:repeatingSection(Item)?[\s/>]/.test(props.rawPropertiesXml ?? '');
}

/** True if the control is bound to a Custom XML data store (`w:dataBinding`). */
export function isDataBound(props: SdtProperties): boolean {
  return props.dataBinding != null;
}

/**
 * Strip `<w:showingPlcHdr/>` from a raw `w:sdtPr` string. When real content is
 * written into a control that was showing its placeholder, the flag must go or
 * Word keeps rendering the (now-stale) placeholder styling over real content.
 */
export function clearShowingPlaceholderXml(raw: string | undefined): string | undefined {
  if (!raw) return raw;
  return raw
    .replace(/<w:showingPlcHdr\b[^>]*\/>/g, '')
    .replace(/<w:showingPlcHdr\b[^>]*>[\s\S]*?<\/w:showingPlcHdr>/g, '');
}

/** Properties for a control after real content is written: placeholder flag cleared. */
function propsAfterContentWrite(props: SdtProperties): SdtProperties {
  if (!props.showingPlaceholder && !/showingPlcHdr/.test(props.rawPropertiesXml ?? '')) {
    return props;
  }
  const next: SdtProperties = { ...props, showingPlaceholder: false };
  const cleaned = clearShowingPlaceholderXml(props.rawPropertiesXml);
  if (cleaned !== undefined) next.rawPropertiesXml = cleaned;
  return next;
}

function paragraph(text: string): BlockContent {
  return {
    type: 'paragraph',
    content: text ? [{ type: 'run', content: [{ type: 'text', text }] }] : [],
  };
}

/**
 * Turn a string into paragraphs (one per newline), or deep-clone block input.
 * A `plainText` control is single-paragraph in OOXML, so its string content is
 * collapsed to one paragraph rather than split — multiple paragraphs would make
 * Word repair the control on open.
 */
function toBlocks(
  replacement: string | BlockContent[],
  opts: { singleParagraph?: boolean } = {}
): BlockContent[] {
  if (typeof replacement !== 'string') {
    // Clone so the caller can't later mutate content shared with the result.
    return structuredClone(replacement);
  }
  if (opts.singleParagraph) return [paragraph(replacement)];
  return replacement.split('\n').map(paragraph);
}

export type ControlOp = (control: BlockSdt) => BlockContent[];

/**
 * Rebuild `blocks`, applying `op` to the first control matching `filter`. The
 * op's result (0, 1, or many blocks) is spliced in at the control's own level
 * — including when the control is nested inside another control — so a
 * remove/unwrap never leaves a placeholder behind. `state.done` stops the
 * walk after the first match.
 */
export function applyToFirst(
  blocks: BlockContent[],
  filter: ContentControlFilter,
  op: ControlOp,
  state: { done: boolean }
): BlockContent[] {
  const out: BlockContent[] = [];
  for (const block of blocks) {
    if (state.done) {
      out.push(block);
      continue;
    }
    // Controls are searched at body level and inside other controls. Table
    // cells are not searched: the current model types a cell as
    // (Paragraph | Table)[], and the table parser does not yet surface a
    // cell-level w:sdt (which OOXML's CT_Tc does permit) — see CONTENT-CONTROLS.md.
    if (block.type === 'blockSdt') {
      if (matches(block.properties, filter)) {
        out.push(...op(block));
        state.done = true;
        continue;
      }
      out.push({ ...block, content: applyToFirst(block.content, filter, op, state) });
    } else {
      out.push(block);
    }
  }
  return out;
}

export function rebuild(doc: Document, content: BlockContent[]): Document {
  return {
    ...doc,
    package: {
      ...doc.package,
      document: { ...doc.package.document, content },
    },
  };
}

// ============================================================================
// INLINE MUTATION (write/remove an inline control by tag)
// ============================================================================

/** One node of an inline SDT's content. */
export type InlineContent = InlineSdt['content'][number];

/** Op applied to a matched block control (alias of {@link ControlOp}). */
export type BlockControlOp = ControlOp;

/** Op applied to a matched inline control; its result is spliced among the paragraph's inline siblings. */
export type InlineControlOp = (control: InlineSdt) => InlineContent[];

/**
 * A `BlockContent[]` replacement was targeted at an inline control, which can
 * only hold inline content (runs, hyperlinks, fields, nested inline SDTs, math).
 * Splicing a paragraph into a paragraph would be invalid OOXML — pass a string,
 * or a single paragraph whose content is entirely inline.
 */
export class ContentControlKindError extends Error {
  constructor(detail: string) {
    super(detail);
    this.name = 'ContentControlKindError';
  }
}

/** Type guard: a paragraph-content node that is also valid inside an inline SDT. */
export function isInlineContent(node: Paragraph['content'][number]): node is InlineContent {
  switch (node.type) {
    case 'run':
    case 'hyperlink':
    case 'simpleField':
    case 'complexField':
    case 'inlineSdt':
    case 'mathEquation':
      return true;
    default:
      return false;
  }
}

/** Formatting for a filled inline run: the placeholder's first run formatting, so the value matches the field it replaces. */
function resolveFillFormatting(control: InlineSdt): TextFormatting | undefined {
  const firstRun = control.content.find((c): c is Run => c.type === 'run');
  return firstRun?.formatting ? structuredClone(firstRun.formatting) : undefined;
}

/** A `w:t` content node, preserving boundary whitespace (`xml:space="preserve"`). */
function makeText(text: string): Run['content'][number] {
  const node: { type: 'text'; text: string; preserveSpace?: boolean } = { type: 'text', text };
  if (/^\s|\s$/.test(text)) node.preserveSpace = true;
  return node;
}

/**
 * Build inline content for filling an inline control. A string becomes a single
 * run inheriting the placeholder's formatting; for a richText control `\n`
 * becomes a `w:br` within that run, while a plainText control never gets a break
 * (it serializes as `<w:text/>`, which Word repairs otherwise). An empty string
 * yields no content. A `BlockContent[]` is rejected unless it is a single
 * all-inline paragraph (whose inline content is lifted, dropping paragraph-level
 * metadata).
 */
function toInline(replacement: string | BlockContent[], control: InlineSdt): InlineContent[] {
  if (typeof replacement !== 'string') {
    if (
      replacement.length === 1 &&
      replacement[0].type === 'paragraph' &&
      replacement[0].content.every(isInlineContent)
    ) {
      return structuredClone(replacement[0].content) as InlineContent[];
    }
    throw new ContentControlKindError(
      'Cannot place block content (paragraphs/tables) inside an inline content control. ' +
        'Pass a string, or a single paragraph whose content is entirely inline.'
    );
  }
  if (replacement === '') return [];
  const fmt = resolveFillFormatting(control);
  const isPlain = control.properties.sdtType === 'plainText';
  const runContent: Run['content'] = [];
  if (isPlain || !replacement.includes('\n')) {
    runContent.push(makeText(replacement));
  } else {
    replacement.split('\n').forEach((line, i) => {
      if (i > 0) runContent.push({ type: 'break', breakType: 'textWrapping' });
      if (line) runContent.push(makeText(line));
    });
  }
  const run: Run = { type: 'run', content: runContent };
  if (fmt) run.formatting = fmt;
  return [run];
}

/**
 * Walk state shared by the control mutators. `matched` counts the controls the
 * op has been applied to so far; `stopAtFirst` halts the walk after the first
 * match (the default). The `{ all: true }` mutator option clears `stopAtFirst`
 * so every matching control is mutated in one pass.
 */
type ControlWalkState = { matched: number; stopAtFirst: boolean };

/** A first-match walk is finished once it has mutated one control. */
const walkDone = (state: ControlWalkState): boolean => state.stopAtFirst && state.matched > 0;

/**
 * Rebuild a paragraph's (or inline SDT's) inline content, applying `inlineOp` to
 * the first matching inline control (or every match, when the walk isn't
 * stop-at-first) and recursing into nested inline SDTs. Returns the same array
 * reference when nothing changed (purity / cheap diff).
 */
function applyToFirstInlineContent(
  content: Paragraph['content'],
  filter: ContentControlFilter,
  inlineOp: InlineControlOp,
  state: ControlWalkState
): Paragraph['content'] {
  let changed = false;
  const out: Paragraph['content'] = [];
  for (const node of content) {
    if (walkDone(state) || node.type !== 'inlineSdt') {
      out.push(node);
      continue;
    }
    if (matches(node.properties, filter)) {
      out.push(...inlineOp(node));
      state.matched += 1;
      changed = true;
      continue;
    }
    const inner = node.content as Paragraph['content'];
    const nested = applyToFirstInlineContent(inner, filter, inlineOp, state);
    if (nested !== inner) {
      // Inside an inline SDT every node is inline-only, so the narrow cast is sound.
      out.push({ ...node, content: nested as InlineSdt['content'] });
      changed = true;
    } else {
      out.push(node);
    }
  }
  return changed ? out : content;
}

/** Rebuild a table, applying the ops to the first matching control in any cell (recursing nested tables). */
function applyToFirstInTable(
  table: Table,
  filter: ContentControlFilter,
  blockOp: BlockControlOp,
  inlineOp: InlineControlOp,
  state: ControlWalkState
): Table {
  if (walkDone(state)) return table;
  let changed = false;
  const rows = table.rows.map((row) => {
    if (walkDone(state)) return row;
    let rowChanged = false;
    const cells = row.cells.map((cell) => {
      if (walkDone(state)) return cell;
      const cellContent = cell.content as BlockContent[];
      const next = applyToFirstControl(cellContent, filter, blockOp, inlineOp, state);
      if (next !== cellContent) {
        rowChanged = true;
        // A cell cannot hold a block SDT in the model, so this stays (Paragraph | Table)[].
        return { ...cell, content: next as typeof cell.content };
      }
      return cell;
    });
    if (rowChanged) {
      changed = true;
      return { ...row, cells };
    }
    return row;
  });
  return changed ? { ...table, rows } : table;
}

/**
 * Rebuild `blocks`, applying the kind-appropriate op to the first matching
 * control — block controls via `blockOp`, inline controls (including inside
 * table cells and nested tables) via `inlineOp`. The result is spliced at the
 * control's own level. The walk stops after the first match unless `state` is in
 * apply-to-all mode (`stopAtFirst === false`), in which case every match is hit.
 */
function applyToFirstControl(
  blocks: BlockContent[],
  filter: ContentControlFilter,
  blockOp: BlockControlOp,
  inlineOp: InlineControlOp,
  state: ControlWalkState
): BlockContent[] {
  let changed = false;
  const out: BlockContent[] = [];
  for (const block of blocks) {
    if (walkDone(state)) {
      out.push(block);
      continue;
    }
    if (block.type === 'blockSdt') {
      if (matches(block.properties, filter)) {
        out.push(...blockOp(block));
        state.matched += 1;
        changed = true;
        continue;
      }
      const nested = applyToFirstControl(block.content, filter, blockOp, inlineOp, state);
      if (nested !== block.content) {
        out.push({ ...block, content: nested });
        changed = true;
      } else {
        out.push(block);
      }
    } else if (block.type === 'table') {
      const next = applyToFirstInTable(block, filter, blockOp, inlineOp, state);
      if (next !== block) {
        out.push(next);
        changed = true;
      } else {
        out.push(block);
      }
    } else if (block.type === 'paragraph') {
      const next = applyToFirstInlineContent(block.content, filter, inlineOp, state);
      if (next !== block.content) {
        out.push({ ...block, content: next });
        changed = true;
      } else {
        out.push(block);
      }
    } else {
      out.push(block);
    }
  }
  return changed ? out : blocks;
}

/** Replace one header/footer part immutably (mirrors {@link rebuild} for the HF maps). */
function rebuildPart(
  doc: Document,
  kind: 'header' | 'footer',
  rId: string,
  part: HeaderFooter
): Document {
  const key = kind === 'header' ? 'headers' : 'footers';
  const nextMap = new Map(doc.package[key]);
  nextMap.set(rId, part);
  return { ...doc, package: { ...doc.package, [key]: nextMap } };
}

/**
 * Apply a content-control mutation across the body and — when
 * `includeHeadersFooters` — header/footer parts (headers then footers, by rId).
 * By default it stops at the first match; with `all: true` it applies the op to
 * every matching control in document order. `finalizeBody` post-processes the
 * rebuilt body content (e.g. the empty-body backstop for removals).
 *
 * The op's own guards (lock/type/data-binding) throw on the first offending
 * match, so an `all` run is atomic: it either mutates every match or — if one is
 * refused — throws and leaves the document untouched (the caller's input is never
 * mutated; a partially-rebuilt tree is discarded with the exception). Throws
 * {@link ContentControlNotFoundError} if nothing matched.
 */
export function applyControlMutation(
  doc: Document,
  filter: ContentControlFilter,
  blockOp: BlockControlOp,
  inlineOp: InlineControlOp,
  includeHeadersFooters: boolean,
  finalizeBody: (content: BlockContent[]) => BlockContent[] = (c) => c,
  all = false
): Document {
  const state: ControlWalkState = { matched: 0, stopAtFirst: !all };
  let next = doc;

  const bodyContent = applyToFirstControl(
    doc.package.document.content,
    filter,
    blockOp,
    inlineOp,
    state
  );
  if (bodyContent !== doc.package.document.content) {
    next = rebuild(next, finalizeBody(bodyContent));
  }
  if (walkDone(state)) return next;

  if (includeHeadersFooters) {
    for (const kind of ['header', 'footer'] as const) {
      const parts = kind === 'header' ? next.package.headers : next.package.footers;
      if (!parts) continue;
      for (const rId of [...parts.keys()].sort()) {
        const part = parts.get(rId)!;
        const nextContent = applyToFirstControl(part.content, filter, blockOp, inlineOp, state);
        if (nextContent !== part.content) {
          next = rebuildPart(next, kind, rId, { ...part, content: nextContent });
        }
        if (walkDone(state)) return next;
      }
    }
  }
  if (state.matched === 0) throw new ContentControlNotFoundError(filter);
  return next;
}

/** Shared write guards for block + inline content writes (all property-only, so kind-agnostic). */
function assertContentWritable(props: SdtProperties, force: boolean | undefined): void {
  if (!force && isContentLocked(props.lock))
    throw new ContentControlLockedError(props.lock, 'edit');
  if (!force && !isTextReplaceable(props.sdtType)) throw new ContentControlTypeError(props.sdtType);
  if (!force && isDataBound(props)) throw new ContentControlBoundError();
}

/** Shared deletion guards for block + inline removal. */
function assertRemovable(
  props: SdtProperties,
  options: { force?: boolean; keepContent?: boolean }
): void {
  if (!options.force && isDeletionLocked(props.lock)) {
    throw new ContentControlLockedError(props.lock, 'remove');
  }
  if (options.keepContent && !options.force && hasRepeatingSection(props)) {
    throw new ContentControlLockedError(props.lock, 'remove');
  }
}

/**
 * Replace the content of the first control matching `filter` — block-level OR
 * inline (including inside table cells and, with `includeHeadersFooters: true`,
 * headers and footers). `replacement` may be a string or block content.
 *
 * - For a **block** control the string is split into paragraphs on newlines
 *   (a `plainText` control collapses to one paragraph); block content is used
 *   as-is (cloned).
 * - For an **inline** control the string becomes a single run that inherits the
 *   placeholder's formatting (so the value matches the field it replaces). A
 *   richText control turns `\n` into a line break; a plainText control never
 *   gets one. Passing `BlockContent[]` to an inline control throws
 *   {@link ContentControlKindError} unless it is a single all-inline paragraph.
 *
 * Pass `{ all: true }` to fill **every** control matching `filter` (one logical
 * value that recurs under a shared tag — e.g. a name in the body, a running
 * header, and several table cells) instead of just the first.
 *
 * The control's properties, tag/alias, and lossless raw `w:sdtPr` are preserved.
 * When the control was showing its placeholder (`w:showingPlcHdr`), that flag is
 * cleared so Word doesn't render the new content as placeholder text.
 *
 * Throws {@link ContentControlNotFoundError} if nothing matches,
 * {@link ContentControlLockedError} if the lock forbids editing,
 * {@link ContentControlTypeError} for a typed (dropdown/date/…) control, and
 * {@link ContentControlBoundError} for a data-bound control. Pass
 * `{ force: true }` to override the guards.
 */
export function setContentControlContent(
  doc: Document,
  filter: ContentControlFilter,
  replacement: string | BlockContent[],
  options: { force?: boolean; includeHeadersFooters?: boolean; all?: boolean } = {}
): Document {
  const blockOp: BlockControlOp = (control) => {
    assertContentWritable(control.properties, options.force);
    return [
      {
        ...control,
        properties: propsAfterContentWrite(control.properties),
        content: toBlocks(replacement, {
          singleParagraph: control.properties.sdtType === 'plainText',
        }),
      },
    ];
  };
  const inlineOp: InlineControlOp = (control) => {
    assertContentWritable(control.properties, options.force);
    return [
      {
        ...control,
        properties: propsAfterContentWrite(control.properties),
        content: toInline(replacement, control),
      },
    ];
  };
  return applyControlMutation(
    doc,
    filter,
    blockOp,
    inlineOp,
    options.includeHeadersFooters ?? false,
    undefined,
    options.all ?? false
  );
}

/**
 * Remove the first control matching `filter` — block-level OR inline (incl.
 * inside table cells and, with `includeHeadersFooters: true`, headers/footers).
 * With `keepContent: true` the control's content is unwrapped in place (the box goes
 * away, the content stays) — block content lifts to its block siblings, inline
 * content stays inline in the enclosing paragraph. Otherwise the control and
 * its content are deleted.
 *
 * Pass `{ all: true }` to remove **every** control matching `filter` (e.g. to
 * flatten a finished template by unwrapping all controls) instead of the first.
 *
 * Unwrapping a repeating-section (item) is refused unless `force`, since lifting
 * its blocks out would orphan the (w15) repeating structure.
 *
 * Throws {@link ContentControlNotFoundError} / {@link ContentControlLockedError}
 * as {@link setContentControlContent} does.
 */
export function removeContentControl(
  doc: Document,
  filter: ContentControlFilter,
  options: {
    force?: boolean;
    keepContent?: boolean;
    includeHeadersFooters?: boolean;
    all?: boolean;
  } = {}
): Document {
  const blockOp: BlockControlOp = (control) => {
    assertRemovable(control.properties, options);
    return options.keepContent ? control.content : [];
  };
  const inlineOp: InlineControlOp = (control) => {
    assertRemovable(control.properties, options);
    return options.keepContent ? control.content : [];
  };
  // Never leave a structurally empty body (an empty <w:body> is invalid for Word
  // consumers). Body-only: an inline removal keeps the enclosing paragraph, and
  // HF parts are rebuilt separately.
  const finalizeBody = (content: BlockContent[]): BlockContent[] =>
    content.length > 0 ? content : [{ type: 'paragraph' as const, content: [] }];
  return applyControlMutation(
    doc,
    filter,
    blockOp,
    inlineOp,
    options.includeHeadersFooters ?? false,
    finalizeBody,
    options.all ?? false
  );
}
