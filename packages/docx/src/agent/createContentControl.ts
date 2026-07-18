/**
 * Create a new inline content control (`w:sdt`) by wrapping a text span.
 *
 * Complements the discovery/edit functions in {@link ./contentControls}: where
 * those find and mutate existing controls, this wraps an exact run of text
 * inside a paragraph in a new control with a synthesized, Word-correct
 * `w:sdtPr`. Pure — the input {@link Document} is not mutated.
 */

import type {
  Document,
  BlockContent,
  InlineSdt,
  Paragraph,
  Run,
  RunContent,
  SdtType,
  SdtProperties,
} from '../types/document';
import { getParagraphText, getRunText, getHyperlinkText } from './text-utils';
import { synthesizeSdtPr } from './synthesizeSdtPr';
import {
  rebuild,
  findContentControl,
  findContentControls,
  isInlineContent,
  type ContentControlInfo,
} from './contentControls';

/** A create request failed: the target couldn't be resolved, or the wrap is invalid. */
export class ContentControlCreateError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ContentControlCreateError';
  }
}

/**
 * Where to create a control: an exact text span inside a paragraph. The
 * paragraph is located by Word `w14:paraId`, and the chosen `occurrence` of
 * `text` is wrapped in an inline control — including inside a table cell, where
 * block-level controls aren't allowed.
 */
export interface CreateContentControlTarget {
  /** Word `w14:paraId` of the paragraph containing the text. */
  paraId: string;
  /** Exact substring to wrap. */
  text: string;
  /** Which occurrence of `text` to wrap when it repeats (1-based; default 1). */
  occurrence?: number;
}

/** Modeled properties for a control created by {@link createContentControl}. */
export interface NewContentControlProps {
  /** Control type (default `richText`). */
  sdtType?: SdtType;
  /** Developer identifier (`w:tag`). */
  tag?: string;
  /** Friendly name (`w:alias`). */
  alias?: string;
  /** Numeric id (`w:id`). Default: auto-assigned, unique across the document. */
  id?: number;
  /** Lock setting (`w:lock`). */
  lock?: SdtProperties['lock'];
  /** Dropdown/combobox list items. */
  listItems?: { displayText: string; value: string }[];
  /** Date display format (`w:date/w:dateFormat`), for `date` controls. */
  dateFormat?: string;
  /** Initial checkbox state, for `checkbox` controls. */
  checked?: boolean;
  /** Whether the control starts in placeholder state (`w:showingPlcHdr`). */
  showingPlaceholder?: boolean;
}

/** Length an inline item contributes to {@link getParagraphText} (runs + hyperlinks only). */
function itemTextLength(item: Paragraph['content'][number]): number {
  if (item.type === 'run') return getRunText(item).length;
  if (item.type === 'hyperlink') return getHyperlinkText(item).length;
  return 0;
}

/** A `w:t` text node sliced to `[a, b)`, preserving boundary whitespace. */
function sliceText(node: Extract<RunContent, { type: 'text' }>, a: number, b?: number): RunContent {
  const text = node.text.slice(a, b);
  const next: Extract<RunContent, { type: 'text' }> = { type: 'text', text };
  if (node.preserveSpace || /^\s|\s$/.test(text)) next.preserveSpace = true;
  return next;
}

/**
 * Split a run's content at text-character offset `n` into `[left, right]` runs,
 * preserving the run's formatting and tracked property changes. Non-text run
 * content (tabs, breaks, drawings) lands on the side its position falls on.
 */
function splitRunAtChar(run: Run, n: number): [Run, Run] {
  const left: RunContent[] = [];
  const right: RunContent[] = [];
  let count = 0;
  for (const c of run.content) {
    if (c.type === 'text') {
      if (count >= n) {
        right.push(c);
      } else if (count + c.text.length <= n) {
        left.push(c);
        count += c.text.length;
      } else {
        const k = n - count;
        left.push(sliceText(c, 0, k));
        right.push(sliceText(c, k));
        count += c.text.length;
      }
    } else {
      (count >= n ? right : left).push(c);
    }
  }
  return [
    { ...run, content: left },
    { ...run, content: right },
  ];
}

/**
 * Wrap the character range `[start, end)` of a paragraph's inline `content` in
 * `sdt`, splitting runs at the boundaries. Items fully inside the range become
 * the control's content (preserving interior fields/tabs/breaks wholesale);
 * only an item straddling a boundary must be a run (to split).
 *
 * Throws {@link ContentControlCreateError} if a boundary falls inside a non-run
 * item (hyperlink/field), if the range overlaps an existing inline control, or
 * if the range contains content that cannot live inside an inline control.
 */
function wrapInlineSpan(
  content: Paragraph['content'],
  start: number,
  end: number,
  sdt: InlineSdt
): Paragraph['content'] {
  const before: Paragraph['content'] = [];
  const matched: InlineSdt['content'] = [];
  const after: Paragraph['content'] = [];
  let pos = 0;
  for (const item of content) {
    const len = itemTextLength(item);
    const iStart = pos;
    const iEnd = pos + len;
    pos = iEnd;
    if (iEnd <= start) {
      before.push(item);
      continue;
    }
    if (iStart >= end) {
      after.push(item);
      continue;
    }
    if (iStart >= start && iEnd <= end) {
      // Fully inside the span.
      if (item.type === 'inlineSdt') {
        throw new ContentControlCreateError(
          'The text span overlaps an existing inline content control; cannot wrap it.'
        );
      }
      if (!isInlineContent(item)) {
        throw new ContentControlCreateError(
          `The text span contains a '${item.type}' that cannot be placed inside an inline content control.`
        );
      }
      matched.push(item);
      continue;
    }
    // Straddles a boundary: must be a run to split cleanly.
    if (item.type !== 'run') {
      throw new ContentControlCreateError(
        item.type === 'inlineSdt'
          ? 'The text span overlaps an existing inline content control; cannot wrap it.'
          : `A content-control boundary falls inside a '${item.type}'; wrap on a run boundary instead.`
      );
    }
    const a = Math.max(start, iStart) - iStart;
    const b = Math.min(end, iEnd) - iStart;
    const [leftRun, rest] = splitRunAtChar(item, a);
    const [midRun, rightRun] = splitRunAtChar(rest, b - a);
    if (leftRun.content.length > 0) before.push(leftRun);
    matched.push(midRun);
    if (rightRun.content.length > 0) after.push(rightRun);
  }
  return [...before, { ...sdt, content: matched }, ...after];
}

/** Highest `w:id` across every control in the document (body + header/footer). */
function maxControlId(doc: Document): number {
  return findContentControls(doc, {}, { includeHeadersFooters: true }).reduce(
    (max, c) => (c.id != null && c.id > max ? c.id : max),
    0
  );
}

/**
 * SDT types {@link synthesizeSdtPr} emits a type marker for. The rest (`group`,
 * `buildingBlockGallery`, `equation`, `citation`, `bibliography`, `unknown`)
 * would serialize without a marker and silently reparse as `richText`, so a
 * created control is rejected rather than lose its type on round-trip.
 */
const SYNTHESIZABLE_SDT_TYPES = new Set<SdtType>([
  'richText',
  'plainText',
  'date',
  'dropDownList',
  'comboBox',
  'checkbox',
  'picture',
]);

/** Build the {@link SdtProperties} for a created control, with a synthesized raw `w:sdtPr`. */
function buildCreatedProps(props: NewContentControlProps, id: number): SdtProperties {
  const sdtType = props.sdtType ?? 'richText';
  if (!SYNTHESIZABLE_SDT_TYPES.has(sdtType)) {
    throw new ContentControlCreateError(
      `Cannot create a '${sdtType}' content control; supported types: ${[...SYNTHESIZABLE_SDT_TYPES].join(', ')}.`
    );
  }
  const out: SdtProperties = { sdtType, id };
  if (props.tag != null) out.tag = props.tag;
  if (props.alias != null) out.alias = props.alias;
  if (props.lock != null) out.lock = props.lock;
  if (props.listItems != null) out.listItems = props.listItems;
  if (props.dateFormat != null) out.dateFormat = props.dateFormat;
  if (props.checked != null) out.checked = props.checked;
  if (props.showingPlaceholder) out.showingPlaceholder = true;
  out.rawPropertiesXml = synthesizeSdtPr(out);
  return out;
}

/**
 * Rebuild body `blocks`, applying `transform` to the FIRST paragraph matching
 * `predicate` (descending block SDTs and table cells), immutably. `state.done`
 * stops the walk after the first match. Returns the same reference if unchanged.
 */
function mapFirstParagraph(
  blocks: BlockContent[],
  predicate: (p: Paragraph) => boolean,
  transform: (p: Paragraph) => Paragraph,
  state: { done: boolean }
): BlockContent[] {
  let changed = false;
  const out: BlockContent[] = [];
  for (const block of blocks) {
    if (state.done) {
      out.push(block);
      continue;
    }
    if (block.type === 'paragraph') {
      if (predicate(block)) {
        out.push(transform(block));
        state.done = true;
        changed = true;
      } else {
        out.push(block);
      }
    } else if (block.type === 'blockSdt') {
      const next = mapFirstParagraph(block.content, predicate, transform, state);
      if (next !== block.content) {
        out.push({ ...block, content: next });
        changed = true;
      } else {
        out.push(block);
      }
    } else if (block.type === 'table') {
      let tableChanged = false;
      const rows = block.rows.map((row) => {
        if (state.done) return row;
        let rowChanged = false;
        const cells = row.cells.map((cell) => {
          if (state.done) return cell;
          const cellContent = cell.content as BlockContent[];
          const next = mapFirstParagraph(cellContent, predicate, transform, state);
          if (next !== cellContent) {
            rowChanged = true;
            return { ...cell, content: next as typeof cell.content };
          }
          return cell;
        });
        if (rowChanged) {
          tableChanged = true;
          return { ...row, cells };
        }
        return row;
      });
      if (tableChanged) {
        out.push({ ...block, rows });
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

/**
 * Wrap an exact text span inside a paragraph in a new inline content control
 * (`w:sdt`), returning a new {@link Document} and the created control's
 * {@link ContentControlInfo}. Pure — the input is not mutated. This is the form
 * needed inside table cells and mid-sentence, where block controls aren't
 * allowed: runs are split at the span boundaries (formatting preserved) and
 * interior fields/tabs/breaks are kept wholesale.
 *
 * The control's `w:sdtPr` is synthesized from `props`, and its `w:id` is
 * auto-assigned (unique across the document) when `props.id` is omitted, so the
 * control round-trips and `findContentControl(doc, { tag })` resolves it after a
 * save/reload.
 *
 * **Body only:** the search covers body paragraphs and block/table content —
 * paragraphs inside headers or footers are not reachable. Passing a `paraId`
 * from a header/footer part throws {@link ContentControlCreateError} with a
 * "No paragraph found" message.
 *
 * @throws {@link ContentControlCreateError} when the paragraph or text isn't
 * found, the span overlaps an existing control or crosses a non-run boundary,
 * the `sdtType` can't be synthesized, or a supplied `id` already exists.
 */
export function createContentControl(
  doc: Document,
  target: CreateContentControlTarget,
  props: NewContentControlProps = {}
): { doc: Document; control: ContentControlInfo } {
  if (!target.text) {
    throw new ContentControlCreateError('`text` must be a non-empty string.');
  }
  const occurrence = target.occurrence ?? 1;
  if (occurrence < 1) {
    throw new ContentControlCreateError('`occurrence` must be >= 1.');
  }
  if (
    props.id != null &&
    findContentControl(doc, { id: props.id }, { includeHeadersFooters: true })
  ) {
    throw new ContentControlCreateError(
      `A content control with id ${props.id} already exists; omit \`id\` to auto-assign a unique one.`
    );
  }
  const id = props.id ?? maxControlId(doc) + 1;
  const properties = buildCreatedProps(props, id);

  const transform = (para: Paragraph): Paragraph => {
    const full = getParagraphText(para);
    let from = -1;
    let count = 0;
    for (
      let i = full.indexOf(target.text);
      i !== -1;
      i = full.indexOf(target.text, i + target.text.length)
    ) {
      if (++count === occurrence) {
        from = i;
        break;
      }
    }
    if (from === -1) {
      throw new ContentControlCreateError(
        count === 0
          ? `Text ${JSON.stringify(target.text)} was not found in the paragraph.`
          : `Occurrence ${occurrence} of ${JSON.stringify(target.text)} not found (only ${count} present).`
      );
    }
    const sdt: InlineSdt = { type: 'inlineSdt', properties, content: [] };
    return {
      ...para,
      content: wrapInlineSpan(para.content, from, from + target.text.length, sdt),
    };
  };

  const state = { done: false };
  const content = mapFirstParagraph(
    doc.package.document.content,
    (p) => p.paraId === target.paraId,
    transform,
    state
  );
  if (!state.done) {
    throw new ContentControlCreateError(
      `No paragraph found with paraId ${JSON.stringify(target.paraId)}.`
    );
  }
  const nextDoc = rebuild(doc, content);

  const control = findContentControl(nextDoc, { id });
  if (!control) {
    // Unreachable: the control was just inserted with this unique id.
    throw new ContentControlCreateError('Created control could not be located after insertion.');
  }
  return { doc: nextDoc, control };
}
