import type { DocxTextRange } from './edits';
import type { YrsCellLoc, YrsLoc } from './index';
import { displayPositionToYrsLoc } from './inputPositionMap';
import type { YrsPointerProjectionTarget } from './yrsPositionProjection';
import { findWordBoundaries } from '../utils/textSelection';
import { DisplayPositionIndex } from './displayPositionIndex';
import { acceptedOffsetOf } from './pointPosition';
import {
  storyOffsetForLoc,
  storyPlainText,
  tablePlainText,
  type TablePayloadRow,
} from './storyPlainText';

/** @internal How far a selection grows around a display position. */
export type DocxSelectionUnit = 'word' | 'paragraph' | 'story';

/** @internal A selection in the display positions of one root story's layout. */
export interface DocxDisplayRange {
  anchor: number;
  head: number;
}

/** @internal The text a display selection covers. */
export interface DocxDisplaySelectionText {
  /** Plain text as the clipboard takes it: tabs and breaks as characters, tables as tab-separated rows. */
  text: string;
  /** The accepted-view range, or null when the selection spans stories. */
  range: DocxTextRange | null;
}

export interface DocxDisplaySelectionInfo {
  paraId: string | null;
  selectedText: string;
  paragraphText: string;
  before: string;
  after: string;
}

interface ResolvedPosition {
  target: YrsPointerProjectionTarget;
  loc: YrsLoc;
}

function resolvePosition(
  index: DisplayPositionIndex,
  rootStory: string,
  position: number
): ResolvedPosition | null {
  const projection = index.projection(rootStory);
  if (!projection || !Number.isFinite(position) || position < 0 || position > projection.size) {
    return null;
  }
  const target = projection.targetAt(position);
  const block = projection.blockBoundaryAt(position);
  if (block) return { target, loc: block };
  const map = index.inputMap(target.story);
  const loc = map ? displayPositionToYrsLoc(map, target.displayPosition) : null;
  return loc ? { target, loc } : null;
}

/** The paragraph's story units as text, one U+FFFC per embed, so indices are loc offsets. */
function paragraphUnits(index: DisplayPositionIndex, loc: YrsLoc): string | null {
  let text = '';
  for (const segment of index.reader.storySegments(loc.story)) {
    if (segment.kind === 'pilcrow') {
      if (segment.paraId === loc.paraId) return text;
      text = '';
    } else {
      text += segment.kind === 'text' ? segment.text : '￼';
    }
  }
  return null;
}

/** @internal The word, paragraph or whole story around `position`, in display positions. */
export function resolveSelectionUnit(
  index: DisplayPositionIndex,
  rootStory: string,
  position: number,
  unit: DocxSelectionUnit,
  expectVersion: string
): DocxDisplayRange | null {
  if (index.reader.version() !== expectVersion) return null;
  if (unit === 'story') {
    const projection = index.projection(rootStory);
    return projection && projection.size > 0 ? { anchor: 0, head: projection.size } : null;
  }
  const resolved = resolvePosition(index, rootStory, position);
  if (!resolved) return null;
  const { loc } = resolved;
  const text = paragraphUnits(index, loc);
  if (text === null) return null;
  const [start, end] = unit === 'word' ? findWordBoundaries(text, loc.offset) : [0, text.length];
  if (unit === 'word' && start >= end) return null;
  const anchor = index.positionOf({ ...loc, offset: start }, rootStory);
  const head = index.positionOf({ ...loc, offset: end }, rootStory);
  return anchor === null || head === null ? null : { anchor, head };
}

function sameTable(a: YrsCellLoc, b: YrsCellLoc): boolean {
  return a.story === b.story && a.tableIndex === b.tableIndex;
}

/** The `tableIndex`-th table embed of `story` and its story offset. */
function tableEmbed(
  index: DisplayPositionIndex,
  story: string,
  tableIndex: number
): { offset: number; rows: TablePayloadRow[]; grid?: unknown[] } | null {
  let offset = 0;
  let tables = 0;
  for (const segment of index.reader.storySegments(story)) {
    if (segment.kind === 'embed' && segment.embedKind === 'table') {
      if (tables === tableIndex) {
        const { rows, grid } = segment.payload;
        return Array.isArray(rows)
          ? { offset, rows: rows as TablePayloadRow[], ...(Array.isArray(grid) ? { grid } : {}) }
          : null;
      }
      tables += 1;
    }
    offset += segment.kind === 'text' ? segment.text.length : 1;
  }
  return null;
}

/** The top-level table of `rootStory` holding `position`, if any. */
function rootTableAt(index: DisplayPositionIndex, rootStory: string, position: number) {
  const projection = index.projection(rootStory);
  let table = projection?.tableAtPosition(position) ?? null;
  while (table && table.story !== rootStory) {
    table = projection!.tableAtPosition(table.start - 1);
  }
  return table;
}

/** The root story offset where a selection end at `position` falls; a table end takes the whole table. */
function rootStoryOffset(
  index: DisplayPositionIndex,
  rootStory: string,
  resolved: ResolvedPosition,
  position: number,
  side: 'start' | 'end'
): number | null {
  if (resolved.loc.story === rootStory) return storyOffsetForLoc(index.reader, resolved.loc);
  const table = rootTableAt(index, rootStory, position);
  const embed = table ? tableEmbed(index, rootStory, table.tableIndex) : null;
  if (!embed) return null;
  return side === 'start' ? embed.offset : embed.offset + 1;
}

/** @internal The plain text and accepted-view range of the selection `[anchor, head]`. */
export function resolveSelectionText(
  index: DisplayPositionIndex,
  rootStory: string,
  anchor: number,
  head: number,
  expectVersion: string
): DocxDisplaySelectionText | null {
  const reader = index.reader;
  if (reader.version() !== expectVersion) return null;
  const from = Math.min(anchor, head);
  const to = Math.max(anchor, head);
  const start = resolvePosition(index, rootStory, from);
  const end = resolvePosition(index, rootStory, to);
  if (!start || !end) return null;
  if (start.loc.story === end.loc.story) {
    const story = start.loc.story;
    const startOffset = acceptedOffsetOf(reader, start.loc);
    const endOffset = acceptedOffsetOf(reader, end.loc);
    return {
      text: storyPlainText(
        reader,
        story,
        storyOffsetForLoc(reader, start.loc),
        storyOffsetForLoc(reader, end.loc)
      ),
      range:
        startOffset === null || endOffset === null
          ? null
          : {
              story,
              start: { paraId: start.loc.paraId, offset: startOffset },
              end: { paraId: end.loc.paraId, offset: endOffset },
              view: 'accepted',
            },
    };
  }
  const startCell = start.target.cell;
  const endCell = end.target.cell;
  if (startCell && endCell && sameTable(startCell, endCell)) {
    const table = tableEmbed(index, startCell.story, startCell.tableIndex);
    if (!table) return null;
    const text = tablePlainText(reader, table, {
      top: Math.min(startCell.row, endCell.row),
      bottom: Math.max(startCell.row, endCell.row),
      left: Math.min(startCell.column, endCell.column),
      right: Math.max(startCell.column, endCell.column),
    }).join('\n');
    return { text, range: null };
  }
  const startOffset = rootStoryOffset(index, rootStory, start, from, 'start');
  const endOffset = rootStoryOffset(index, rootStory, end, to, 'end');
  if (startOffset === null || endOffset === null) return null;
  return { text: storyPlainText(reader, rootStory, startOffset, endOffset), range: null };
}

/** @internal The selection info for a display range. */
export function resolveSelectionInfo(
  index: DisplayPositionIndex,
  rootStory: string,
  anchor: number,
  head: number,
  expectVersion: string
): DocxDisplaySelectionInfo | null {
  const reader = index.reader;
  if (reader.version() !== expectVersion) return null;
  const a = resolvePosition(index, rootStory, anchor);
  const b = resolvePosition(index, rootStory, head);
  if (!a || !b) return null;
  if (a.loc.story !== b.loc.story) {
    const selection = resolveSelectionText(index, rootStory, anchor, head, expectVersion);
    return selection ? {
      paraId: null, selectedText: selection.text, paragraphText: '', before: '', after: '',
    } : null;
  }
  const [start, end] =
    storyOffsetForLoc(reader, a.loc) <= storyOffsetForLoc(reader, b.loc)
      ? [a.loc, b.loc]
      : [b.loc, a.loc];
  try {
    return reader.selectionText({
      story: start.story,
      start: { paraId: start.paraId, offset: start.offset },
      end: { paraId: end.paraId, offset: end.offset },
    });
  } catch {
    return null;
  }
}

function storyOffsetToLoc(index: DisplayPositionIndex, story: string, offset: number): YrsLoc | null {
  const paragraphs = index.inputMap(story)?.paragraphs;
  if (!paragraphs?.length) return null;
  for (const paragraph of paragraphs) {
    const span = index.reader.locateParagraph(story, paragraph.paraId);
    if (offset <= span.end) {
      return {
        story,
        paraId: paragraph.paraId,
        offset: Math.min(Math.max(0, offset - span.start), span.end - span.start),
      };
    }
  }
  const last = paragraphs[paragraphs.length - 1];
  const span = index.reader.locateParagraph(story, last.paraId);
  return { story, paraId: last.paraId, offset: span.end - span.start };
}

function displayRange(
  index: DisplayPositionIndex,
  rootStory: string,
  start: YrsLoc,
  end: YrsLoc
): DocxDisplayRange | null {
  const projection = index.projection(rootStory);
  if (!projection) return null;
  const anchor = projection.positionForLoc(start);
  const head = projection.positionForLoc(end);
  return anchor === null || head === null ? null : { anchor, head };
}

/** @internal The paragraph range across stories, mapped into the root layout. */
export function resolveParagraphTarget(
  index: DisplayPositionIndex,
  rootStory: string,
  paraId: string,
  expectVersion: string
): DocxDisplayRange | null {
  if (index.reader.version() !== expectVersion) return null;
  for (const story of index.reader.storyIds()) {
    if (!index.reader.paragraphs(story).some((paragraph) => paragraph.paraId === paraId)) continue;
    const span = index.reader.locateParagraph(story, paraId);
    return displayRange(index, rootStory, { story, paraId, offset: 0 },
      { story, paraId, offset: Math.max(0, span.end - span.start) });
  }
  return null;
}

/** @internal The first comment anchor in the root story's display positions. */
export function resolveCommentTarget(
  index: DisplayPositionIndex,
  rootStory: string,
  commentId: string,
  expectVersion: string
): DocxDisplayRange | null {
  if (index.reader.version() !== expectVersion) return null;
  try {
    const anchor = index.reader.resolveComment(commentId)[0];
    if (!anchor) return null;
    const start = storyOffsetToLoc(index, anchor.story, anchor.start);
    const end = storyOffsetToLoc(index, anchor.story, anchor.end);
    return start && end ? displayRange(index, rootStory, start, end) : null;
  } catch {
    return null;
  }
}

/** @internal The revision range in the root story's display positions. */
export function resolveRevisionTarget(
  index: DisplayPositionIndex,
  rootStory: string,
  revisionId: string,
  expectVersion: string
): DocxDisplayRange | null {
  if (index.reader.version() !== expectVersion) return null;
  const revision = index.reader.listRevisions().find((candidate) => candidate.revisionId === revisionId);
  return revision
    ? displayRange(
        index,
        rootStory,
        { story: revision.story, ...revision.range.start },
        { story: revision.story, ...revision.range.end }
      )
    : null;
}

/** @internal The display position of bookmark `name` in root story `rootStory`, or null. */
export function resolveBookmarkPosition(
  index: DisplayPositionIndex,
  rootStory: string,
  name: string,
  expectVersion: string
): number | null {
  if (index.reader.version() !== expectVersion) return null;
  return index.projection(rootStory)?.bookmarkPosition(name) ?? null;
}
