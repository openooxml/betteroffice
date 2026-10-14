import { headerFooterStory } from './sessionInternals';
import type { YrsCellLoc, YrsLoc, YrsSession, YrsStorySegment } from './index';
import type { PointPosition } from '../plugin-api';
import {
  createYrsInputPositionMap,
  yrsLocToDisplayPosition as yrsLocToLocalDisplayPosition,
  type YrsInputPositionMap,
} from './inputPositionMap';

type PositionReader = Pick<YrsSession, 'hasStory' | 'storySegments'>;

export type YrsLocProjection = Pick<YrsPositionProjection, 'positionForLoc'>;

export type YrsPositionOutline = Record<string, {
  contentStart: number;
  size: number;
  /** `length` counts display units; `controls` lists inline controls as `[story offset, display units]`. */
  paragraphs: Array<{ paraId: string; displayStart: number; length: number; leading: number; controls?: InlineControls }>;
}>;

/** Inline controls of a paragraph as `[story offset within the content, display units]`, in order. */
type InlineControls = ReadonlyArray<readonly [number, number]>;

/** Display units of an inline control with `payload`: its content and its two boundaries, as the layout counts them. */
function inlineControlDisplaySize(payload: Record<string, unknown>): number {
  if (!Array.isArray(payload.content)) return 2;
  let size = 2;
  for (const child of payload.content as unknown[]) {
    if (!child || typeof child !== 'object' || Array.isArray(child)) continue;
    const item = child as Record<string, unknown>;
    if (item.kind === 'text') size += typeof item.text === 'string' ? item.text.length : 0;
    else if (item.kind !== 'sdt') size += 1;
    else size += item.payload && typeof item.payload === 'object' && !Array.isArray(item.payload)
      ? inlineControlDisplaySize(item.payload as Record<string, unknown>) : 2;
  }
  return size;
}

/** Display units from a paragraph's content start to story offset `offset` in it. */
function controlDisplayOffset(offset: number, controls: InlineControls | undefined): number {
  let expanded = 0;
  for (const [at, size] of controls ?? []) {
    if (at >= offset) break;
    expanded += size - 1;
  }
  return offset + expanded;
}

/** Story offset at display units `display` from a paragraph's content start; inside a control, its nearer edge. */
function controlStoryOffset(display: number, controls: InlineControls): number {
  let expanded = 0;
  for (const [at, size] of controls) {
    const start = at + expanded;
    if (display <= start) break;
    if (display < start + size) return display - start < start + size - display ? at : at + 1;
    expanded += size - 1;
  }
  return display - expanded;
}

export function createYrsLocProjectionFromOutline(outline: YrsPositionOutline): YrsLocProjection {
  const stories = new Map(Object.entries(outline).map(([id, story]) => {
    const paragraphs = new Map<string, (typeof story.paragraphs)[number]>();
    for (const paragraph of story.paragraphs) {
      if (!paragraphs.has(paragraph.paraId)) paragraphs.set(paragraph.paraId, paragraph);
    }
    return [id, { contentStart: story.contentStart, paragraphs }] as const;
  }));
  return {
    positionForLoc: (loc) => {
      const story = stories.get(loc.story);
      const paragraph = story?.paragraphs.get(loc.paraId);
      if (!story || !paragraph) return null;
      return story.contentStart + paragraph.displayStart + 1 +
        Math.min(controlDisplayOffset(Math.max(0, loc.offset - paragraph.leading), paragraph.controls), paragraph.length);
    },
  };
}

/** @internal */
export interface YrsProjectedNode {
  kind: string;
  start: number;
  nodeSize: number;
  attrs: Record<string, unknown>;
  story: string;
}

/** @internal */
export interface YrsProjectedTable extends YrsProjectedNode {
  kind: 'table';
  tableIndex: number;
  widthsTwips: number[];
  rowCount: number;
  cells: YrsProjectedCell[];
}

/** @internal */
export interface YrsProjectedCell {
  row: number;
  column: number;
  rowspan: number;
  colspan: number;
  story: string;
  start: number;
  contentStart: number;
  nodeSize: number;
  attrs: Record<string, unknown>;
}

interface StoryProjection {
  story: string;
  contentStart: number;
  size: number;
  depth: number;
  cell?: YrsCellLoc;
  paragraphs: Array<{
    paraId: string;
    displayStart: number;
    inputStart: number;
    /** Display units of the content. */
    length: number;
    /** Leading block embeds counted by story offsets. */
    leading: number;
    controls: InlineControls;
  }>;
  tables: YrsProjectedTable[];
}

/** @internal */
export interface YrsPointerProjectionTarget {
  story: string;
  displayPosition: number;
  cell?: YrsCellLoc;
}

/** @internal */
export interface YrsStorySegmentSource {
  segments(story: string): YrsStorySegment[];
}

/** @internal */
export function createYrsPositionProjection(
  session: PositionReader,
  rootStory: string,
  source?: YrsStorySegmentSource
): YrsPositionProjection | null {
  if (!session.hasStory(rootStory)) return null;
  return new YrsPositionProjection(session, rootStory, source);
}

/** @internal */
export function projectYrsDisplayPosition(
  position: number | PointPosition,
  getProjection: (rootStory: string) => YrsPositionProjection | null,
  session?: object | null
): YrsPointerProjectionTarget | null {
  let rootStory = 'body';
  if (typeof position !== 'number') {
    switch (position.region) {
      case 'header':
      case 'footer':
        if (!position.rId) return null;
        rootStory = headerFooterStory(session, position.rId);
        break;
      case 'footnote':
      case 'endnote':
        if (position.noteId === undefined) return null;
        rootStory = `${position.region === 'footnote' ? 'fn' : 'en'}:${position.noteId}`;
        break;
    }
  }
  const offset = typeof position === 'number' ? position : position.position;
  const projection = getProjection(rootStory);
  if (!projection || !Number.isFinite(offset) || offset < 0 || offset > projection.size) {
    return null;
  }
  return projection.targetAt(offset);
}

/** @internal */
export class YrsPositionProjection {
  private readonly stories = new Map<string, StoryProjection>();
  private readonly nodes = new Map<number, YrsProjectedNode>();
  private readonly tables: YrsProjectedTable[] = [];
  private readonly blockBoundaries = new Map<number, YrsLoc>();
  private readonly source: YrsStorySegmentSource;
  private readonly paragraphsById = new Map<
    StoryProjection,
    Map<string, StoryProjection['paragraphs'][number]>
  >();

  constructor(
    session: Pick<YrsSession, 'storySegments'>,
    readonly rootStory: string,
    source?: YrsStorySegmentSource
  ) {
    this.source = source ?? { segments: (story) => session.storySegments(story) };
    this.buildStory(rootStory, 0, 0);
  }

  get size(): number {
    return this.stories.get(this.rootStory)?.size ?? 0;
  }

  nodeAt(position: number): YrsProjectedNode | null {
    return this.nodes.get(position) ?? null;
  }

  allTables(): readonly YrsProjectedTable[] {
    return this.tables;
  }

  tableAtPosition(position: number): YrsProjectedTable | null {
    let result: YrsProjectedTable | null = null;
    for (const table of this.tables) {
      if (position < table.start || position > table.start + table.nodeSize) continue;
      if (!result || table.nodeSize < result.nodeSize) result = table;
    }
    return result;
  }

  tableAtStart(start: number): YrsProjectedTable | null {
    return this.tables.find((table) => table.start === start) ?? null;
  }

  tableStartForLoc(at: YrsCellLoc): number | null {
    return (
      this.stories
        .get(at.story)
        ?.tables.find((table) => table.tableIndex === at.tableIndex)?.start ?? null
    );
  }

  cellPosition(tableStart: number, row: number, column: number): number | null {
    const table = this.tableAtStart(tableStart);
    if (!table) return null;
    const cell = table.cells.find(
      (candidate) =>
        candidate.row <= row &&
        row < candidate.row + candidate.rowspan &&
        candidate.column <= column &&
        column < candidate.column + candidate.colspan
    );
    return cell?.start ?? null;
  }

  targetAt(position: number): YrsPointerProjectionTarget {
    let target = this.stories.get(this.rootStory) ?? null;
    for (const story of this.stories.values()) {
      if (
        position >= story.contentStart &&
        position <= story.contentStart + story.size &&
        (!target || story.depth > target.depth)
      ) {
        target = story;
      }
    }
    return {
      story: target?.story ?? this.rootStory,
      displayPosition: target ? this.inputPositionAt(target, position) : 0,
      ...(target?.cell ? { cell: target.cell } : {}),
    };
  }

  positionForLoc(loc: YrsLoc): number | null {
    const story = this.stories.get(loc.story);
    const paragraph = story ? this.paragraphIndex(story).get(loc.paraId) : undefined;
    if (!story || !paragraph) return null;
    return (
      story.contentStart +
      paragraph.displayStart +
      1 +
      Math.min(
        controlDisplayOffset(Math.max(0, loc.offset - paragraph.leading), paragraph.controls),
        paragraph.length
      )
    );
  }

  /** The location of a block embed's start or end at `position`, in the offsets of the paragraph after it. */
  blockBoundaryAt(position: number): YrsLoc | null {
    return this.blockBoundaries.get(position) ?? null;
  }

  private setBlockBoundary(position: number, loc: YrsLoc): void {
    if (!this.blockBoundaries.has(position)) this.blockBoundaries.set(position, loc);
  }

  /** First paragraph for each paraId. */
  private paragraphIndex(
    story: StoryProjection
  ): Map<string, StoryProjection['paragraphs'][number]> {
    let index = this.paragraphsById.get(story);
    if (!index) {
      index = new Map();
      for (const paragraph of story.paragraphs) {
        if (!index.has(paragraph.paraId)) index.set(paragraph.paraId, paragraph);
      }
      this.paragraphsById.set(story, index);
    }
    return index;
  }

  bookmarkPosition(name: string): number | null {
    for (const story of this.stories.values()) {
      for (const paragraph of story.paragraphs) {
        const node = this.nodes.get(story.contentStart + paragraph.displayStart);
        const bookmarks = node?.attrs.bookmarks;
        if (
          Array.isArray(bookmarks) &&
          bookmarks.some(
            (bookmark) =>
              bookmark != null &&
              typeof bookmark === 'object' &&
              String((bookmark as Record<string, unknown>).name ?? '') === name
          )
        ) {
          return story.contentStart + paragraph.displayStart;
        }
      }
    }
    return null;
  }

  private buildStory(
    storyId: string,
    contentStart: number,
    depth: number,
    cell?: YrsCellLoc
  ): StoryProjection {
    const existing = this.stories.get(storyId);
    if (existing) return existing;
    const story: StoryProjection = {
      story: storyId,
      contentStart,
      size: 0,
      depth,
      cell,
      paragraphs: [],
      tables: [],
    };
    this.stories.set(storyId, story);

    const segments = this.source.segments(storyId);
    let cursor = 0;
    let inlineLength = 0;
    let inlineUnits = 0;
    let controls: Array<[number, number]> = [];
    let leading = 0;
    let paragraphStart = 0;
    let inputStart = 0;
    let tableIndex = 0;
    let blocks: Array<{ start: number; size: number; offset: number }> = [];
    for (const segment of segments) {
      if (segment.kind === 'text') {
        inlineLength += segment.text.length;
        inlineUnits += segment.text.length;
        continue;
      }
      if (segment.kind === 'pilcrow') {
        const nodeSize = inlineLength + 2;
        const start = contentStart + paragraphStart;
        const node: YrsProjectedNode = {
          kind: 'paragraph',
          start,
          nodeSize,
          attrs: { ...segment.properties, paraId: segment.paraId },
          story: storyId,
        };
        this.nodes.set(start, node);
        for (const block of blocks) {
          this.setBlockBoundary(block.start, { story: storyId, paraId: segment.paraId, offset: block.offset });
          this.setBlockBoundary(block.start + block.size, { story: storyId, paraId: segment.paraId, offset: block.offset + 1 });
        }
        blocks = [];
        story.paragraphs.push({
          paraId: segment.paraId,
          displayStart: paragraphStart,
          inputStart,
          length: inlineLength,
          leading,
          controls,
        });
        inputStart += inlineUnits + 2 + leading;
        cursor = paragraphStart + nodeSize;
        paragraphStart = cursor;
        inlineLength = 0;
        inlineUnits = 0;
        controls = [];
        leading = 0;
        continue;
      }

      if (segment.embedKind === 'table') {
        const table = this.buildTable(segment, storyId, tableIndex, contentStart + cursor, depth);
        story.tables.push(table);
        this.tables.push(table);
        this.nodes.set(table.start, table);
        blocks.push({ start: table.start, size: table.nodeSize, offset: leading });
        tableIndex += 1;
        cursor += table.nodeSize;
        paragraphStart = cursor;
        leading += 1;
        continue;
      }
      if (segment.embedKind === 'blockSdt') {
        const childStory = String(segment.payload.story ?? `${storyId}:sdt0`);
        const start = contentStart + cursor;
        const child = this.buildStory(childStory, start + 1, depth + 1, cell);
        const node: YrsProjectedNode = {
          kind: 'blockSdt',
          start,
          nodeSize: child.size + 2,
          attrs: segment.payload,
          story: storyId,
        };
        this.nodes.set(start, node);
        blocks.push({ start, size: node.nodeSize, offset: leading });
        cursor += node.nodeSize;
        paragraphStart = cursor;
        leading += 1;
        continue;
      }
      if (
        (segment.embedKind === 'pageBreak' || segment.embedKind === 'columnBreak') &&
        inlineLength === 0
      ) {
        const node: YrsProjectedNode = {
          kind: segment.embedKind,
          start: contentStart + cursor,
          nodeSize: 1,
          attrs: segment.payload,
          story: storyId,
        };
        this.nodes.set(node.start, node);
        blocks.push({ start: node.start, size: node.nodeSize, offset: leading });
        cursor += 1;
        paragraphStart = cursor;
        leading += 1;
        continue;
      }

      const start = contentStart + paragraphStart + 1 + inlineLength;
      const nodeSize = segment.embedKind === 'sdt' ? inlineControlDisplaySize(segment.payload) : 1;
      this.nodes.set(start, {
        kind: segment.embedKind,
        start,
        nodeSize,
        attrs: { ...segment.payload, ...segment.attributes },
        story: storyId,
      });
      if (segment.embedKind === 'sdt') controls.push([inlineUnits, nodeSize]);
      inlineLength += nodeSize;
      inlineUnits += 1;
    }

    story.size = Math.max(cursor, paragraphStart + (inlineLength > 0 ? inlineLength + 2 : 0));
    return story;
  }

  private inputPositionAt(story: StoryProjection, position: number): number {
    const projected = Math.max(0, position - story.contentStart);
    let paragraph = story.paragraphs[0];
    for (const candidate of story.paragraphs) {
      if (candidate.displayStart > projected) break;
      paragraph = candidate;
    }
    if (!paragraph) return 0;
    const display = Math.min(Math.max(0, projected - paragraph.displayStart), paragraph.length + 1);
    return (
      paragraph.inputStart +
      paragraph.leading +
      (display === 0 ? 0 : 1 + controlStoryOffset(display - 1, paragraph.controls))
    );
  }

  private buildTable(
    segment: Extract<YrsStorySegment, { kind: 'embed' }>,
    storyId: string,
    tableIndex: number,
    start: number,
    depth: number
  ): YrsProjectedTable {
    const rows = Array.isArray(segment.payload.rows)
      ? (segment.payload.rows as Array<Record<string, unknown>>)
      : [];
    const occupied: boolean[][] = Array.from({ length: rows.length }, () => []);
    const cells: YrsProjectedCell[] = [];
    let tableContentSize = 0;
    rows.forEach((rowPayload, rowIndex) => {
      const rowStart = start + 1 + tableContentSize;
      const payloadCells = Array.isArray(rowPayload.cells)
        ? (rowPayload.cells as Array<Record<string, unknown>>)
        : [];
      let rowContentSize = 0;
      let column = 0;
      payloadCells.forEach((cellPayload, cellIndex) => {
        while (occupied[rowIndex]?.[column]) column += 1;
        const attrs = asObject(cellPayload.tcPr);
        const rowspan = positiveSpan(attrs.rowspan);
        const colspan = positiveSpan(attrs.colspan);
        const childStory = String(
          cellPayload.story ?? `${storyId}:t${tableIndex}:r${rowIndex}c${cellIndex}`
        );
        const cellStart = rowStart + 1 + rowContentSize;
        const cellLoc: YrsCellLoc = { story: storyId, tableIndex, row: rowIndex, column };
        const child = this.buildStory(childStory, cellStart + 1, depth + 1, cellLoc);
        const nodeSize = child.size + 2;
        cells.push({
          row: rowIndex,
          column,
          rowspan,
          colspan,
          story: childStory,
          start: cellStart,
          contentStart: cellStart + 1,
          nodeSize,
          attrs,
        });
        for (let targetRow = rowIndex; targetRow < rowIndex + rowspan; targetRow += 1) {
          occupied[targetRow] ??= [];
          for (let targetColumn = column; targetColumn < column + colspan; targetColumn += 1) {
            occupied[targetRow][targetColumn] = true;
          }
        }
        rowContentSize += nodeSize;
        column += colspan;
      });
      tableContentSize += rowContentSize + 2;
    });
    return {
      kind: 'table',
      start,
      nodeSize: tableContentSize + 2,
      attrs: segment.payload,
      story: storyId,
      tableIndex,
      widthsTwips: Array.isArray(segment.payload.grid)
        ? segment.payload.grid.filter((width): width is number => typeof width === 'number')
        : [],
      rowCount: rows.length,
      cells,
    };
  }
}

function asObject(value: unknown): Record<string, unknown> {
  return value != null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};
}

function positiveSpan(value: unknown): number {
  return typeof value === 'number' && Number.isInteger(value) && value > 0 ? value : 1;
}

/** @internal */
export function yrsLocToProjectedDisplayPosition(
  reader: Pick<YrsSession, 'hasStory' | 'paragraphSpans'>,
  projectionFor: (rootStory: string) => YrsLocProjection | null,
  loc: YrsLoc,
  activeRootStory = 'body',
  inputMap?: (story: string) => YrsInputPositionMap | null
): number | null {
  const map = inputMap
    ? inputMap(loc.story)
    : reader.hasStory(loc.story)
      ? createYrsInputPositionMap(loc.story, reader.paragraphSpans(loc.story))
      : null;
  if (!map) return null;
  const rootStory =
    loc.story === 'body' || loc.story.startsWith('body:') ? 'body' : activeRootStory;
  return (
    projectionFor(rootStory)?.positionForLoc(loc) ??
    (loc.story === rootStory ? yrsLocToLocalDisplayPosition(map, loc) : null)
  );
}
