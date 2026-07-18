import type { YrsCellLoc, YrsLoc, YrsSession, YrsStorySegment } from '@betteroffice/docx/yrs';

export interface YrsProjectedNode {
  kind: string;
  start: number;
  nodeSize: number;
  attrs: Record<string, unknown>;
  story: string;
}

export interface YrsProjectedTable extends YrsProjectedNode {
  kind: 'table';
  tableIndex: number;
  widthsTwips: number[];
  rowCount: number;
  cells: YrsProjectedCell[];
}

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
  paragraphs: Array<{ paraId: string; displayStart: number; length: number }>;
  tables: YrsProjectedTable[];
}

export interface YrsPointerProjectionTarget {
  story: string;
  displayPosition: number;
  cell?: YrsCellLoc;
}

export class YrsPositionProjection {
  private readonly stories = new Map<string, StoryProjection>();
  private readonly nodes = new Map<number, YrsProjectedNode>();
  private readonly tables: YrsProjectedTable[] = [];

  constructor(
    private readonly session: YrsSession,
    readonly rootStory: string
  ) {
    this.buildStory(rootStory, 0, 0);
  }

  get size(): number {
    return this.stories.get(this.rootStory)?.size ?? 0;
  }

  nodeAt(position: number): YrsProjectedNode | null {
    return this.nodes.get(position) ?? null;
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
      displayPosition: Math.max(0, position - (target?.contentStart ?? 0)),
      ...(target?.cell ? { cell: target.cell } : {}),
    };
  }

  positionForLoc(loc: YrsLoc): number | null {
    const story = this.stories.get(loc.story);
    const paragraph = story?.paragraphs.find((candidate) => candidate.paraId === loc.paraId);
    if (!story || !paragraph) return null;
    return (
      story.contentStart +
      paragraph.displayStart +
      1 +
      Math.min(Math.max(0, loc.offset), paragraph.length)
    );
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

    const segments = this.session.storySegments(storyId);
    let cursor = 0;
    let inlineLength = 0;
    let paragraphStart = 0;
    let tableIndex = 0;
    for (const segment of segments) {
      if (segment.kind === 'text') {
        inlineLength += segment.text.length;
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
        story.paragraphs.push({
          paraId: segment.paraId,
          displayStart: paragraphStart,
          length: inlineLength,
        });
        cursor = paragraphStart + nodeSize;
        paragraphStart = cursor;
        inlineLength = 0;
        continue;
      }

      if (segment.embedKind === 'table') {
        const table = this.buildTable(segment, storyId, tableIndex, contentStart + cursor, depth);
        story.tables.push(table);
        this.tables.push(table);
        this.nodes.set(table.start, table);
        tableIndex += 1;
        cursor += table.nodeSize;
        paragraphStart = cursor;
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
        cursor += node.nodeSize;
        paragraphStart = cursor;
        continue;
      }
      if (segment.embedKind === 'pageBreak' && inlineLength === 0) {
        const node: YrsProjectedNode = {
          kind: 'pageBreak',
          start: contentStart + cursor,
          nodeSize: 1,
          attrs: segment.payload,
          story: storyId,
        };
        this.nodes.set(node.start, node);
        cursor += 1;
        paragraphStart = cursor;
        continue;
      }

      const start = contentStart + paragraphStart + 1 + inlineLength;
      this.nodes.set(start, {
        kind: segment.embedKind,
        start,
        nodeSize: 1,
        attrs: { ...segment.payload, ...segment.attributes },
        story: storyId,
      });
      inlineLength += 1;
    }

    story.size = Math.max(cursor, paragraphStart + (inlineLength > 0 ? inlineLength + 2 : 0));
    return story;
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
