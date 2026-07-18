import type {
  YrsCellBorders,
  YrsCellLoc,
  YrsContentControlValue,
  YrsImageGeometry,
  YrsLoc,
  YrsParagraphAttrs,
  YrsSession,
  YrsStoryRange,
  YrsTableLoc,
  YrsTableRange,
} from '@betteroffice/docx/yrs';
import { computeSplitDialogDefaults, pixelsToEmu } from '@betteroffice/docx/utils';
import type { ImageLayoutTarget, SetImageWrapTypeOptions } from '@betteroffice/docx/docx';
import type { TableContextInfo } from './types';

/** Non-toolbar writes that become yrs-authoritative with `?yrsInput=1`. */
export type YrsEditorCommand =
  | { type: 'imageGeometry'; pmPos: number; patch: Readonly<Record<string, unknown>> }
  | {
      type: 'imageWrap';
      pmPos: number;
      target: ImageLayoutTarget;
      options?: SetImageWrapTypeOptions;
    }
  | {
      type: 'imageTransform';
      pmPos: number;
      action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV';
    }
  | { type: 'insertImage'; image: Readonly<Record<string, unknown>> }
  | {
      type: 'contentControlValue';
      pmPos: number;
      embedId?: string;
      value: YrsContentControlValue;
    }
  | { type: 'insertPageBreak' }
  | {
      type: 'insertSectionBreak';
      breakType: 'nextPage' | 'continuous' | 'oddPage' | 'evenPage';
    }
  | { type: 'paragraphAttrs'; attrs: YrsParagraphAttrs }
  | { type: 'removeTabStop'; positionTwips: number }
  | {
      type: 'setHyperlink';
      href: string;
      tooltip?: string;
      displayText?: string;
      editExisting?: boolean;
      matchHref?: string;
    }
  | { type: 'removeHyperlink'; href?: string }
  | { type: 'insertTable'; rows: number; columns: number }
  | { type: 'tableInsertRow'; side: 'above' | 'below'; at?: YrsCellLoc }
  | { type: 'tableInsertColumn'; side: 'left' | 'right'; at?: YrsCellLoc }
  | { type: 'tableDeleteRow' }
  | { type: 'tableDeleteColumn' }
  | { type: 'tableDelete' }
  | { type: 'tableMergeCells' }
  | { type: 'tableSplitCell'; rows?: number; columns?: number }
  | { type: 'tableCellShading'; color: string | null }
  | {
      type: 'tableProperties';
      properties: {
        width?: number | null;
        widthType?: string | null;
        justification?: 'left' | 'center' | 'right' | null;
      };
    }
  | { type: 'tableSetBorders'; borders: YrsCellBorders }
  | {
      type: 'tableColumnWidths';
      pmStart: number;
      widths: ReadonlyArray<{ column: number; widthTwips: number }>;
    }
  | { type: 'tableSelect'; target: 'table' | 'row' | 'column' };

export interface YrsTableTarget {
  focused: YrsCellLoc;
  range: YrsTableRange;
}

export interface YrsHyperlinkHit {
  href: string;
  tooltip?: string;
  text: string;
  range: YrsStoryRange;
}

export function yrsStoryOffsetForLoc(session: YrsSession, loc: YrsLoc): number {
  const span = session.locateParagraph(loc.story, loc.paraId);
  return span.start + loc.offset;
}

export function yrsLocForStoryOffset(
  session: YrsSession,
  story: string,
  offset: number
): YrsLoc | null {
  const paragraphs = session.paragraphs(story);
  for (const paragraph of paragraphs) {
    const span = session.locateParagraph(story, paragraph.paraId);
    if (offset <= span.end) {
      return {
        story,
        paraId: paragraph.paraId,
        offset: Math.min(Math.max(0, offset - span.start), span.end - span.start),
      };
    }
  }
  return null;
}

/** Current sticky selection normalized into document order. */
export function currentYrsSelectionRange(session: YrsSession): YrsStoryRange | null {
  const selection = session.selection();
  if (!selection || selection.anchor.story !== selection.head.story) return null;
  const anchorOffset = yrsStoryOffsetForLoc(session, selection.anchor);
  const headOffset = yrsStoryOffsetForLoc(session, selection.head);
  const [start, end] =
    anchorOffset <= headOffset
      ? [selection.anchor, selection.head]
      : [selection.head, selection.anchor];
  return {
    story: start.story,
    start: { paraId: start.paraId, offset: start.offset },
    end: { paraId: end.paraId, offset: end.offset },
  };
}

interface FlatStorySegment {
  start: number;
  end: number;
  text: string;
  attributes: Record<string, unknown>;
}

function flatStorySegments(session: YrsSession, story: string): FlatStorySegment[] {
  const result: FlatStorySegment[] = [];
  let offset = 0;
  for (const segment of session.storySegments(story)) {
    const length = segment.kind === 'text' ? segment.text.length : 1;
    result.push({
      start: offset,
      end: offset + length,
      text: segment.kind === 'text' ? segment.text : segment.kind === 'pilcrow' ? '\n' : '',
      attributes: segment.attributes,
    });
    offset += length;
  }
  return result;
}

function objectValue(value: unknown): Record<string, unknown> | null {
  return value != null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function hyperlinkAttrs(segment: FlatStorySegment): Record<string, unknown> | null {
  return objectValue(segment.attributes.hyperlink);
}

function sameHyperlink(segment: FlatStorySegment, href: string): boolean {
  return String(hyperlinkAttrs(segment)?.href ?? '') === href;
}

/** Resolve the hyperlink under the current Yrs caret/selection. */
export function yrsHyperlinkAtSelection(
  session: YrsSession,
  matchHref?: string
): YrsHyperlinkHit | null {
  const selection = session.selection();
  if (!selection) return null;
  const story = selection.head.story;
  const offset = yrsStoryOffsetForLoc(session, selection.head);
  const segments = flatStorySegments(session, story);
  let index = segments.findIndex((segment) => {
    const link = hyperlinkAttrs(segment);
    if (!link || (matchHref && String(link.href ?? '') !== matchHref)) return false;
    return offset > segment.start && offset < segment.end;
  });
  if (index < 0) {
    index = segments.findIndex((segment) => {
      const link = hyperlinkAttrs(segment);
      if (!link || (matchHref && String(link.href ?? '') !== matchHref)) return false;
      return offset === segment.start || offset === segment.end;
    });
  }
  if (index < 0) return null;
  const attrs = hyperlinkAttrs(segments[index]);
  const href = String(attrs?.href ?? '');
  if (!href) return null;
  let first = index;
  let last = index;
  while (first > 0 && segments[first - 1].end === segments[first].start && sameHyperlink(segments[first - 1], href)) first -= 1;
  while (last + 1 < segments.length && segments[last].end === segments[last + 1].start && sameHyperlink(segments[last + 1], href)) last += 1;
  const start = yrsLocForStoryOffset(session, story, segments[first].start);
  const end = yrsLocForStoryOffset(session, story, segments[last].end);
  if (!start || !end) return null;
  return {
    href,
    ...(typeof attrs?.tooltip === 'string' ? { tooltip: attrs.tooltip } : {}),
    text: segments.slice(first, last + 1).map((segment) => segment.text).join(''),
    range: {
      story,
      start: { paraId: start.paraId, offset: start.offset },
      end: { paraId: end.paraId, offset: end.offset },
    },
  };
}

/** Plain text covered by the current Yrs selection. */
export function yrsSelectedText(session: YrsSession): string {
  const range = currentYrsSelectionRange(session);
  if (!range) return '';
  const start = yrsStoryOffsetForLoc(session, { story: range.story, ...range.start });
  const end = yrsStoryOffsetForLoc(session, { story: range.story, ...range.end });
  if (start === end) return '';
  return flatStorySegments(session, range.story)
    .map((segment) => {
      const from = Math.max(start, segment.start);
      const to = Math.min(end, segment.end);
      return from < to ? segment.text.slice(from - segment.start, to - segment.start) : '';
    })
    .join('');
}

interface TablePayloadCell {
  story: string;
  tcPr?: Record<string, unknown>;
}

interface TablePayloadRow {
  cells: TablePayloadCell[];
}

interface TablePayload {
  tblPr?: Record<string, unknown>;
  grid?: unknown[];
  rows: TablePayloadRow[];
}

interface TableCellAnchor {
  row: number;
  column: number;
  rowspan: number;
  colspan: number;
  story: string;
}

function positiveSpan(value: unknown): number {
  return typeof value === 'number' && Number.isInteger(value) && value > 0 ? value : 1;
}

function tablePayload(session: YrsSession, table: YrsTableLoc): TablePayload | null {
  let tableIndex = 0;
  for (const segment of session.storySegments(table.story)) {
    if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
    if (tableIndex === table.tableIndex) {
      const rows = segment.payload.rows;
      if (!Array.isArray(rows)) return null;
      return {
        tblPr:
          segment.payload.tblPr && typeof segment.payload.tblPr === 'object'
            ? (segment.payload.tblPr as Record<string, unknown>)
            : undefined,
        grid: Array.isArray(segment.payload.grid) ? segment.payload.grid : undefined,
        rows: rows as TablePayloadRow[],
      };
    }
    tableIndex += 1;
  }
  return null;
}

/** Current table properties for the properties dialog. */
export function currentYrsTableProperties(
  session: YrsSession
): Record<string, unknown> | undefined {
  const target = currentYrsTableTarget(session);
  return target ? tablePayload(session, target.focused)?.tblPr : undefined;
}

function tableAnchors(payload: TablePayload): { anchors: TableCellAnchor[]; columns: number } {
  const occupied: boolean[][] = Array.from({ length: payload.rows.length }, () => []);
  const anchors: TableCellAnchor[] = [];
  let columns = payload.grid?.length ?? 0;

  payload.rows.forEach((row, rowIndex) => {
    let column = 0;
    for (const cell of row.cells ?? []) {
      while (occupied[rowIndex]?.[column]) column += 1;
      const rowspan = positiveSpan(cell.tcPr?.rowspan);
      const colspan = positiveSpan(cell.tcPr?.colspan);
      anchors.push({ row: rowIndex, column, rowspan, colspan, story: cell.story });
      for (let targetRow = rowIndex; targetRow < rowIndex + rowspan; targetRow += 1) {
        const slots = occupied[targetRow] ?? [];
        occupied[targetRow] = slots;
        for (let targetColumn = column; targetColumn < column + colspan; targetColumn += 1) {
          slots[targetColumn] = true;
        }
      }
      column += colspan;
      columns = Math.max(columns, column);
    }
  });

  return { anchors, columns };
}

function sameTable(a: YrsTableLoc, b: YrsTableLoc): boolean {
  return a.story === b.story && a.tableIndex === b.tableIndex;
}

function sameCell(a: YrsCellLoc, b: YrsCellLoc): boolean {
  return sameTable(a, b) && a.row === b.row && a.column === b.column;
}

function storyContains(parent: string, story: string): boolean {
  return story === parent || story.startsWith(`${parent}:`);
}

/** Resolve the innermost table-cell identity encoded in a yrs story id. */
export function yrsCellLocFromStory(story: string): YrsCellLoc | null {
  const matches = Array.from(story.matchAll(/:t(\d+):r(\d+)c(\d+)(?=:|$)/g));
  const match = matches[matches.length - 1];
  if (!match || match.index == null) return null;
  const tableIndex = Number(match[1]);
  const row = Number(match[2]);
  const column = Number(match[3]);
  if (![tableIndex, row, column].every(Number.isSafeInteger)) return null;
  return { story: story.slice(0, match.index), tableIndex, row, column };
}

/** Resolve a current grid cell back to its stable independent story id. */
export function yrsCellStory(session: YrsSession, at: YrsCellLoc): string | null {
  const payload = tablePayload(session, at);
  if (!payload) return null;
  const { anchors } = tableAnchors(payload);
  return (
    anchors.find(
      (anchor) =>
        anchor.row <= at.row &&
        at.row < anchor.row + anchor.rowspan &&
        anchor.column <= at.column &&
        at.column < anchor.column + anchor.colspan
    )?.story ?? null
  );
}

/**
 * Resolve the table target owned by the current yrs text/cell selections.
 * The cell selection is sticky by stable story identity, so it remains correct
 * after preceding row/column edits even though authored story ids do not rename.
 */
export function currentYrsTableTarget(session: YrsSession): YrsTableTarget | null {
  const selection = session.selection();
  if (!selection || !yrsCellLocFromStory(selection.head.story)) return null;

  let selected: YrsTableRange | null;
  try {
    selected = session.cellSelection();
  } catch {
    // Undo/remote deletion can invalidate the ephemeral cell selection before
    // the hidden input repairs its caret to a surviving body paragraph.
    return null;
  }
  if (selected && sameTable(selected.anchor, selected.head)) {
    const anchorStory = yrsCellStory(session, selected.anchor);
    const headStory = yrsCellStory(session, selected.head);
    const focused =
      anchorStory && storyContains(anchorStory, selection.head.story)
        ? selected.anchor
        : headStory && storyContains(headStory, selection.head.story)
          ? selected.head
          : selected.anchor;
    const range =
      sameCell(focused, selected.anchor) || sameCell(focused, selected.head)
        ? selected
        : { anchor: focused, head: focused };
    return { focused, range };
  }

  // Compatibility fallback for programmatic selections that predate the
  // pointer adapter's cell-selection publication.
  const focused = yrsCellLocFromStory(selection.head.story);
  return focused ? { focused, range: { anchor: focused, head: focused } } : null;
}

function cellBorderColor(tcPr: Record<string, unknown> | undefined): TableContextInfo['cellBorderColor'] {
  const borders = tcPr?.borders;
  if (!borders || typeof borders !== 'object') return undefined;
  for (const value of Object.values(borders)) {
    if (!value || typeof value !== 'object') continue;
    const color = (value as { color?: unknown }).color;
    if (typeof color === 'string') return { rgb: color.replace(/^#/, '') };
    if (color && typeof color === 'object') {
      return color as NonNullable<TableContextInfo['cellBorderColor']>;
    }
  }
  return undefined;
}

/** Build the toolbar/context-menu table state from the authoritative yrs payload. */
export function currentYrsTableContext(session: YrsSession): TableContextInfo | null {
  const target = currentYrsTableTarget(session);
  if (!target) return null;
  const payload = tablePayload(session, target.focused);
  if (!payload) return null;
  const { anchors, columns } = tableAnchors(payload);
  const focusedAnchor = anchors.find(
    (anchor) =>
      anchor.row <= target.focused.row &&
      target.focused.row < anchor.row + anchor.rowspan &&
      anchor.column <= target.focused.column &&
      target.focused.column < anchor.column + anchor.colspan
  );
  const focusedCell = focusedAnchor
    ? payload.rows
        .flatMap((row) => row.cells ?? [])
        .find((cell) => cell.story === focusedAnchor.story)
    : undefined;
  const justification = payload.tblPr?.justification;
  const backgroundColor = focusedCell?.tcPr?.backgroundColor;
  return {
    isInTable: true,
    table:
      typeof justification === 'string'
        ? { attrs: { justification } }
        : undefined,
    rowIndex: target.focused.row,
    columnIndex: target.focused.column,
    rowCount: payload.rows.length,
    columnCount: columns,
    hasMultiCellSelection:
      target.range.anchor.row !== target.range.head.row ||
      target.range.anchor.column !== target.range.head.column,
    canSplitCell: true,
    cellBorderColor: cellBorderColor(focusedCell?.tcPr),
    cellBackgroundColor:
      typeof backgroundColor === 'string' ? backgroundColor.replace(/^#/, '') : undefined,
  };
}

/** Resolve the live cell's minimum and suggested split-dialog dimensions. */
export function currentYrsSplitCellConfig(session: YrsSession): {
  initialRows: number;
  initialCols: number;
  minRows: number;
  minCols: number;
} | null {
  const target = currentYrsTableTarget(session);
  if (!target) return null;
  const payload = tablePayload(session, target.focused);
  if (!payload) return null;
  const anchor = tableAnchors(payload).anchors.find(
    (candidate) =>
      candidate.row <= target.focused.row &&
      target.focused.row < candidate.row + candidate.rowspan &&
      candidate.column <= target.focused.column &&
      target.focused.column < candidate.column + candidate.colspan
  );
  return anchor ? computeSplitDialogDefaults(anchor.rowspan, anchor.colspan) : null;
}

/** Build the row/column/table range used by selection commands. */
export function yrsTableSelectionRange(
  session: YrsSession,
  focused: YrsCellLoc,
  target: 'table' | 'row' | 'column'
): YrsTableRange | null {
  const payload = tablePayload(session, focused);
  if (!payload || payload.rows.length === 0) return null;
  const { columns } = tableAnchors(payload);
  if (columns === 0) return null;
  const endRow = payload.rows.length - 1;
  const endColumn = columns - 1;
  if (target === 'row') {
    return {
      anchor: { ...focused, column: 0 },
      head: { ...focused, column: endColumn },
    };
  }
  if (target === 'column') {
    return {
      anchor: { ...focused, row: 0 },
      head: { ...focused, row: endRow },
    };
  }
  return {
    anchor: { ...focused, row: 0, column: 0 },
    head: { ...focused, row: endRow, column: endColumn },
  };
}

/** Author a table-wide property directly on its structural yrs embed. */
export function setYrsTableProperty(
  session: YrsSession,
  table: YrsTableLoc,
  key: string,
  value: unknown
): boolean {
  let offset = 0;
  let tableIndex = 0;
  for (const segment of session.storySegments(table.story)) {
    if (segment.kind === 'text') {
      offset += segment.text.length;
      continue;
    }
    if (segment.kind === 'pilcrow') {
      offset += 1;
      continue;
    }
    if (segment.embedKind === 'table') {
      if (tableIndex === table.tableIndex) {
        const current =
          segment.payload[key] && typeof segment.payload[key] === 'object'
            ? (segment.payload[key] as Record<string, unknown>)
            : {};
        session.applyRawOps(table.story, [
          { op: 'setEmbedAttr', index: offset, key, value: { ...current, ...(value as object) } },
        ]);
        return true;
      }
      tableIndex += 1;
    }
    offset += 1;
  }
  return false;
}

/** Move the sticky text caret into a currently resolved table cell. */
export function setYrsSelectionInCell(session: YrsSession, at: YrsCellLoc): boolean {
  const story = yrsCellStory(session, at);
  if (!story) return false;
  const paragraph = session.paragraphs(story)[0];
  if (!paragraph) return false;
  session.setSelection({ story, paraId: paragraph.paraId, offset: 0 });
  return true;
}

/** Pick a surviving parent-story paragraph near a table before deleting it. */
export function yrsSelectionNearTable(session: YrsSession, table: YrsTableLoc) {
  const segments = session.storySegments(table.story);
  let tableIndex = 0;
  let targetIndex = -1;
  for (let index = 0; index < segments.length; index += 1) {
    const segment = segments[index];
    if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
    if (tableIndex === table.tableIndex) {
      targetIndex = index;
      break;
    }
    tableIndex += 1;
  }
  if (targetIndex < 0) return null;
  const following = segments.slice(targetIndex + 1).find((segment) => segment.kind === 'pilcrow');
  const preceding = segments
    .slice(0, targetIndex)
    .reverse()
    .find((segment) => segment.kind === 'pilcrow');
  const paragraph = following ?? preceding;
  return paragraph?.kind === 'pilcrow'
    ? { story: table.story, paraId: paragraph.paraId, offset: 0 }
    : null;
}

/**
 * Run native history without leaving the live caret inside a cell story that
 * the structural transaction may delete. A surviving cell selection is
 * restored after the transaction; otherwise the caret stays beside the table.
 */
export function performYrsHistoryAction(session: YrsSession, redo: boolean): boolean {
  const before = session.selection();
  const cell = before ? yrsCellLocFromStory(before.head.story) : null;
  const nearby = cell ? yrsSelectionNearTable(session, cell) : null;
  if (nearby) session.setSelection(nearby);

  const changed = redo ? session.redo() : session.undo();
  if (!changed || !before || !nearby) return changed;

  const restoredCell = yrsCellLocFromStory(before.head.story);
  const cellIsLive =
    !restoredCell || yrsCellStory(session, restoredCell) === before.head.story;
  if (!cellIsLive) return changed;
  try {
    const paragraphs = session.paragraphs(before.head.story);
    const paraIds = new Set(paragraphs.map((paragraph) => paragraph.paraId));
    if (paraIds.has(before.anchor.paraId) && paraIds.has(before.head.paraId)) {
      session.setSelection(before.anchor, before.head);
    }
  } catch {
    // The structural history operation removed the selected story.
  }
  return changed;
}

function authoredId(value: unknown): string | null {
  if (typeof value === 'string') return value.length > 0 ? value : null;
  if (typeof value === 'number' && Number.isFinite(value)) return String(value);
  return null;
}

export interface YrsProjectedEmbedNode {
  kind: string;
  attrs: Record<string, unknown>;
}

/** Resolve the same stable payload keys accepted by the Rust embed-id helper. */
export function yrsEmbedIdForProjectedNode(node: YrsProjectedEmbedNode): string | null {
  const explicit = authoredId(node.attrs.embedId);
  if (explicit) return explicit;
  if (node.kind === 'image') return authoredId(node.attrs.rId);
  if (node.kind === 'sdt' || node.kind === 'blockSdt') {
    return authoredId(node.attrs.id);
  }
  return authoredId(node.attrs.id) ?? authoredId(node.attrs.rId);
}

function finiteNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function nullableEmu(value: unknown): number | null {
  const number = finiteNumber(value);
  return number == null ? null : pixelsToEmu(number);
}

function nullableString(value: unknown): string | null {
  return typeof value === 'string' && value.length > 0 ? value : null;
}

function normalizedPatch(patch: Readonly<Record<string, unknown>>): Record<string, unknown | null> {
  return Object.fromEntries(
    Object.entries(patch).map(([key, value]) => [key, value === undefined ? null : value])
  );
}

/**
 * Translate projected image attrs into the typed facade geometry payload.
 *
 * The typed EMU fields are the authored command contract. `other` carries the
 * coexistence payload vocabulary used by today's render projection until that
 * bridge is retired; this keeps the canvas and the faithful PM follower on the
 * same geometry without dispatching a PM transaction.
 */
export function yrsImageGeometryForProjectedNode(
  node: YrsProjectedEmbedNode,
  patch: Readonly<Record<string, unknown>>
): YrsImageGeometry | null {
  if (node.kind !== 'image') return null;
  const next = { ...node.attrs, ...patch } as Record<string, unknown>;
  const width = finiteNumber(next.width) ?? 0;
  const height = finiteNumber(next.height) ?? 0;
  const wrap = next.wrapType;
  const position =
    next.position !== null && typeof next.position === 'object'
      ? (next.position as {
          horizontal?: { posOffset?: unknown; relativeTo?: unknown };
          vertical?: { posOffset?: unknown; relativeTo?: unknown };
        })
      : undefined;

  return {
    widthEmu: pixelsToEmu(width),
    heightEmu: pixelsToEmu(height),
    ...(wrap === 'inline' ||
    wrap === 'square' ||
    wrap === 'tight' ||
    wrap === 'through' ||
    wrap === 'topAndBottom' ||
    wrap === 'behind' ||
    wrap === 'inFront'
      ? { wrap }
      : {}),
    hOffsetEmu: finiteNumber(position?.horizontal?.posOffset),
    vOffsetEmu: finiteNumber(position?.vertical?.posOffset),
    distTopEmu: nullableEmu(next.distTop),
    distBottomEmu: nullableEmu(next.distBottom),
    distLeftEmu: nullableEmu(next.distLeft),
    distRightEmu: nullableEmu(next.distRight),
    relativeFromHorizontal: nullableString(position?.horizontal?.relativeTo),
    relativeFromVertical: nullableString(position?.vertical?.relativeTo),
    other: normalizedPatch(patch),
  };
}

/** Resolve a rotate/flip toolbar action against current projected image attrs. */
export function yrsImageTransformForProjectedNode(
  node: YrsProjectedEmbedNode,
  action: Extract<YrsEditorCommand, { type: 'imageTransform' }>['action']
): string | null {
  if (node.kind !== 'image') return null;
  const current = typeof node.attrs.transform === 'string' ? node.attrs.transform : '';
  const rotateMatch = current.match(/rotate\((-?\d+(?:\.\d+)?)deg\)/);
  let rotation = rotateMatch ? Number.parseFloat(rotateMatch[1]) : 0;
  let flipH = /scaleX\(-1\)/.test(current);
  let flipV = /scaleY\(-1\)/.test(current);
  if (action === 'rotateCW') rotation = (rotation + 90) % 360;
  else if (action === 'rotateCCW') rotation = (rotation - 90 + 360) % 360;
  else if (action === 'flipH') flipH = !flipH;
  else flipV = !flipV;

  const parts: string[] = [];
  if (rotation !== 0) parts.push(`rotate(${rotation}deg)`);
  if (flipH) parts.push('scaleX(-1)');
  if (flipV) parts.push('scaleY(-1)');
  return parts.length > 0 ? parts.join(' ') : null;
}
