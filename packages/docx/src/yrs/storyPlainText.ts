import type { YrsLoc, YrsSession } from './index';

/** @internal */
export type PlainTextReader = Pick<YrsSession, 'storySegments'>;

/** @internal */
export const MAX_TABLE_COLUMNS = 16_384;
const PLAIN_TEXT_SLOT_LIMIT = 1 << 20;

interface TablePayloadCell {
  story: string;
  tcPr?: Record<string, unknown>;
}

/** @internal */
export interface TablePayloadRow {
  cells: TablePayloadCell[];
}

/** @internal */
export interface TablePayload {
  tblPr?: Record<string, unknown>;
  grid?: unknown[];
  rows: TablePayloadRow[];
}

/** @internal */
export interface TableCellAnchor {
  row: number;
  column: number;
  rowspan: number;
  colspan: number;
  story: string;
}

interface TableVerticalMerge {
  start: number;
  end: number;
  lastRow: number;
}

function positiveSpan(value: unknown): number {
  return typeof value === 'number' && Number.isInteger(value) && value > 0 ? value : 1;
}

function objectValue(value: unknown): Record<string, unknown> | null {
  return value != null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function embedPlainText(reader: PlainTextReader, kind: string, payload: Record<string, unknown>): string {
  switch (kind) {
    case 'tab':
      return '\t';
    case 'break':
      return '\n';
    case 'field':
      return typeof payload.displayText === 'string' ? payload.displayText : '';
    case 'math':
      return typeof payload.plainText === 'string' ? payload.plainText : '';
    case 'sdt':
      return Array.isArray(payload.content)
        ? payload.content
            .map((item: { kind?: unknown; text?: unknown; payload?: unknown }) =>
              item.kind === 'text' && typeof item.text === 'string'
                ? item.text
                : embedPlainText(reader, String(item.kind), objectValue(item.payload) ?? {})
            )
            .join('')
        : '';
    case 'blockSdt':
      return typeof payload.story === 'string'
        ? `${storyPlainText(reader, payload.story).replace(/\n$/, '')}\n`
        : '';
    case 'table':
      return Array.isArray(payload.rows)
        ? tablePlainText(reader, { rows: payload.rows as TablePayloadRow[] })
            .map((row) => `${row}\n`)
            .join('')
        : '';
    default:
      return '';
  }
}

/** A cell as one tab-separated field, quoted as spreadsheets do when it holds a tab, break or quote. */
function cellPlainText(reader: PlainTextReader, story: string): string {
  const text = storyPlainText(reader, story).replace(/\n$/, '');
  return /[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

/** Tab-separated rows, without grid padding for oversized blocks. @internal */
export function tablePlainText(
  reader: PlainTextReader,
  payload: TablePayload,
  range?: { top: number; bottom: number; left: number; right: number }
): string[] {
  const { anchors, columns } = tableAnchors(payload);
  const top = range?.top ?? 0;
  const bottom = range?.bottom ?? payload.rows.length - 1;
  const left = range?.left ?? 0;
  const right = range?.right ?? columns - 1;
  const lines: string[] = [];
  if ((bottom - top + 1) * (right - left + 1) > PLAIN_TEXT_SLOT_LIMIT) {
    let index = 0;
    for (let row = top; row <= bottom; row += 1) {
      const texts: string[] = [];
      while (index < anchors.length && anchors[index].row < row) index += 1;
      while (index < anchors.length && anchors[index].row === row) {
        const cell = anchors[index++];
        if (cell.column >= left && cell.column <= right) {
          texts.push(cell.story ? cellPlainText(reader, cell.story) : '');
        }
      }
      lines.push(texts.join('\t'));
    }
    return lines;
  }
  const byGrid = new Map(anchors.map((cell) => [`${cell.row}:${cell.column}`, cell.story]));
  for (let row = top; row <= bottom; row += 1) {
    const texts: string[] = [];
    for (let column = left; column <= right; column += 1) {
      const story = byGrid.get(`${row}:${column}`);
      texts.push(story ? cellPlainText(reader, story) : '');
    }
    lines.push(texts.join('\t'));
  }
  return lines;
}

/** Plain text of story units `[from, to)`: paragraphs end in newlines, tables become tab-separated rows. @internal */
export function storyPlainText(reader: PlainTextReader, story: string, from = 0, to = Infinity): string {
  let text = '';
  let offset = 0;
  for (const segment of reader.storySegments(story)) {
    const start = offset;
    offset += segment.kind === 'text' ? segment.text.length : 1;
    if (offset <= from) continue;
    if (start >= to) break;
    if (segment.kind === 'text') {
      text += segment.text.slice(Math.max(from, start) - start, Math.min(to, offset) - start);
    } else if (segment.kind === 'pilcrow') {
      text += '\n';
    } else {
      text += embedPlainText(reader, segment.embedKind, segment.payload);
    }
  }
  return text;
}

/** @internal */
export function tableAnchors(payload: TablePayload): { anchors: TableCellAnchor[]; columns: number } {
  let activeMerges: TableVerticalMerge[] = [];
  const anchors: TableCellAnchor[] = [];
  let columns = Math.min(payload.grid?.length ?? 0, MAX_TABLE_COLUMNS);

  payload.rows.forEach((row, rowIndex) => {
    activeMerges = activeMerges.filter((merge) => merge.lastRow >= rowIndex);
    const nextMerges: TableVerticalMerge[] = [];
    let mergeIndex = 0;
    let column = 0;
    for (const cell of row.cells ?? []) {
      while (mergeIndex < activeMerges.length && activeMerges[mergeIndex].start <= column) {
        const merge = activeMerges[mergeIndex++];
        column = Math.max(column, merge.end);
        nextMerges.push(merge);
      }
      const rowspan = Math.max(
        1,
        Math.min(positiveSpan(cell.tcPr?.rowspan), payload.rows.length - rowIndex)
      );
      const colspan = Math.max(
        1,
        Math.min(positiveSpan(cell.tcPr?.colspan), MAX_TABLE_COLUMNS - column)
      );
      anchors.push({ row: rowIndex, column, rowspan, colspan, story: cell.story });
      if (rowspan > 1) {
        nextMerges.push({ start: column, end: column + colspan, lastRow: rowIndex + rowspan - 1 });
      }
      column += colspan;
      columns = Math.max(columns, column);
    }
    while (mergeIndex < activeMerges.length) nextMerges.push(activeMerges[mergeIndex++]);
    activeMerges = nextMerges;
  });

  return { anchors, columns };
}

/** @internal Story offset of `loc`: the paragraph's span start plus the loc offset. */
export function storyOffsetForLoc(reader: Pick<YrsSession, 'locateParagraph'>, loc: YrsLoc): number {
  const span = reader.locateParagraph(loc.story, loc.paraId);
  return span.start + loc.offset;
}
