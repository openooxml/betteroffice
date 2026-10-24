import type {
  YrsAuthor,
  YrsLoc,
  YrsSession,
  YrsStoryRange,
  YrsStorySegment,
  YrsTableRange,
  YrsInlineFormatDelta,
} from './index';
import {
  embedPlainText,
  storyPlainText,
  tableAnchors,
  tablePlainText,
  type TablePayload,
} from './storyPlainText';

/** @internal Version of the internal clipboard payload; a reader ignores other versions. */
export const DOCX_CLIPBOARD_FRAGMENT_VERSION = 2;

/**
 * @internal One unit of a table cell's story. `attributes` are raw yrs run or embed attributes as
 * `storySegments` reports them; `format` is a semantic patch applied on top (HTML import uses it).
 */
export type DocxFragmentSegment =
  | { kind: 'text'; text: string; attributes?: Record<string, unknown>; format?: YrsInlineFormatDelta }
  | { kind: 'tab' | 'break'; payload?: Record<string, unknown>; attributes?: Record<string, unknown> }
  | { kind: 'pilcrow'; properties?: Record<string, unknown>; attributes?: Record<string, unknown> }
  | { kind: 'table'; table: DocxTableFragment; attributes?: Record<string, unknown> };

/** @internal A cell's `content` ends with a pilcrow; `shading` is a hex fill applied like `setCellShading`. */
export interface DocxTableFragmentCell {
  tcPr: Record<string, unknown>;
  content: DocxFragmentSegment[];
  shading?: string;
}

/** @internal */
export interface DocxTableFragmentRow {
  trPr: Record<string, unknown>;
  cells: DocxTableFragmentCell[];
}

/** @internal A rectangular table, its cells in row order with `rowspan`/`colspan` in `tcPr`. */
export interface DocxTableFragment {
  tblPr: Record<string, unknown>;
  /** Column widths in twips; empty splits the default table width evenly. */
  grid: number[];
  rows: DocxTableFragmentRow[];
}

export type DocxClipboardInline = { text: string } | { image: Record<string, unknown> };

/** @internal A paragraph's leading tables and ordered inline content. */
export interface DocxClipboardParagraph {
  tables?: DocxTableFragment[];
  text: string;
  inlines?: DocxClipboardInline[];
}

/** @internal Copied or pasted content in document order; consecutive paragraphs are separate paragraphs. */
export interface DocxClipboardFragment {
  paragraphs: DocxClipboardParagraph[];
}

/** @internal What a copy writes: plain text plus a fragment for tables or pictures. */
export interface DocxClipboardSlice {
  text: string;
  fragment: DocxClipboardFragment | null;
}

/** @internal */
export function clipboardFragmentHasPictures(fragment: DocxClipboardFragment | null): boolean {
  return fragment?.paragraphs.some(paragraph => paragraph.inlines?.some(inline => 'image' in inline)) ?? false;
}

const MAX_FRAGMENT_CELLS = 10_000;
const MAX_FRAGMENT_PARAGRAPHS = 1_000;
const MAX_FRAGMENT_TEXT = 1_000_000;
const MAX_FRAGMENT_DEPTH = 8;
const MAX_FRAGMENT_MEDIA = 64 * 1024 * 1024;
const IMAGE_DATA_URL = /^data:image\/[a-z0-9.+-]+;base64,[a-z0-9+/]+={0,2}$/iu;
const IMAGE_IDENTITY_KEYS = new Set(['rId', '_kind', 'id', 'embedId', 'docPrId', 'drawingId']);

function independentImage(payload: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(payload).filter(([key]) => !IMAGE_IDENTITY_KEYS.has(key)));
}

function validImage(value: unknown): value is Record<string, unknown> {
  return plainObject(value) && jsonValue(value) && !Object.hasOwn(value, '_kind') &&
    typeof value.src === 'string' && value.src.length <= MAX_FRAGMENT_MEDIA && IMAGE_DATA_URL.test(value.src) &&
    ['widthEmu', 'heightEmu'].every(key => value[key] === undefined ||
      (typeof value[key] === 'number' && Number.isFinite(value[key]) && value[key] > 0));
}
const LINE_BREAK = /[\n\r\u000B\u2028\u2029]/u;

function plainObject(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function jsonValue(value: unknown, depth = 0): boolean {
  if (depth > 32) return false;
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return true;
  if (typeof value === 'number') return Number.isFinite(value);
  if (Array.isArray(value)) return value.every(entry => jsonValue(entry, depth + 1));
  return plainObject(value) && Object.values(value).every(entry => jsonValue(entry, depth + 1));
}

function optionalMap(value: unknown): boolean {
  return value === undefined || (plainObject(value) && jsonValue(value));
}

/** @internal Strict shape and size check for a pasted or copied clipboard fragment. */
export function validClipboardFragment(value: unknown): value is DocxClipboardFragment {
  const budget = { cells: MAX_FRAGMENT_CELLS, text: MAX_FRAGMENT_TEXT, media: MAX_FRAGMENT_MEDIA };
  const text = (entry: unknown, allowEmpty: boolean): boolean => {
    if (typeof entry !== 'string' || (!allowEmpty && entry === '') || LINE_BREAK.test(entry)) return false;
    budget.text -= entry.length;
    return budget.text >= 0;
  };
  const spanOf = (tcPr: Record<string, unknown>, key: 'rowspan' | 'colspan'): number | null => {
    const span = tcPr[key];
    if (span === undefined || span === null) return 1;
    if (typeof span !== 'number' || !Number.isInteger(span) || span < 1 || span > 1000) return null;
    return span;
  };
  const formatOk = (format: unknown): boolean => {
    if (format === undefined) return true;
    if (!plainObject(format) || !jsonValue(format)) return false;
    const boolOrNull = (entry: unknown): boolean => entry === undefined || entry === null || typeof entry === 'boolean';
    if (!boolOrNull(format.bold) || !boolOrNull(format.italic)) return false;
    const underline = format.underline;
    if (underline !== undefined && underline !== null && typeof underline !== 'boolean' &&
      !(plainObject(underline) &&
        (underline.style === undefined || typeof underline.style === 'string') &&
        (underline.color === undefined || typeof underline.color === 'string'))) return false;
    const strike = format.strike;
    if (strike !== undefined && strike !== null && typeof strike !== 'boolean' &&
      !(plainObject(strike) && (strike.double === undefined || typeof strike.double === 'boolean'))) return false;
    const color = format.color;
    if (color !== undefined && color !== null) {
      if (!plainObject(color)) return false;
      const rgb = typeof color.rgb === 'string';
      const theme = typeof color.themeColor === 'string';
      if (rgb === theme) return false;
      if (color.rgb !== undefined && !rgb) return false;
      if (color.themeColor !== undefined && !theme) return false;
    }
    if (format.highlight !== undefined && format.highlight !== null && typeof format.highlight !== 'string') return false;
    const fontSize = format.fontSize;
    if (fontSize !== undefined && fontSize !== null &&
      (typeof fontSize !== 'number' || !Number.isFinite(fontSize) || fontSize <= 0)) return false;
    const fontFamily = format.fontFamily;
    if (fontFamily !== undefined && fontFamily !== null &&
      !(plainObject(fontFamily) && typeof fontFamily.ascii === 'string' &&
        (fontFamily.hAnsi === undefined || typeof fontFamily.hAnsi === 'string'))) return false;
    const other = format.other;
    if (other !== undefined &&
      !(plainObject(other) && !('ins' in other) && !('del' in other) && !('hyperlink' in other))) return false;
    return true;
  };
  const segment = (entry: unknown, depth: number): boolean => {
    if (!plainObject(entry) || !optionalMap(entry.attributes)) return false;
    switch (entry.kind) {
      case 'text': return text(entry.text, false) && formatOk(entry.format);
      case 'tab': case 'break': return optionalMap(entry.payload);
      case 'pilcrow': return optionalMap(entry.properties);
      case 'table': return table(entry.table, depth + 1);
      default: return false;
    }
  };
  const table = (entry: unknown, depth: number): boolean => {
    if (depth > MAX_FRAGMENT_DEPTH || !plainObject(entry) || !optionalMap(entry.tblPr) || entry.tblPr === undefined) return false;
    if (!Array.isArray(entry.grid) || !entry.grid.every(width => typeof width === 'number' && Number.isFinite(width) && width > 0)) return false;
    if (!Array.isArray(entry.rows) || entry.rows.length === 0 || entry.rows.length > 1000) return false;
    const rowsOk = entry.rows.every(row => plainObject(row) && row.trPr !== undefined && optionalMap(row.trPr) &&
      Array.isArray(row.cells) && row.cells.length <= 1000 && row.cells.every(cell => {
        if (!plainObject(cell) || cell.tcPr === undefined || !optionalMap(cell.tcPr)) return false;
        const tcPr = cell.tcPr as Record<string, unknown>;
        if (spanOf(tcPr, 'rowspan') === null || spanOf(tcPr, 'colspan') === null) return false;
        if (cell.shading !== undefined && (typeof cell.shading !== 'string' || !/^[0-9a-f]{6}$/iu.test(cell.shading))) return false;
        if (--budget.cells < 0 || !Array.isArray(cell.content) || cell.content.length === 0) return false;
        const last = cell.content.at(-1);
        return plainObject(last) && last.kind === 'pilcrow' && cell.content.every(item => segment(item, depth));
      }));
    if (!rowsOk) return false;
    const rows = entry.rows as Array<{ cells: Array<{ tcPr: Record<string, unknown> }> }>;
    const occupied: boolean[][] = rows.map(() => []);
    let columns = 0;
    let cellCount = 0;
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      let column = 0;
      for (const cell of rows[rowIndex]!.cells) {
        while (occupied[rowIndex]![column]) column += 1;
        const rowspan = spanOf(cell.tcPr, 'rowspan')!;
        const colspan = spanOf(cell.tcPr, 'colspan')!;
        if (rowIndex + rowspan > rows.length || column + colspan > 1000) return false;
        for (let r = rowIndex; r < rowIndex + rowspan; r += 1) {
          for (let c = column; c < column + colspan; c += 1) {
            if (occupied[r]![c]) return false;
            occupied[r]![c] = true;
          }
        }
        cellCount += 1;
        column += colspan;
        columns = Math.max(columns, column);
      }
    }
    if (cellCount === 0) return false;
    if (entry.grid.length > 0 && entry.grid.length !== columns) return false;
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      if (rows[rowIndex]!.cells.length === 0) {
        for (let column = 0; column < columns; column += 1) {
          if (!occupied[rowIndex]![column]) return false;
        }
      }
    }
    return true;
  };
  if (!plainObject(value) || !Array.isArray(value.paragraphs) || value.paragraphs.length === 0 ||
    value.paragraphs.length > MAX_FRAGMENT_PARAGRAPHS) return false;
  return value.paragraphs.every(paragraph => {
    if (!plainObject(paragraph) || !text(paragraph.text, true) ||
      !(paragraph.tables === undefined || (Array.isArray(paragraph.tables) && paragraph.tables.every(entry => table(entry, 1))))) return false;
    if (paragraph.inlines === undefined) return true;
    if (!Array.isArray(paragraph.inlines) || paragraph.inlines.length > MAX_FRAGMENT_TEXT) return false;
    let plain = '';
    for (const inline of paragraph.inlines) {
      if (!plainObject(inline) || Object.keys(inline).length !== 1) return false;
      if ('text' in inline) {
        if (typeof inline.text !== 'string' || inline.text === '' || LINE_BREAK.test(inline.text)) return false;
        plain += inline.text;
      } else if ('image' in inline) {
        if (!validImage(inline.image)) return false;
        budget.media -= (inline.image.src as string).length;
        if (budget.media < 0) return false;
      } else return false;
    }
    return plain === paragraph.text;
  });
}

/** @internal Sessions a table read needs: the raw payload plus story segments. */
export type ClipboardTableReader = Pick<YrsSession, 'tablePayload' | 'storySegments'> & { mediaDataUrl?(source: string): string | null | undefined };

/** @internal Engine surface a fragment paste needs. */
export type ClipboardPasteEngine = Pick<
  YrsSession,
  'deleteRange' | 'splitParagraph' | 'insertText' | 'insertTableFragment' | 'validateTableFragment' | 'insertImage'
>;

/** @internal Inclusive grid bounds for {@link readTableFragment}. */
export interface ClipboardTableRect {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

/** @internal Copy source: a multi-cell range, or story units `[from, to)`. */
export type ClipboardSliceTarget = { cells: YrsTableRange } | { story: string; from: number; to: number };

function record(value: unknown): Record<string, unknown> | null {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function copiedMap(value: unknown): Record<string, unknown> | undefined {
  const map = record(value);
  return map && Object.keys(map).length > 0 ? { ...map } : undefined;
}

interface RawCell {
  story: string;
  tcPr?: unknown;
}

/** @internal Raw table payload lowered to the anchor reader's shape, or null. */
export function clipboardTablePayload(value: unknown): TablePayload | null {
  const map = record(value);
  if (!map || !Array.isArray(map.rows)) return null;
  return {
    ...(record(map.tblPr) ? { tblPr: map.tblPr as Record<string, unknown> } : {}),
    ...(Array.isArray(map.grid) ? { grid: map.grid } : {}),
    rows: map.rows as TablePayload['rows'],
  };
}

function rawRows(payload: { rows?: unknown }): Array<{ cells?: unknown; trPr?: unknown }> {
  if (!payload || !Array.isArray(payload.rows) || payload.rows.length === 0) {
    throw new Error('Table payload holds no rows');
  }
  return payload.rows as Array<{ cells?: unknown; trPr?: unknown }>;
}

function cellContent(reader: ClipboardTableReader, story: string, onUnsupported?: () => void): DocxFragmentSegment[] {
  const content: DocxFragmentSegment[] = [];
  for (const segment of reader.storySegments(story)) {
    if (segment.kind === 'text') {
      if (segment.text === '' || segment.attributes?.del != null) continue;
      content.push({
        kind: 'text',
        text: segment.text,
        ...(Object.keys(segment.attributes).length > 0 ? { attributes: { ...segment.attributes } } : {}),
      });
    } else if (segment.kind === 'pilcrow') {
      content.push({
        kind: 'pilcrow',
        ...(Object.keys(segment.properties).length > 0 ? { properties: { ...segment.properties } } : {}),
        ...(Object.keys(segment.attributes).length > 0 ? { attributes: { ...segment.attributes } } : {}),
      });
    } else if (segment.embedKind === 'tab' || segment.embedKind === 'break') {
      content.push({
        kind: segment.embedKind,
        ...(Object.keys(segment.payload).length > 0 ? { payload: { ...segment.payload } } : {}),
        ...(Object.keys(segment.attributes).length > 0 ? { attributes: { ...segment.attributes } } : {}),
      });
    } else if (segment.embedKind === 'table') {
      content.push({
        kind: 'table',
        table: fragmentFromPayload(reader, segment.payload, undefined, onUnsupported),
        ...(Object.keys(segment.attributes).length > 0 ? { attributes: { ...segment.attributes } } : {}),
      });
    } else {
      for (const part of embedTextSegments(reader, segment, onUnsupported)) content.push(part);
    }
  }
  return content;
}

function embedTextSegments(
  reader: ClipboardTableReader,
  segment: Extract<YrsStorySegment, { kind: 'embed' }>,
  onUnsupported?: () => void,
): DocxFragmentSegment[] {
  const out: DocxFragmentSegment[] = [];
  const text = embedPlainText(reader, segment.embedKind, segment.payload);
  if (onUnsupported && !clipboardTextEmbed(reader, segment.embedKind, segment.payload, text)) onUnsupported();
  const parts = text.split('\n');
  const attributes = Object.keys(segment.attributes).length > 0 ? { ...segment.attributes } : undefined;
  parts.forEach((text, index) => {
    if (index > 0) out.push({ kind: 'pilcrow' });
    if (text !== '') out.push({ kind: 'text', text, ...(attributes ? { attributes: { ...attributes } } : {}) });
  });
  return out;
}

function clipboardControlText(
  reader: ClipboardTableReader,
  kind: 'sdt' | 'blockSdt',
  payload: Record<string, unknown>,
  depth = 0,
): boolean {
  if (depth > MAX_FRAGMENT_DEPTH) return false;
  const textLike = (itemKind: unknown, value: unknown): boolean => {
    switch (itemKind) {
      case 'tab': case 'break': case 'pilcrow':
        return true;
      case 'sdt': case 'blockSdt': {
        const nested = record(value);
        return nested !== null && clipboardControlText(reader, itemKind, nested, depth + 1);
      }
      default:
        return false;
    }
  };
  if (kind === 'blockSdt') {
    return typeof payload.story === 'string' && reader.storySegments(payload.story).every(segment =>
      segment.kind === 'text' || segment.kind === 'pilcrow' || textLike(segment.embedKind, segment.payload));
  }
  return Array.isArray(payload.content) && payload.content.every(item => plainObject(item) &&
    (item.kind === 'text' ? typeof item.text === 'string' : textLike(item.kind, item.payload)));
}

function clipboardTextEmbed(reader: ClipboardTableReader, kind: string, payload: Record<string, unknown>, text: string): boolean {
  switch (kind) {
    case 'bookmark': case 'bookmarkStart': case 'bookmarkEnd':
    case 'commentAnchor': case 'commentReference': case 'commentRangeStart': case 'commentRangeEnd':
    case 'rangeMarker': case 'rangeStart': case 'rangeEnd':
    case 'permStart': case 'permEnd':
    case 'moveFromRangeStart': case 'moveFromRangeEnd': case 'moveToRangeStart': case 'moveToRangeEnd':
      return true;
    case 'field':
      return payload.modelKind === 'commentReference' || text !== '';
    case 'sdt': case 'blockSdt':
      return text !== '' && clipboardControlText(reader, kind, payload);
    case 'math':
      return text !== '';
    case 'shape': case 'chart': case 'opaqueDrawing': case 'horizontalRule':
    case 'noteRef': case 'pageBreak': case 'columnBreak': case 'opaque':
      return false;
    default:
      return false;
  }
}

function shadingOf(tcPr: Record<string, unknown>): string | undefined {
  const color = typeof tcPr.backgroundColor === 'string' ? tcPr.backgroundColor.replace(/^#+/, '') : null;
  return color && /^[0-9a-f]{6}$/iu.test(color) ? color : undefined;
}

function fragmentFromPayload(
  reader: ClipboardTableReader,
  payload: { tblPr?: unknown; grid?: unknown; rows?: unknown },
  rect?: ClipboardTableRect,
  onUnsupported?: () => void,
): DocxTableFragment {
  const rows = rawRows(payload);
  const tablePayload: TablePayload = { rows: rows as TablePayload['rows'] };
  const { anchors, columns } = tableAnchors(tablePayload);
  const rowCount = rows.length;
  const full: ClipboardTableRect = {
    top: 0,
    bottom: rowCount - 1,
    left: 0,
    right: Math.max(columns - 1, 0),
  };
  const clamp = (value: number, max: number): number => Math.min(Math.max(value, 0), Math.max(max, 0));
  let next = rect
    ? {
        top: clamp(Math.min(rect.top, rect.bottom), rowCount - 1),
        bottom: clamp(Math.max(rect.top, rect.bottom), rowCount - 1),
        left: clamp(Math.min(rect.left, rect.right), columns - 1),
        right: clamp(Math.max(rect.left, rect.right), columns - 1),
      }
    : { ...full };
  for (;;) {
    const before = `${next.top}:${next.bottom}:${next.left}:${next.right}`;
    for (const anchor of anchors) {
      if (
        anchor.row <= next.bottom &&
        anchor.row + anchor.rowspan - 1 >= next.top &&
        anchor.column <= next.right &&
        anchor.column + anchor.colspan - 1 >= next.left
      ) {
        next = {
          top: Math.min(next.top, anchor.row),
          bottom: Math.max(next.bottom, anchor.row + anchor.rowspan - 1),
          left: Math.min(next.left, anchor.column),
          right: Math.max(next.right, anchor.column + anchor.colspan - 1),
        };
      }
    }
    if (`${next.top}:${next.bottom}:${next.left}:${next.right}` === before) break;
  }
  const { top, bottom, left, right } = next;
  const byRow = new Map<number, typeof anchors>();
  for (const anchor of anchors) {
    if (
      anchor.row >= top &&
      anchor.row + anchor.rowspan - 1 <= bottom &&
      anchor.column >= left &&
      anchor.column + anchor.colspan - 1 <= right
    ) {
      const list = byRow.get(anchor.row) ?? [];
      list.push(anchor);
      byRow.set(anchor.row, list);
    }
  }
  if (byRow.size === 0) throw new Error('Table selection holds no cells');
  const fragmentRows: DocxTableFragmentRow[] = [];
  for (let rowIndex = top; rowIndex <= bottom; rowIndex += 1) {
    const cells = (byRow.get(rowIndex) ?? []).sort((a, b) => a.column - b.column);
    fragmentRows.push({
      trPr: copiedMap(rows[rowIndex]?.trPr) ?? {},
      cells: cells.map((anchor) => {
        const raw = (rows[anchor.row]?.cells as RawCell[] | undefined)?.find(
          (cell) => cell?.story === anchor.story,
        );
        const tcPr = { ...(record(raw?.tcPr) ?? {}) };
        const shading = shadingOf(tcPr);
        if (shading !== undefined) delete tcPr.backgroundColor;
        return {
          tcPr: {
            ...tcPr,
            rowspan: anchor.rowspan,
            colspan: anchor.colspan,
          },
          content: cellContent(reader, anchor.story, onUnsupported),
          ...(shading !== undefined ? { shading } : {}),
        };
      }),
    });
  }
  const grid =
    Array.isArray(payload.grid) &&
    payload.grid.every((width) => typeof width === 'number' && Number.isFinite(width) && width > 0)
      ? (payload.grid as number[]).slice(left, right + 1)
      : [];
  const tblPr = { ...(record(payload.tblPr) ?? {}) };
  if (
    (left > 0 || right < columns - 1) &&
    tblPr.widthType === 'dxa' &&
    typeof tblPr.width === 'number' &&
    grid.length > 0
  ) {
    tblPr.width = grid.reduce((sum, width) => sum + width, 0);
  }
  return { tblPr, grid, rows: fragmentRows };
}

/**
 * Reads one rectangular table as a fragment. `rect` is inclusive; it expands
 * to fully cover every merged cell it touches.
 * @internal
 */
export function readTableFragment(
  reader: ClipboardTableReader,
  table: { story: string; tableIndex: number },
  rect?: ClipboardTableRect,
): DocxTableFragment {
  const raw = reader.tablePayload(table.story, table.tableIndex);
  if (!raw) throw new Error(`Table ${table.tableIndex} was not found in story ${table.story}`);
  const payload = clipboardTablePayload(raw);
  if (!payload) throw new Error(`Table ${table.tableIndex} in story ${table.story} holds no rows`);
  return fragmentFromPayload(reader, payload, rect);
}

/**
 * Reads the clipboard slice for a copy: always plain text, plus a fragment
 * when the selection holds tables or pictures.
 * @internal
 */
export function readClipboardSlice(
  reader: ClipboardTableReader,
  target: ClipboardSliceTarget,
): DocxClipboardSlice {
  if ('cells' in target) {
    const { anchor, head } = target.cells;
    const payload = clipboardTablePayload(reader.tablePayload(anchor.story, anchor.tableIndex));
    const range = {
      top: Math.min(anchor.row, head.row),
      bottom: Math.max(anchor.row, head.row),
      left: Math.min(anchor.column, head.column),
      right: Math.max(anchor.column, head.column),
    };
    const text = payload ? tablePlainText(reader, payload, range).join('\n') : '';
    return {
      text,
      fragment: payload
        ? {
            paragraphs: [
              {
                tables: [fragmentFromPayload(reader, payload, range)],
                text: '',
              },
            ],
          }
        : null,
    };
  }
  const from = Math.min(target.from, target.to);
  const to = Math.max(target.from, target.to);
  const text = storyPlainText(reader, target.story, from, to);
  const paragraphs: DocxClipboardParagraph[] = [];
  let current: { tables: DocxTableFragment[]; text: string; inlines: DocxClipboardInline[] } = { tables: [], text: '', inlines: [] };
  let sawRichContent = false;
  let sawBoundary = false;
  let sawUnrepresentableEmbed = false;
  const push = (): void => {
    paragraphs.push({
      ...(current.tables.length > 0 ? { tables: current.tables } : {}),
      text: current.text,
      ...(current.inlines.some(inline => 'image' in inline) ? { inlines: current.inlines } : {}),
    });
    current = { tables: [], text: '', inlines: [] };
  };
  const appendText = (value: string): void => {
    if (value === '') return;
    current.text += value;
    current.inlines.push({ text: value });
  };
  const appendPlainText = (value: string): void => {
    const parts = value.split('\n');
    appendText(parts[0] ?? '');
    for (const part of parts.slice(1)) {
      sawBoundary = true;
      push();
      appendText(part);
    }
  };
  let offset = 0;
  for (const segment of reader.storySegments(target.story)) {
    const start = offset;
    offset += segment.kind === 'text' ? segment.text.length : 1;
    if (offset <= from || start >= to) continue;
    if (segment.kind === 'text') {
      appendText(segment.text.slice(Math.max(from, start) - start, Math.min(to, offset) - start));
    } else if (segment.kind === 'pilcrow') {
      sawBoundary = true;
      push();
    } else if (segment.embedKind === 'table') {
      if (current.inlines.length > 0) push();
      current.tables.push(fragmentFromPayload(reader, segment.payload, undefined, () => { sawUnrepresentableEmbed = true; }));
      sawRichContent = true;
    } else if (segment.embedKind === 'image') {
      const image = independentImage(segment.payload);
      const source = image.src;
      image.src = typeof source === 'string' && source.startsWith('media:') ? reader.mediaDataUrl?.(source) : source;
      if (!validImage(image)) throw new Error('Picture clipboard media is unavailable or invalid');
      current.inlines.push({ image });
      sawRichContent = true;
    } else if (segment.embedKind === 'tab') {
      appendText('\t');
    } else if (segment.embedKind === 'break') {
      sawBoundary = true;
      push();
    } else {
      const plain = embedPlainText(reader, segment.embedKind, segment.payload);
      if (!clipboardTextEmbed(reader, segment.embedKind, segment.payload, plain)) sawUnrepresentableEmbed = true;
      appendPlainText(plain);
    }
  }
  if (sawBoundary || current.tables.length > 0 || current.inlines.length > 0) push();
  const fragment = sawRichContent ? { paragraphs } : null;
  if (fragment && clipboardFragmentHasPictures(fragment)) {
    if (sawUnrepresentableEmbed) throw new Error('Picture selection contains unsupported clipboard objects');
    if (!validClipboardFragment(fragment)) throw new Error('Clipboard fragment exceeds the supported size or shape');
  }
  return { text, fragment };
}

/**
 * Pastes a fragment over a story range, returning the caret. Deletes a
 * non-collapsed range first, then replays each paragraph: splits for
 * paragraphs after the first, inserts tables, then text.
 * @internal
 */
export function applyClipboardFragment(
  engine: ClipboardPasteEngine,
  range: YrsStoryRange,
  fragment: DocxClipboardFragment,
  suggesting?: YrsAuthor,
): YrsLoc {
  for (const paragraph of fragment.paragraphs) {
    for (const table of paragraph.tables ?? []) engine.validateTableFragment(table);
    for (const inline of paragraph.inlines ?? []) {
      if ('image' in inline && !validImage(inline.image)) throw new Error('Invalid clipboard picture');
    }
  }
  let caret: YrsLoc = {
    story: range.story,
    paraId: range.start.paraId,
    offset: range.start.offset,
  };
  if (range.start.paraId !== range.end.paraId || range.start.offset !== range.end.offset) {
    const removed = engine.deleteRange(range, suggesting);
    if (!removed.range) throw new Error('Delete receipt lacks range');
    caret = {
      story: removed.range.story,
      paraId: removed.range.start.paraId,
      offset: removed.range.start.offset,
    };
  }
  fragment.paragraphs.forEach((paragraph, index) => {
    if (index > 0) {
      const split = engine.splitParagraph(caret, suggesting);
      caret = { story: caret.story, paraId: split.secondParaId, offset: 0 };
    }
    for (const table of paragraph.tables ?? []) {
      caret = engine.insertTableFragment(caret, table, suggesting).after;
    }
    for (const inline of paragraph.inlines ?? (paragraph.text === '' ? [] : [{ text: paragraph.text }])) {
      const inserted = 'image' in inline
        ? engine.insertImage(caret, independentImage(inline.image), suggesting)
        : engine.insertText(caret, inline.text, suggesting);
      if (!inserted.range) throw new Error('Clipboard insertion receipt lacks range');
      caret = { story: inserted.range.story, ...inserted.range.end };
    }
  });
  return caret;
}
