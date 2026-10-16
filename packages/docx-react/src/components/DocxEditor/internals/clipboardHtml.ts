import {
  DOCX_CLIPBOARD_FRAGMENT_VERSION,
  validClipboardFragment,
  type DocxClipboardFragment,
  type DocxClipboardParagraph,
  type DocxFragmentSegment,
  type DocxTableFragment,
  type DocxTableFragmentCell,
} from '@betteroffice/docx/yrs';
import type { YrsInlineFormatDelta } from '@betteroffice/docx/yrs';

const NBSP = '\u00a0';
const LINE_SEP = '\u2028';

const DEFAULT_BORDERS = () => ({
  top: { style: 'single', size: 4, color: { rgb: '000000' } },
  bottom: { style: 'single', size: 4, color: { rgb: '000000' } },
  left: { style: 'single', size: 4, color: { rgb: '000000' } },
  right: { style: 'single', size: 4, color: { rgb: '000000' } },
});

const BLOCK_TAGS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER']);
const SKIP_TAGS = new Set(['STYLE', 'SCRIPT', 'TITLE', 'META', 'HEAD']);

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function objectValue(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function stringValue(value: unknown): string | null {
  return typeof value === 'string' && value.length > 0 ? value : null;
}

function finiteNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function spanStyleForText(attributes: Record<string, unknown> | undefined, format: YrsInlineFormatDelta | undefined): string {
  const parts: string[] = [];
  let bold = false;
  let italic = false;
  let underline = false;
  let strike = false;
  let color: string | null = null;
  let fontSizePt: number | null = null;
  let fontFamily: string | null = null;
  let highlight: string | null = null;
  const readRaw = (attrs: Record<string, unknown>) => {
    if (attrs.bold === true) bold = true;
    if (attrs.italic === true) italic = true;
    if (attrs.underline === true || objectValue(attrs.underline)) underline = true;
    if (attrs.strike === true || objectValue(attrs.strike)) strike = true;
    const rgb = objectValue(attrs.textColor) ? stringValue(objectValue(attrs.textColor)?.rgb) : null;
    if (rgb) color = rgb;
    const half = objectValue(attrs.fontSize) ? finiteNumber(objectValue(attrs.fontSize)?.size) : null;
    if (half !== null && half !== undefined && half > 0) fontSizePt = half / 2;
    const ascii = objectValue(attrs.fontFamily) ? stringValue(objectValue(attrs.fontFamily)?.ascii) : null;
    if (ascii) fontFamily = ascii;
    const highlightMap = objectValue(attrs.highlight);
    const highlightColor = typeof attrs.highlight === 'string' ? attrs.highlight : highlightMap ? stringValue(highlightMap.color) : null;
    if (highlightColor) highlight = highlightColor;
  };
  if (attributes) readRaw(attributes);
  if (format) {
    if (format.bold === true) bold = true;
    if (format.italic === true) italic = true;
    if (format.underline === true || objectValue(format.underline as unknown)) underline = true;
    if (format.strike === true || objectValue(format.strike as unknown)) strike = true;
    const formatRgb = objectValue(format.color as unknown) ? stringValue(objectValue(format.color as unknown)?.rgb) : null;
    if (formatRgb) color = formatRgb;
    if (typeof format.fontSize === 'number' && Number.isFinite(format.fontSize) && format.fontSize > 0) {
      fontSizePt = format.fontSize;
    }
    if (format.fontFamily && typeof format.fontFamily.ascii === 'string' && format.fontFamily.ascii) {
      fontFamily = format.fontFamily.ascii;
    }
    if (typeof format.highlight === 'string' && format.highlight) highlight = format.highlight;
  }
  if (bold) parts.push('font-weight:bold');
  if (italic) parts.push('font-style:italic');
  const decorations: string[] = [];
  if (underline) decorations.push('underline');
  if (strike) decorations.push('line-through');
  if (decorations.length > 0) parts.push(`text-decoration:${decorations.join(' ')}`);
  if (color) parts.push(`color:#${color.replace(/^#/, '')}`);
  if (fontSizePt !== null) parts.push(`font-size:${fontSizePt}pt`);
  if (fontFamily) parts.push(`font-family:${fontFamily}`);
  if (highlight) {
    const value = /^[0-9a-f]{6}$/i.test(highlight) ? `#${highlight}` : highlight;
    parts.push(`background-color:${value}`);
  }
  return parts.join(';');
}

function segmentStyle(segment: DocxFragmentSegment): string {
  if (segment.kind !== 'text') return '';
  return spanStyleForText(segment.attributes, segment.format);
}

function serializeRuns(segments: DocxFragmentSegment[]): string {
  let html = '';
  for (const segment of segments) {
    if (segment.kind === 'text') {
      const style = segmentStyle(segment);
      const text = escapeHtml(segment.text);
      html += style ? `<span style="${escapeHtml(style)}">${text}</span>` : text;
    } else if (segment.kind === 'tab') {
      html += `<span style="${escapeHtml('white-space:pre')}">\t</span>`;
    } else if (segment.kind === 'break') {
      html += '<br>';
    } else if (segment.kind === 'table') {
      html += serializeTable(segment.table);
    }
  }
  return html;
}

function cellParagraphs(content: DocxFragmentSegment[]): DocxFragmentSegment[][] {
  const paragraphs: DocxFragmentSegment[][] = [];
  let current: DocxFragmentSegment[] = [];
  for (const segment of content) {
    if (segment.kind === 'pilcrow') {
      paragraphs.push(current);
      current = [];
    } else if (segment.kind === 'table') {
      if (current.length > 0) {
        paragraphs.push(current);
        current = [];
      }
      paragraphs.push([segment]);
    } else {
      current.push(segment);
    }
  }
  if (current.length > 0) paragraphs.push(current);
  return paragraphs;
}

function tcPrSpan(tcPr: Record<string, unknown>): { colspan: number; rowspan: number } {
  const colspanValue = tcPr.colspan;
  const rowspanValue = tcPr.rowspan;
  const colspan = typeof colspanValue === 'number' && Number.isInteger(colspanValue) && colspanValue >= 1 ? colspanValue : 1;
  const rowspan = typeof rowspanValue === 'number' && Number.isInteger(rowspanValue) && rowspanValue >= 1 ? rowspanValue : 1;
  return { colspan, rowspan };
}

function cellShading(cell: DocxTableFragmentCell): string | null {
  if (typeof cell.shading === 'string' && /^[0-9a-f]{6}$/i.test(cell.shading)) return cell.shading;
  const tcPrShading = stringValue((cell.tcPr as Record<string, unknown>).shading);
  if (tcPrShading && /^[0-9a-f]{6}$/i.test(tcPrShading.replace(/^#/, ''))) {
    return tcPrShading.replace(/^#/, '');
  }
  return null;
}

function serializeCell(cell: DocxTableFragmentCell): string {
  const tcPr = cell.tcPr ?? {};
  const { colspan, rowspan } = tcPrSpan(tcPr);
  const styles: string[] = [];
  const borders = objectValue(tcPr.borders);
  if (borders && Object.keys(borders).length > 0) styles.push('border:1px solid #000');
  const shading = cellShading(cell);
  if (shading) styles.push(`background-color:#${shading}`);
  const attrs: string[] = [];
  if (colspan > 1) attrs.push(` colspan="${colspan}"`);
  if (rowspan > 1) attrs.push(` rowspan="${rowspan}"`);
  if (styles.length > 0) attrs.push(` style="${escapeHtml(styles.join(';'))}"`);
  let html = `<td${attrs.join('')}>`;
  for (const paragraph of cellParagraphs(cell.content)) {
    if (paragraph.length === 1 && paragraph[0]!.kind === 'table') {
      html += serializeTable((paragraph[0] as { kind: 'table'; table: DocxTableFragment }).table);
    } else {
      html += `<p style="${escapeHtml('margin:0')}">${serializeRuns(paragraph)}</p>`;
    }
  }
  html += '</td>';
  return html;
}

function serializeTable(table: DocxTableFragment): string {
  let html = `<table style="${escapeHtml('border-collapse:collapse')}">`;
  if (table.grid.length > 0) {
    html += '<colgroup>';
    for (const twips of table.grid) {
      html += `<col style="${escapeHtml(`width:${twips / 20}pt`)}">`;
    }
    html += '</colgroup>';
  }
  for (const row of table.rows) {
    html += '<tr>';
    for (const cell of row.cells) html += serializeCell(cell);
    html += '</tr>';
  }
  html += '</table>';
  return html;
}

/** HTML clipboard form of a fragment for Word, Google Docs and internal paste. */
export function clipboardFragmentHtml(fragment: DocxClipboardFragment): string {
  const payload = escapeHtml(JSON.stringify({ v: DOCX_CLIPBOARD_FRAGMENT_VERSION, fragment }));
  let html = `<div data-betteroffice-docx="${payload}">`;
  for (const paragraph of fragment.paragraphs) {
    for (const table of paragraph.tables ?? []) html += serializeTable(table);
    if (paragraph.text !== '' || (paragraph.tables ?? []).length === 0) {
      html += `<p>${escapeHtml(paragraph.text)}</p>`;
    }
  }
  html += '</div>';
  return html;
}

function parseStyleMap(style: string | null): Map<string, string> {
  const map = new Map<string, string>();
  if (!style) return map;
  for (const part of style.split(';')) {
    const index = part.indexOf(':');
    if (index < 0) continue;
    const key = part.slice(0, index).trim().toLowerCase();
    const value = part.slice(index + 1).trim();
    if (key) map.set(key, value);
  }
  return map;
}

function styleOf(element: Element): Map<string, string> {
  return parseStyleMap(element.getAttribute('style'));
}

function widthAttrTwips(element: Element): number | null {
  const style = styleOf(element).get('width') ?? element.getAttribute('width');
  if (!style) return null;
  const match = /([\d.]+)\s*(pt|px)/i.exec(style);
  if (!match) return null;
  const value = Number.parseFloat(match[1]!);
  if (!Number.isFinite(value) || value <= 0) return null;
  return Math.round(value * (match[2]!.toLowerCase() === 'px' ? 15 : 20));
}

function colorToHex(value: string | null | undefined): string | null {
  if (!value) return null;
  const text = value.trim();
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(text);
  if (hex) {
    const digits = hex[1]!;
    const full = digits.length === 3 ? digits.split('').map((c) => c + c).join('') : digits;
    return full.toUpperCase();
  }
  const rgb = /^rgba?\(\s*([\d.]+%?)\s*,\s*([\d.]+%?)\s*,\s*([\d.]+%?)/i.exec(text);
  if (rgb) {
    const parts = [rgb[1]!, rgb[2]!, rgb[3]!].map((part) => {
      const number = part.includes('%') ? (Number.parseFloat(part) / 100) * 255 : Number.parseFloat(part);
      return Math.max(0, Math.min(255, Math.round(number)));
    });
    return parts.map((part) => part.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  return null;
}

function fontSizeToPt(value: string | null | undefined): number | null {
  if (!value) return null;
  const match = /([\d.]+)\s*(pt|px)/i.exec(value.trim());
  if (!match) return null;
  const number = Number.parseFloat(match[1]!);
  if (!Number.isFinite(number) || number <= 0) return null;
  return match[2]!.toLowerCase() === 'px' ? number * 0.75 : number;
}

function fontFamilyFirst(value: string | null | undefined): string | null {
  if (!value) return null;
  const first = value.split(',')[0]!.trim().replace(/^["']+|["']+$/g, '').trim();
  return first ? first : null;
}

function fontWeightBold(value: string | null | undefined): boolean | undefined {
  if (!value) return undefined;
  const text = value.trim().toLowerCase();
  if (text === 'bold') return true;
  if (text === 'normal') return false;
  const number = Number.parseFloat(text);
  if (Number.isFinite(number)) {
    if (number >= 600) return true;
    if (number <= 500) return false;
  }
  return undefined;
}

interface InlineFormatState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: string | null;
  fontSize: number | null;
  fontFamily: string | null;
  highlight: string | null;
}

function blankInlineFormat(): InlineFormatState {
  return { bold: false, italic: false, underline: false, strike: false, color: null, fontSize: null, fontFamily: null, highlight: null };
}

function applyInlineElement(state: InlineFormatState, element: Element, isCell: boolean): void {
  const tag = element.tagName.toUpperCase();
  let boldSpec: boolean | undefined;
  let italicSpec: boolean | undefined;
  let underlineSpec: boolean | undefined;
  let strikeSpec: boolean | undefined;
  if (tag === 'B' || tag === 'STRONG') boldSpec = true;
  if (tag === 'I' || tag === 'EM') italicSpec = true;
  if (tag === 'U' || tag === 'INS') underlineSpec = true;
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') strikeSpec = true;
  const style = styleOf(element);
  const weight = fontWeightBold(style.get('font-weight'));
  if (weight !== undefined) boldSpec = weight;
  const fontStyle = (style.get('font-style') ?? '').trim().toLowerCase();
  if (fontStyle === 'italic' || fontStyle === 'oblique') italicSpec = true;
  else if (fontStyle === 'normal') italicSpec = false;
  const decoration = (style.get('text-decoration-line') ?? style.get('text-decoration') ?? '').toLowerCase();
  if (decoration) {
    if (/\bnone\b/.test(decoration)) {
      underlineSpec = false;
      strikeSpec = false;
    } else {
      if (decoration.includes('underline')) underlineSpec = true;
      if (decoration.includes('line-through')) strikeSpec = true;
    }
  }
  if (boldSpec !== undefined) state.bold = boldSpec;
  if (italicSpec !== undefined) state.italic = italicSpec;
  if (underlineSpec !== undefined) state.underline = underlineSpec;
  if (strikeSpec !== undefined) state.strike = strikeSpec;
  const color = colorToHex(style.get('color'));
  if (color) state.color = color;
  const size = fontSizeToPt(style.get('font-size'));
  if (size !== null) state.fontSize = size;
  const family = fontFamilyFirst(style.get('font-family'));
  if (family) state.fontFamily = family;
  if (!isCell) {
    const highlight = colorToHex(style.get('background-color') ?? style.get('background'));
    if (highlight) state.highlight = highlight;
  }
}

function formatDelta(state: InlineFormatState): YrsInlineFormatDelta | undefined {
  const delta: YrsInlineFormatDelta = {};
  if (state.bold) delta.bold = true;
  if (state.italic) delta.italic = true;
  if (state.underline) delta.underline = true;
  if (state.strike) delta.strike = true;
  if (state.color) delta.color = { rgb: state.color };
  if (state.fontSize !== null) delta.fontSize = state.fontSize;
  if (state.fontFamily) delta.fontFamily = { ascii: state.fontFamily };
  if (state.highlight) delta.highlight = state.highlight;
  return Object.keys(delta).length > 0 ? delta : undefined;
}

function normalizeTextPart(part: string): string {
  return part.replace(/[ \n\r\f\v\u00a0\u2028\u2029]+/g, ' ');
}

function normalizeParagraphText(raw: string): string {
  const withSpaces = raw.split(NBSP).join(' ');
  const parts = withSpaces.split('\t').map(normalizeTextPart);
  const joined = parts.join('\t');
  const trimmed = joined.replace(/^ +| +$/g, '');
  if (trimmed === '') return '';
  return trimmed;
}

function textNodeText(node: Text): string {
  return node.data.split(NBSP).join(' ');
}

function blockLines(block: Element): string[] {
  const lines: string[] = [];
  let current = '';
  const visit = (node: Node) => {
    if (node.nodeType === 8) return;
    if (node.nodeType === 3) {
      current += textNodeText(node as Text);
      return;
    }
    if (node.nodeType !== 1) return;
    const element = node as Element;
    if (element.tagName.toUpperCase() === 'BR') {
      lines.push(current);
      current = '';
      return;
    }
    if (element.tagName.toUpperCase() === 'TABLE') return;
    for (const child of Array.from(element.childNodes)) visit(child);
  };
  for (const child of Array.from(block.childNodes)) visit(child);
  lines.push(current);
  return lines.map(normalizeParagraphText);
}

function containsTable(element: Element): boolean {
  if (element.tagName.toUpperCase() === 'TABLE') return true;
  for (const child of Array.from(element.children)) {
    if (containsTable(child)) return true;
  }
  return false;
}

function childTag(parent: Element, tag: string): Element[] {
  return Array.from(parent.children).filter((child) => child.tagName.toUpperCase() === tag);
}

function tableRows(table: HTMLTableElement): HTMLTableRowElement[] {
  const rows: HTMLTableRowElement[] = [];
  for (const head of childTag(table, 'THEAD')) {
    for (const row of childTag(head, 'TR')) rows.push(row as HTMLTableRowElement);
  }
  for (const child of Array.from(table.children)) {
    const tag = child.tagName.toUpperCase();
    if (tag === 'TBODY') {
      for (const row of childTag(child, 'TR')) rows.push(row as HTMLTableRowElement);
    } else if (tag === 'TR') {
      rows.push(child as HTMLTableRowElement);
    }
  }
  for (const foot of childTag(table, 'TFOOT')) {
    for (const row of childTag(foot, 'TR')) rows.push(row as HTMLTableRowElement);
  }
  return rows;
}

function rowCells(row: HTMLTableRowElement): Element[] {
  return Array.from(row.children).filter((child) => {
    const tag = child.tagName.toUpperCase();
    return tag === 'TD' || tag === 'TH';
  });
}

function directCols(table: HTMLTableElement): Element[] {
  const cols: Element[] = [];
  for (const child of Array.from(table.children)) {
    if (child.tagName.toUpperCase() === 'COLGROUP') {
      for (const col of Array.from(child.children)) {
        if (col.tagName.toUpperCase() === 'COL') cols.push(col);
      }
    } else if (child.tagName.toUpperCase() === 'COL') {
      cols.push(child);
    }
  }
  return cols;
}

function intAttr(element: Element, name: string): number {
  const raw = element.getAttribute(name);
  if (!raw) return 1;
  const value = Number.parseInt(raw, 10);
  if (!Number.isInteger(value) || value < 1) return 1;
  return Math.min(value, 1000);
}

function appendTextWithTabs(
  target: DocxFragmentSegment[],
  text: string,
  format: YrsInlineFormatDelta | undefined
): void {
  const parts = text.split('\t');
  parts.forEach((part, index) => {
    if (index > 0) target.push({ kind: 'tab' });
    if (part !== '') {
      target.push(format ? { kind: 'text', text: part, format } : { kind: 'text', text: part });
    }
  });
}

function parseCellContent(cell: Element): DocxFragmentSegment[] {
  const segments: DocxFragmentSegment[] = [];
  let current: DocxFragmentSegment[] = [];
  const flushCurrent = (force: boolean) => {
    if (current.length > 0) {
      const onlySpaces = current.every((segment) => segment.kind === 'text' && /^[ ]+$/.test(segment.text));
      if (onlySpaces) current = [];
    }
    if (current.length > 0 || force) {
      segments.push(...current);
      segments.push({ kind: 'pilcrow' });
      current = [];
    }
  };
  const pushInlineText = (text: string, ancestors: Element[]) => {
    const state = blankInlineFormat();
    ancestors.forEach((ancestor, index) => {
      if (index === 0) return;
      applyInlineElement(state, ancestor, ancestor === cell);
    });
    const format = formatDelta(state);
    appendTextWithTabs(current, text, format);
  };
  const walkInline = (node: Node, ancestors: Element[]) => {
    if (node.nodeType === 8) return;
    if (node.nodeType === 3) {
      const raw = textNodeText(node as Text);
      const parts = raw.split('\t').map(normalizeTextPart);
      const restored = parts.join('\t');
      if (restored === '') return;
      if (/^[ ]+$/.test(restored) && !restored.includes('\t') && current.length === 0) return;
      pushInlineText(restored, ancestors);
      return;
    }
    if (node.nodeType !== 1) return;
    const element = node as Element;
    const tag = element.tagName.toUpperCase();
    if (tag === 'BR') {
      flushCurrent(true);
      return;
    }
    if (tag === 'TABLE') {
      flushCurrent(false);
      segments.push({ kind: 'table', table: parseTableElement(element as HTMLTableElement) });
      return;
    }
    const next = [...ancestors, element];
    for (const child of Array.from(element.childNodes)) walkInline(child, next);
  };
  const walkBlockChildren = (nodes: Node[], ancestors: Element[]) => {
    for (const node of nodes) {
      if (node.nodeType === 8) continue;
      if (node.nodeType === 1 && (node as Element).tagName.toUpperCase() === 'TABLE') {
        flushCurrent(false);
        segments.push({ kind: 'table', table: parseTableElement(node as unknown as HTMLTableElement) });
        continue;
      }
      if (node.nodeType === 1 && BLOCK_TAGS.has((node as Element).tagName.toUpperCase())) {
        const element = node as Element;
        if (containsTable(element)) {
          walkBlockChildren(Array.from(element.childNodes), [...ancestors, element]);
        } else {
          flushCurrent(false);
          const before = segments.length;
          for (const child of Array.from(element.childNodes)) walkInline(child, [...ancestors, element]);
          flushCurrent(segments.length === before);
        }
        continue;
      }
      walkInline(node, ancestors);
    }
  };
  walkBlockChildren(Array.from(cell.childNodes), [cell]);
  if (current.length > 0) flushCurrent(false);
  if (segments.length === 0) return [{ kind: 'pilcrow' }];
  const last = segments[segments.length - 1]!;
  if (last.kind === 'table') segments.push({ kind: 'pilcrow' });
  return segments;
}

function parseTableElement(table: HTMLTableElement): DocxTableFragment {
  const rows = tableRows(table);
  type ParsedCell = { colspan: number; rowspan: number; shading: string | null; content: DocxFragmentSegment[]; width: number | null };
  const parsedRows: ParsedCell[][] = rows.map((row) => {
    return rowCells(row).map((cell) => {
      const colspan = intAttr(cell, 'colspan');
      const rowspan = intAttr(cell, 'rowspan');
      const background = styleOf(cell).get('background-color') ?? cell.getAttribute('bgcolor') ?? null;
      const shading = colorToHex(background);
      return { colspan, rowspan, shading, content: parseCellContent(cell), width: widthAttrTwips(cell) };
    });
  });
  const rowCount = parsedRows.length;
  const occupied: boolean[][] = parsedRows.map(() => []);
  type PlacedCell = ParsedCell & { column: number };
  const placedRows: PlacedCell[][] = [];
  let columns = 0;
  parsedRows.forEach((cells, rowIndex) => {
    const placed: PlacedCell[] = [];
    let column = 0;
    for (const cell of cells) {
      while (occupied[rowIndex]![column]) column += 1;
      const rowspan = Math.min(cell.rowspan, rowCount - rowIndex);
      const colspan = cell.colspan;
      for (let r = rowIndex; r < rowIndex + rowspan; r += 1) {
        for (let c = column; c < column + colspan; c += 1) occupied[r]![c] = true;
      }
      placed.push({ ...cell, rowspan, colspan, column });
      column += colspan;
      columns = Math.max(columns, column);
    }
    placedRows.push(placed);
  });
  columns = Math.max(columns, 1);
  const filler = (): ParsedCell => ({
    colspan: 1, rowspan: 1, shading: null, content: [{ kind: 'pilcrow' }], width: null,
  });
  const filledRows: ParsedCell[][] = placedRows.map((placed, rowIndex) => {
    const byColumn = new Map(placed.map((cell) => [cell.column, cell as ParsedCell]));
    const cells: ParsedCell[] = [];
    for (let column = 0; column < columns; column += 1) {
      const cell = byColumn.get(column);
      if (cell) cells.push(cell);
      else if (!occupied[rowIndex]![column]) cells.push(filler());
    }
    return cells;
  });
  let grid: number[] = [];
  const cols = directCols(table);
  const colWidths = cols.map((col) => widthAttrTwips(col));
  if (cols.length === columns && colWidths.every((width): width is number => width !== null)) {
    grid = colWidths;
  } else if (filledRows.length > 0) {
    const first = filledRows[0]!;
    if (first.length === columns && first.every((cell) => cell.colspan === 1 && cell.width !== null)) {
      grid = first.map((cell) => cell.width as number);
    }
  }
  const width = grid.length > 0 ? grid.reduce((sum, value) => sum + value, 0) : 9360;
  return {
    tblPr: { width, widthType: 'dxa' },
    grid,
    rows: filledRows.map((cells) => ({
      trPr: {},
      cells: cells.map((cell) => {
        const tcPr: Record<string, unknown> = { borders: DEFAULT_BORDERS() };
        if (cell.colspan > 1) tcPr.colspan = cell.colspan;
        if (cell.rowspan > 1) tcPr.rowspan = cell.rowspan;
        const result: DocxTableFragmentCell = { tcPr, content: cell.content };
        if (cell.shading) result.shading = cell.shading;
        return result;
      }),
    })),
  };
}

/** Parsed table fragment from HTML, or null when the HTML holds no table. */
export function parseClipboardHtml(html: string): DocxClipboardFragment | null {
  const parser = new DOMParser();
  const document_ = parser.parseFromString(html, 'text/html');
  const payloadElement = document_.querySelector('[data-betteroffice-docx]');
  if (payloadElement) {
    const raw = payloadElement.getAttribute('data-betteroffice-docx');
    if (raw) {
      try {
        const parsed = JSON.parse(raw) as { v?: unknown; fragment?: unknown };
        if (parsed.v === DOCX_CLIPBOARD_FRAGMENT_VERSION && validClipboardFragment(parsed.fragment)) {
          return parsed.fragment;
        }
      } catch {
      }
    }
  }
  const body = document_.body;
  if (!body) return null;
  const paragraphs: DocxClipboardParagraph[] = [];
  let activeTablePara: DocxClipboardParagraph | null = null;
  const pushTextLines = (lines: string[]) => {
    for (const line of lines) {
      if (activeTablePara && activeTablePara.text === '') {
        activeTablePara.text = line;
        activeTablePara = null;
      } else {
        paragraphs.push({ text: line });
      }
    }
  };
  const pushTable = (table: DocxTableFragment) => {
    if (activeTablePara) {
      activeTablePara.tables!.push(table);
    } else {
      const paragraph: DocxClipboardParagraph = { tables: [table], text: '' };
      paragraphs.push(paragraph);
      activeTablePara = paragraph;
    }
  };
  const collectInlineRaw = (node: Node, out: string[]) => {
    if (node.nodeType === 8) return;
    if (node.nodeType === 3) {
      out.push(textNodeText(node as Text));
      return;
    }
    if (node.nodeType !== 1) return;
    const element = node as Element;
    const tag = element.tagName.toUpperCase();
    if (tag === 'BR') {
      out.push(LINE_SEP);
      return;
    }
    if (tag === 'TABLE' || SKIP_TAGS.has(tag)) return;
    for (const child of Array.from(element.childNodes)) collectInlineRaw(child, out);
  };
  const walkTop = (nodes: Node[]) => {
    let pendingInline: string[] = [];
    const flushPendingInline = () => {
      if (pendingInline.length === 0) return;
      const raw = pendingInline.join('');
      pendingInline = [];
      const lines = raw.split(LINE_SEP).map(normalizeParagraphText).filter((line) => line !== '');
      pushTextLines(lines);
    };
    for (const node of nodes) {
      if (node.nodeType === 8) continue;
      if (node.nodeType === 3) {
        const raw = textNodeText(node as Text);
        if (raw.trim() === '' && !raw.includes('\t')) continue;
        pendingInline.push(raw);
        continue;
      }
      if (node.nodeType !== 1) continue;
      const element = node as Element;
      const tag = element.tagName.toUpperCase();
      if (SKIP_TAGS.has(tag)) continue;
      if (tag === 'BR') {
        pendingInline.push(LINE_SEP);
        continue;
      }
      if (tag === 'TABLE') {
        flushPendingInline();
        pushTable(parseTableElement(element as HTMLTableElement));
        continue;
      }
      if (BLOCK_TAGS.has(tag)) {
        if (containsTable(element)) {
          flushPendingInline();
          walkTop(Array.from(element.childNodes));
          continue;
        }
        flushPendingInline();
        const lines = blockLines(element);
        pushTextLines(lines.filter((line) => line !== ''));
        continue;
      }
      if (containsTable(element)) {
        flushPendingInline();
        walkTop(Array.from(element.childNodes));
        continue;
      }
      const parts: string[] = [];
      collectInlineRaw(element, parts);
      const raw = parts.join('');
      if (raw.trim() === '' && !raw.includes('\t') && !raw.includes(LINE_SEP)) continue;
      pendingInline.push(raw);
    }
    flushPendingInline();
  };
  walkTop(Array.from(body.childNodes));
  const cleaned = paragraphs.filter((paragraph) => (paragraph.tables ?? []).length > 0 || paragraph.text !== '');
  if (!cleaned.some((paragraph) => (paragraph.tables ?? []).length > 0)) return null;
  const fragment: DocxClipboardFragment = { paragraphs: cleaned };
  return validClipboardFragment(fragment) ? fragment : null;
}
