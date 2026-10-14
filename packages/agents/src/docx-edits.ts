import { tableAnchors, type DocxEditStep, type DocxParagraphTarget, type TableCellAnchor, type TablePayload, type YrsSession, type YrsStorySegment } from '@betteroffice/docx/yrs';
import { plainText } from './prototype';
import { DocumentToolError, type TableCellPosition } from './types';

export interface TableCell {
  story: string;
  tableIndex: number;
  rows: number;
  columns: number;
  anchors: TableCellAnchor[];
  anchor: TableCellAnchor;
  mergedRows: Set<number>;
}

export interface ParagraphStyle { id: string; name?: string; isDefault: boolean }

type Pilcrow = Extract<YrsStorySegment, { kind: 'pilcrow' }>;
const PROTECTED = new Set(['ins', 'del', 'hyperlink']);
const PARAGRAPH_PROPERTIES = new Set([
  'pStyle', 'numPr', 'numPrFromStyle', '_originalFormatting', 'defaultTextFormatting',
  'alignment', 'spaceBefore', 'spaceAfter', 'spaceBeforeLines', 'spaceAfterLines',
  'beforeAutospacing', 'afterAutospacing', 'lineSpacing', 'lineSpacingRule', 'spacingExplicit',
  'indentLeft', 'indentRight', 'indentFirstLine', 'hangingIndent', 'borders', 'shading', 'tabs',
  'pageBreakBefore', 'keepNext', 'keepLines', 'widowControl', 'contextualSpacing', 'snapToGrid',
  'autoSpaceDE', 'autoSpaceDN', 'outlineLevel', 'bidi',
]);

export function tableCells(story: string, segments: readonly YrsStorySegment[]): Map<string, TableCell> {
  const cells = new Map<string, TableCell>();
  let tableIndex = 0;
  for (const segment of segments) {
    if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
    const payload = segment.payload as unknown as TablePayload;
    if (Array.isArray(payload.rows)) {
      const { anchors, columns } = tableAnchors(payload);
      const mergedRows = new Set<number>();
      payload.rows.forEach((row, index) => {
        if (row.cells.some(cell => {
          const original = cell.tcPr?._originalFormatting as Record<string, unknown> | undefined;
          return cell.tcPr?.vMerge !== undefined || original?.vMerge !== undefined;
        })) mergedRows.add(index);
      });
      for (const anchor of anchors) {
        if (anchor.rowspan > 1) for (let row = anchor.row; row < anchor.row + anchor.rowspan; row++) mergedRows.add(row);
        if (anchor.story) cells.set(anchor.story, { story, tableIndex, rows: payload.rows.length, columns, anchors, anchor, mergedRows });
      }
    }
    tableIndex++;
  }
  return cells;
}

export function cellPosition(cell: TableCell): TableCellPosition {
  return { story: cell.story, index: cell.tableIndex + 1, row: cell.anchor.row + 1, column: cell.anchor.column + 1, rows: cell.rows, columns: cell.columns };
}

export function resolveStyle(styles: readonly ParagraphStyle[], wanted: string): string {
  const key = (value: string) => value.toLowerCase().replace(/[^\p{L}\p{N}]/gu, '');
  const target = key(wanted);
  const found = styles.find(style => style.id === wanted) ?? styles.find(style => key(style.id) === target || (style.name !== undefined && key(style.name) === target));
  if (found) return found.id;
  const shared = (value: string) => { let length = 0; while (length < value.length && value[length] === target[length]) length++; return length; };
  const nearby = styles.map(style => ({ id: style.id, score: Math.max(shared(key(style.id)), shared(key(style.name ?? ''))) }))
    .sort((a, b) => b.score - a.score).slice(0, 12).map(style => style.id);
  throw new DocumentToolError('UNKNOWN_STYLE', `Paragraph style "${wanted}" is not defined in this document. Use one of validStyles.`, { style: wanted, validStyles: nearby });
}

export function lineText(text: string) {
  plainText(text);
  if (/[\t\r\n\u000b\u000c￼]/.test(text)) throw new DocumentToolError('INVALID_ARGUMENT', 'Paragraph and cell text is one line of plain text: no tabs, line breaks or object placeholders.');
}

export function requireParagraph(draft: YrsSession, target: DocxParagraphTarget) {
  if (!draft.hasStory(target.story) || draft.paragraphIdCount(target.story, target.paraId) === 0) {
    throw new DocumentToolError('UNKNOWN_REF', 'An earlier edit in this batch removed this paragraph.');
  }
}

export function hostEdit(draft: YrsSession, step: DocxEditStep) {
  const result = draft.applyEdits({ expectVersion: draft.version(), source: 'agent', history: 'none', steps: [step] });
  if (!result.ok) throw new DocumentToolError(result.failure.code.toUpperCase().replaceAll('-', '_'), result.failure.message);
  return result.receipts[0];
}

export function deleteParagraph(draft: YrsSession, target: DocxParagraphTarget) {
  if (draft.paragraphSpans(target.story).length < 2) {
    throw new DocumentToolError('LAST_PARAGRAPH', 'A table cell or story keeps at least one paragraph. Use delete_table_row to remove a row.');
  }
  hostEdit(draft, { op: 'deleteParagraphs', story: target.story, firstParaId: target.paraId, lastParaId: target.paraId });
}

export function notInTable(): never {
  throw new DocumentToolError('NOT_IN_TABLE', 'This paragraph is not in a table cell. Use a ref whose outline entry has table.');
}

function locate(draft: YrsSession, parent: string, story: string): TableCell {
  return tableCells(parent, draft.storySegments(parent)).get(story) ?? notInTable();
}

function covering(anchors: readonly TableCellAnchor[], row: number, column: number) {
  return anchors.find(anchor => anchor.row <= row && row < anchor.row + anchor.rowspan && anchor.column <= column && column < anchor.column + anchor.colspan);
}

function requireUnmergedRow(cell: TableCell) {
  if (cell.mergedRows.has(cell.anchor.row)) {
    throw new DocumentToolError('UNSUPPORTED_TARGET', 'This row participates in a vertical merge. Choose another row.');
  }
}

export function insertTableRow(draft: YrsSession, target: DocxParagraphTarget, parent: string, position: 'before' | 'after', texts: readonly string[]): DocxParagraphTarget[] {
  const cell = locate(draft, parent, target.story);
  requireUnmergedRow(cell);
  const { story, tableIndex, columns, anchors, anchor } = cell;
  const templates = Array.from({ length: columns }, (_, column) => covering(anchors, anchor.row, column)?.story);
  const receipt = draft.insertRow({ story, tableIndex, row: anchor.row, column: anchor.column }, position === 'before' ? 'above' : 'below');
  if (receipt.createdStoryIds.length !== texts.length) {
    throw new DocumentToolError('UNSUPPORTED_TARGET', 'A vertically merged cell crosses this row boundary. Insert next to another row.');
  }
  return receipt.createdStoryIds.map((cell, column) => {
    const paraId = receipt.newParaIds[column];
    const template = templates[column];
    const source: YrsStorySegment[] = template ? draft.storySegments(template) : [];
    const properties = source.find((segment): segment is Pilcrow => segment.kind === 'pilcrow')?.properties;
    for (const [key, value] of Object.entries(properties ?? {})) {
      if (PARAGRAPH_PROPERTIES.has(key)) draft.setParagraphAttr(paraId, key, value);
    }
    const run = source.find(segment => segment.kind === 'text');
    const attrs = Object.fromEntries(Object.entries(run?.attributes ?? {}).filter(([key]) => !PROTECTED.has(key)));
    if (texts[column]) draft.applyRawOps(cell, [{ op: 'insert', index: 0, text: texts[column], attrs }]);
    return { story: cell, paraId };
  });
}

export function deleteTableRow(draft: YrsSession, target: DocxParagraphTarget, parent: string): boolean {
  const cell = locate(draft, parent, target.story);
  requireUnmergedRow(cell);
  if (cell.rows === 1) {
    throw new DocumentToolError('UNSUPPORTED_TARGET', 'A table keeps at least one row.');
  }
  const { story, tableIndex, anchor } = cell;
  const at = { story, tableIndex, row: anchor.row, column: anchor.column };
  return draft.deleteRow({ anchor: at, head: at }).deletedTable;
}
