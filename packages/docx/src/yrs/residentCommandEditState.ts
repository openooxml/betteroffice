import type { YrsHyperlinkAttrs, YrsLoc, YrsSelectionContext } from './index';
import { readCommandSelection } from './residentCommandAnchors';
import { readCellSelection, resolveLiveCell } from './residentCommandTableSelection';
import { tableAnchors, type TablePayload } from './storyPlainText';
import type {
  DocxCommandTableContext,
  DocxEditState,
  DocxStoredMarks,
  ResidentDocxCommandEngine,
  ResidentDocxCommandState,
} from './residentCommandProtocol';

export function withStoredMarks(context: YrsSelectionContext, stored: DocxStoredMarks | null): YrsSelectionContext {
  if (!stored) return context;
  const base = stored.clear ? {
    ...context, bold: false as const, italic: false as const, underline: false as const,
    strike: false as const, superscript: false as const, subscript: false as const,
    fontFamily: null, fontSize: null, color: null, highlight: null,
  } : context;
  const delta = stored.delta;
  const script = delta.other ?? {};
  return {
    ...base,
    bold: delta.bold === undefined ? base.bold : delta.bold === true,
    italic: delta.italic === undefined ? base.italic : delta.italic === true,
    underline: delta.underline === undefined ? base.underline : delta.underline !== false && delta.underline !== null,
    strike: delta.strike === undefined ? base.strike : delta.strike !== false && delta.strike !== null,
    superscript: script.superscript === undefined ? base.superscript : script.superscript === true,
    subscript: script.subscript === undefined ? base.subscript : script.subscript === true,
    fontFamily: delta.fontFamily === undefined ? base.fontFamily : delta.fontFamily?.ascii ?? delta.fontFamily?.hAnsi ?? null,
    fontSize: delta.fontSize === undefined ? base.fontSize : delta.fontSize == null ? null : Math.round(delta.fontSize * 2),
    color: delta.color === undefined ? base.color : delta.color?.rgb ?? delta.color?.themeColor ?? null,
    highlight: delta.highlight === undefined ? base.highlight : delta.highlight,
  };
}

function hyperlinkAt(engine: ResidentDocxCommandEngine, story: string, paraId: string, offset: number): YrsHyperlinkAttrs | null {
  const index = engine.geometryReader.paragraphSpans(story).findIndex((paragraph) => paragraph.paraId === paraId);
  if (index < 0) return null;
  let start = 0;
  let boundary: YrsHyperlinkAttrs | null = null;
  for (const segment of engine.paragraphSegments(story, [index])[0] ?? []) {
    const end = start + (segment.kind === 'text' ? segment.text.length : 1);
    const link = segment.attributes.hyperlink as YrsHyperlinkAttrs | undefined;
    if (link && typeof link.href === 'string' && link.href) {
      const attrs = {
        href: link.href,
        ...(typeof link.tooltip === 'string' ? { tooltip: link.tooltip } : {}),
        ...(typeof link.rId === 'string' ? { rId: link.rId } : {}),
      };
      if (offset > start && offset < end) return attrs;
      if (!boundary && (offset === start || offset === end)) boundary = attrs;
    }
    start = end;
  }
  return boundary;
}

function revisionAt(engine: ResidentDocxCommandEngine, state: ResidentDocxCommandState, head: YrsLoc): string | undefined {
  const reader = engine.geometryReader;
  const at = (paraId: string, offset: number) => reader.locateParagraph(head.story, paraId).start + offset;
  const position = at(head.paraId, head.offset);
  const holding = reader.listRevisions().filter(({ range }) => {
    if (range.story !== head.story) return false;
    try {
      return at(range.start.paraId, range.start.offset) <= position && position <= at(range.end.paraId, range.end.offset);
    } catch {
      return false;
    }
  });
  if (!holding.length) return undefined;
  const proposed = new Set(state.proposals?.snapshot().proposals.flatMap((proposal) => proposal.revisionIds));
  return holding.find(({ revisionId }) => !proposed.has(revisionId))?.revisionId;
}

function object(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
}

function validTablePayload(value: unknown): value is TablePayload {
  return object(value) &&
    (value.tblPr === undefined || object(value.tblPr)) &&
    (value.grid === undefined || Array.isArray(value.grid)) &&
    Array.isArray(value.rows) && value.rows.every((row: unknown) => object(row) &&
      Array.isArray(row.cells) && row.cells.every((cell: unknown) => object(cell) &&
        typeof cell.story === 'string' && (cell.tcPr === undefined || object(cell.tcPr))));
}

function tableContext(engine: ResidentDocxCommandEngine, story: string): DocxCommandTableContext | null {
  const focused = resolveLiveCell(engine, story);
  if (!focused || ![focused.tableIndex, focused.row, focused.column].every(Number.isSafeInteger)) return null;
  const selected = readCellSelection(engine);
  const sameTable = !!selected && selected.anchor.story === focused.story &&
    selected.anchor.tableIndex === focused.tableIndex &&
    selected.head.story === focused.story && selected.head.tableIndex === focused.tableIndex;
  const head = sameTable && selected ? selected.head : focused;
  const payload = engine.tablePayload(focused.story, focused.tableIndex);
  if (!validTablePayload(payload)) return null;
  const { anchors, columns } = tableAnchors(payload);
  const cell = anchors.find((entry) => entry.row <= head.row && head.row < entry.row + entry.rowspan &&
    entry.column <= head.column && head.column < entry.column + entry.colspan);
  if (!cell) return null;
  const focusedCell = anchors.find((entry) => entry.row <= focused.row && focused.row < entry.row + entry.rowspan &&
    entry.column <= focused.column && focused.column < entry.column + entry.colspan);
  if (!focusedCell) return null;
  const properties = payload.rows.flatMap((row) => row.cells ?? []).find((entry) => entry.story === focusedCell.story)?.tcPr;
  const borders = properties?.borders;
  let cellBorderColor: DocxCommandTableContext['cellBorderColor'];
  if (borders && typeof borders === 'object') {
    for (const border of Object.values(borders)) {
      const color = border && typeof border === 'object' ? (border as { color?: unknown }).color : undefined;
      if (typeof color === 'string') { cellBorderColor = { rgb: color.replace(/^#/, '') }; break; }
      if (color && typeof color === 'object') {
        cellBorderColor = color as NonNullable<typeof cellBorderColor>;
        break;
      }
    }
  }
  return {
    story: focused.story, tableIndex: focused.tableIndex, rowIndex: head.row, columnIndex: head.column,
    rowCount: payload.rows.length, columnCount: columns,
    hasMultiCellSelection: !!(sameTable && selected && (selected.anchor.row !== selected.head.row ||
      selected.anchor.column !== selected.head.column)),
    canSplitCell: true,
    rowSpan: focusedCell.rowspan, columnSpan: focusedCell.colspan,
    ...(typeof payload.tblPr?.justification === 'string' ? { justification: payload.tblPr.justification } : {}),
    ...(cellBorderColor ? { cellBorderColor } : {}),
    ...(typeof properties?.backgroundColor === 'string'
      ? { cellBackgroundColor: properties.backgroundColor.replace(/^#/, '') } : {}),
  };
}

export function projectDocxEditState(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  resolved = state.selection ? readCommandSelection(engine, state) : null,
  context?: YrsSelectionContext
): DocxEditState {
  const result: DocxEditState = {
    ...(state.activePart !== undefined ? { activePart: state.activePart } : {}),
    documentVersion: engine.proposalEngine?.version(), readOnly: state.readOnly, toolbar: null, effectiveFormatting: null, storedMarks: null,
    canUndo: engine.canUndo?.() ?? false, canRedo: engine.canRedo?.() ?? false,
    hyperlink: null, tableContext: null,
    ...(state.proposals ? { review: { version: engine.proposalEngine.version(), previewVersion: state.proposals.snapshot().previewVersion } } : {}),
  };
  if (!resolved?.ok) return result;
  const { range, selection } = resolved;
  const collapsed = range.start.paraId === range.end.paraId && range.start.offset === range.end.offset;
  const stored = collapsed && state.storedMarks?.story === range.story && state.storedMarks.paraId === range.start.paraId
    ? state.storedMarks : null;
  const toolbar = withStoredMarks(context ?? engine.selectionContext(range), stored);
  const defaults = toolbar.paragraphProperties.defaultTextFormatting as { fontSize?: number | { size?: number } } | undefined;
  const defaultSize = typeof defaults?.fontSize === 'number' ? defaults.fontSize : defaults?.fontSize?.size;
  const fontSize = toolbar.fontSize ?? (typeof defaultSize === 'number' && Number.isFinite(defaultSize) ? defaultSize : null);
  result.toolbar = toolbar;
  result.storedMarks = stored ? structuredClone(stored) : null;
  result.effectiveFormatting = {
    bold: toolbar.bold, italic: toolbar.italic, underline: toolbar.underline, strike: toolbar.strike,
    superscript: toolbar.superscript, subscript: toolbar.subscript, fontFamily: toolbar.fontFamily,
    fontSizePt: fontSize == null ? null : fontSize / 2, color: toolbar.color, highlight: toolbar.highlight,
    styleId: toolbar.styleId, alignment: toolbar.alignment,
  };
  result.hyperlink = hyperlinkAt(engine, selection.focus.story, selection.focus.paraId, selection.focus.offset);
  result.tableContext = tableContext(engine, selection.focus.story);
  const revision = revisionAt(engine, state, selection.focus);
  if (revision !== undefined) result.currentRevisionId = revision;
  return result;
}
