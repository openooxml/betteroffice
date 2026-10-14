import type { FormulaShapeDraft, FormulaShapeTreeDraft, ShapeSnapshot } from '@betteroffice/vsdx';

export interface ClipboardCellLocator { cellName: string; section?: string; sectionIndex?: number; rowIndex?: number; rowName?: string; rowType?: string; }
export interface ClipboardCell { locator: ClipboardCellLocator; name: string; formula?: string; value?: string; rowType?: string; }
export interface VsdxClipboardGlue { connectorSource: string; endpoint: string; targetSource: string; toCell: string; }
export interface VsdxClipboardEntry { pageId: string; name?: string; cells: ClipboardCell[]; text: string; pinX: number | null; pinY: number | null; pasteCount: number; sourceShapeId?: string; sourceId?: number; copySourceId?: number; copySourcePageId?: number; copyRefusal?: string; children: VsdxClipboardEntry[]; glue: VsdxClipboardGlue[]; }

export const PASTE_OFFSET = { x: 0.25, y: -0.25 } as const;
export const DUPLICATE_OFFSET = { x: -0.25, y: 0.25 } as const;

/** Every shape copies unless its subtree carries unportable content. */
export function canCopyShape(shape: ShapeSnapshot): boolean {
  return copyRefusalReason(shape) == null;
}

/** First unportable reason in the subtree, or null when the tree copies losslessly. */
export function copyRefusalReason(shape: ShapeSnapshot): string | null {
  if (shape.copyRefusal != null) return shape.copyRefusal;
  for (const child of shape.children) {
    const reason = copyRefusalReason(child);
    if (reason != null) return reason;
  }
  return null;
}

/** Chain of groups above a shape, innermost last, or null when the shape is not there. */
function ancestorChain(shapes: readonly ShapeSnapshot[], shapeId: string, depth = 0): ShapeSnapshot[] | null {
  if (depth >= 256) return null;
  for (const shape of shapes) {
    if (shape.id === shapeId) return [];
    const nested = ancestorChain(shape.children, shapeId, depth + 1);
    if (nested) return [shape, ...nested];
  }
  return null;
}

/** Page-space offset of a shape's group-local pins; null when an ancestor rotates or flips. */
export function ancestorPinOffset(shapes: readonly ShapeSnapshot[], shapeId: string): { dx: number; dy: number } | null {
  const chain = ancestorChain(shapes, shapeId);
  if (!chain) return null;
  let dx = 0;
  let dy = 0;
  for (const ancestor of chain) {
    const turned = ['Angle', 'FlipX', 'FlipY'].some((cell) => Math.abs(resolvedNumeric(ancestor, cell) ?? 0) > 1e-9);
    if (turned) return null;
    dx += (resolvedNumeric(ancestor, 'PinX') ?? 0) - (resolvedNumeric(ancestor, 'LocPinX') ?? 0);
    dy += (resolvedNumeric(ancestor, 'PinY') ?? 0) - (resolvedNumeric(ancestor, 'LocPinY') ?? 0);
  }
  return { dx, dy };
}

/** Snapshot a shape into an in-app clipboard entry, preserving every cell formula. */
export function buildClipboardEntry(pageId: string, shape: ShapeSnapshot, text: string, options?: { textFor?: (shape: ShapeSnapshot) => string; glue?: VsdxClipboardGlue[]; pinOffset?: { dx: number; dy: number } }): VsdxClipboardEntry {
  const reason = copyRefusalReason(shape);
  if (reason != null) throw new Error(`vsdx copy is not supported for shape ${shape.id} with ${reason}`);
  const textFor = options?.textFor ?? (() => '');
  const node = buildNode(shape, text, textFor);
  const offset = options?.pinOffset ?? { dx: 0, dy: 0 };
  const pin = (name: 'PinX' | 'PinY', delta: number) => {
    const own = resolvedNumeric(shape, name);
    return own == null ? null : own + delta;
  };
  return { pageId, name: node.name, cells: node.cells, text: node.text, pinX: pin('PinX', offset.dx), pinY: pin('PinY', offset.dy), pasteCount: 0, sourceShapeId: shape.id, sourceId: shape.sourceId, ...(shape.copySourceId != null ? { copySourceId: shape.copySourceId } : {}), ...(shape.copySourcePageId != null ? { copySourcePageId: shape.copySourcePageId } : {}), ...(shape.copyRefusal != null ? { copyRefusal: shape.copyRefusal } : {}), children: node.children, glue: options?.glue ?? [] };
}

function buildNode(shape: ShapeSnapshot, text: string, textFor: (shape: ShapeSnapshot) => string): { name?: string; cells: ClipboardCell[]; text: string; children: VsdxClipboardEntry[] } {
  const cells: ClipboardCell[] = shape.cells
    .filter((cell) => cell.locator.cellName.length > 0)
    .map((cell) => {
      const locator: ClipboardCellLocator = { cellName: cell.locator.cellName };
      if (cell.locator.section != null) locator.section = cell.locator.section;
      if (cell.locator.sectionIndex != null) locator.sectionIndex = cell.locator.sectionIndex;
      const row = cell.locator.row;
      if (row && 'index' in row) locator.rowIndex = (row as { index: number }).index;
      else if (row && 'name' in row) locator.rowName = (row as { name: string }).name;
      if (cell.rowType != null) locator.rowType = cell.rowType;
      const entry: ClipboardCell = { locator, name: cell.locator.cellName };
      if (cell.formula != null) entry.formula = cell.formula;
      if (cell.value != null) entry.value = cell.value;
      if (cell.rowType != null) entry.rowType = cell.rowType;
      return entry;
    });
  return {
    ...(shape.name != null ? { name: shape.name } : {}),
    cells,
    text,
    children: shape.children.map((child) => {
      const node = buildNode(child, textFor(child), textFor);
      return { pageId: '', name: node.name, cells: node.cells, text: node.text, pinX: null, pinY: null, pasteCount: 0, sourceShapeId: child.id, sourceId: child.sourceId, ...(child.copySourceId != null ? { copySourceId: child.copySourceId } : {}), ...(child.copySourcePageId != null ? { copySourcePageId: child.copySourcePageId } : {}), ...(child.copyRefusal != null ? { copyRefusal: child.copyRefusal } : {}), children: node.children, glue: [] };
    }),
  };
}

/** A clipboard entry needs the tree paste path when it carries children or glue. */
export function isTreeEntry(entry: VsdxClipboardEntry): boolean {
  return entry.children.length > 0 || entry.glue.length > 0;
}

/** Draft a pasted shape, offsetting PinX/PinY so the copy lands visibly apart. */
export function draftForPaste(entry: VsdxClipboardEntry, dx: number, dy: number): FormulaShapeDraft {
  const cells = entry.cells.map((cell) => ({
    locator: toDraftLocator(cell),
    name: cell.name,
    ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
    ...(cell.value !== undefined ? { value: cell.value } : {}),
  }));
  applyPinOffset(cells, 'PinX', entry.pinX, dx);
  applyPinOffset(cells, 'PinY', entry.pinY, dy);
  return { ...(entry.name !== undefined ? { name: entry.name } : {}), cells };
}

/** Draft a pasted group subtree, offsetting only the root so children keep relative positions. */
export function draftTreeForPaste(entry: VsdxClipboardEntry, dx: number, dy: number): FormulaShapeTreeDraft {
  return toTreeDraft(entry, dx, dy, true);
}

function toTreeDraft(entry: VsdxClipboardEntry, dx: number, dy: number, isRoot: boolean): FormulaShapeTreeDraft {
  const cells = entry.cells.map((cell) => ({
    locator: toDraftLocator(cell),
    name: cell.name,
    ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
    ...(cell.value !== undefined ? { value: cell.value } : {}),
  }));
  if (isRoot) {
    applyPinOffset(cells, 'PinX', entry.pinX, dx);
    applyPinOffset(cells, 'PinY', entry.pinY, dy);
  }
  return {
    ...(entry.name !== undefined ? { name: entry.name } : {}),
    cells,
    text: entry.text,
    ...(entry.copySourceId !== undefined ? { copySourceId: entry.copySourceId } : {}),
    ...(entry.copySourcePageId !== undefined ? { copySourcePageId: entry.copySourcePageId } : {}),
    ...(entry.sourceShapeId !== undefined ? { sourceShapeId: entry.sourceShapeId } : {}),
    ...(entry.sourceId !== undefined ? { sourceId: entry.sourceId } : {}),
    ...(entry.copyRefusal !== undefined ? { copyRefusal: entry.copyRefusal } : {}),
    ...(isRoot && entry.glue.length > 0 ? { glue: entry.glue.map((glue) => ({ ...glue })) } : {}),
    ...(entry.children.length > 0 ? { children: entry.children.map((child) => toTreeDraft(child, 0, 0, false)) } : {}),
  };
}

/** Resolved numeric value of a root cell, or null when absent or non-numeric. */
export function resolvedNumeric(shape: ShapeSnapshot, name: string): number | null {
  const cell = shape.cells.find((item) => item.locator.section == null && item.locator.row == null && item.locator.cellName === name);
  const raw = cell?.value ?? cell?.formula;
  if (raw == null || !raw.trim()) return null;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : null;
}

function toDraftLocator(cell: ClipboardCell): ClipboardCellLocator {
  const locator: ClipboardCellLocator = { cellName: cell.locator.cellName };
  if (cell.locator.section !== undefined) locator.section = cell.locator.section;
  if (cell.locator.sectionIndex !== undefined) locator.sectionIndex = cell.locator.sectionIndex;
  if (cell.locator.rowIndex !== undefined) locator.rowIndex = cell.locator.rowIndex;
  if (cell.locator.rowName !== undefined) locator.rowName = cell.locator.rowName;
  if (cell.locator.rowType !== undefined) locator.rowType = cell.locator.rowType;
  return locator;
}

function applyPinOffset(cells: Array<{ locator: ClipboardCellLocator; name: string; formula?: string; value?: string }>, name: 'PinX' | 'PinY', base: number | null, delta: number): void {
  const next = base == null ? delta : base + delta;
  const formula = toFormula(next);
  const existing = cells.find((cell) => cell.locator.section === undefined && cell.locator.rowIndex === undefined && cell.locator.rowName === undefined && cell.name === name);
  if (existing) { existing.formula = formula; delete existing.value; }
  else cells.push({ locator: { cellName: name }, name, formula });
}

/** Canonical ShapeSheet number literal. */
export function toFormula(value: number): string {
  const rounded = Number(value.toFixed(6));
  return String(Object.is(rounded, -0) ? 0 : rounded);
}
