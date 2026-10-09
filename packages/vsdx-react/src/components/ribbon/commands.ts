import { createContext, createElement, useContext, useMemo } from 'react';
import type { ReactNode } from 'react';
import type { CellWriteProbe, CellWriteQuery, DiagramHandle, DiagramSnapshot, FormulaShapeDraft, PageDisplayList, PagePrimitive, PageSnapshot, Paint, ShapePrimitive, ShapeSnapshot } from '@betteroffice/vsdx';
import type { VsdxShapeSelection } from '../../VsdxEditor';
import { standardShapeById } from '../shapes/shapeLibrary';
import { DUPLICATE_OFFSET, PASTE_OFFSET, ancestorPinOffset, buildClipboardEntry, canCopyShape, draftForPaste, draftTreeForPaste, isTreeEntry } from './clipboard';
import type { VsdxClipboardEntry } from './clipboard';

export type RibbonCommandId =
  | 'undo' | 'redo' | 'delete' | 'cut' | 'copy' | 'paste' | 'duplicate'
  | 'fillColor' | 'lineColor' | 'lineWeight' | 'linePattern'
  | 'bringToFront' | 'bringForward' | 'sendBackward' | 'sendToBack'
  | 'rotateLeft' | 'rotateRight' | 'flipHorizontal' | 'flipVertical' | 'addShape' | 'download' | 'pageBreaks';

export interface RibbonCommand { id: RibbonCommandId; run: (value?: string) => void; enabled: boolean; active?: boolean; value?: string; }
export type RibbonCommands = Record<RibbonCommandId, RibbonCommand>;

export const RibbonCommandsContext = createContext<RibbonCommands | null>(null);

export interface RibbonCommandsProviderProps {
  handle: DiagramHandle | null;
  snapshot: DiagramSnapshot | null;
  pageId?: string;
  selection: readonly VsdxShapeSelection[];
  frame?: PageDisplayList | null;
  clipboard?: VsdxClipboardEntry | null;
  onClipboardChange?: (next: VsdxClipboardEntry | null) => void;
  onSelectShape?: (selection: VsdxShapeSelection) => void;
  onMutation: () => void;
  onError: (error: unknown) => void;
  onDownload: (bytes: Uint8Array) => void;
  pageBreaks?: PageBreakToggle;
  probes?: ReadonlyMap<string, ShapeWriteProbes>;
  children: ReactNode;
}

/** Whether the page-break overlay is shown, and how to flip it. */
export interface PageBreakToggle { shown: boolean; toggle: () => void; }

export interface ShapePlacement { shape: ShapeSnapshot; index: number; siblings: readonly ShapeSnapshot[]; }

export function findShapePlacement(shapes: readonly ShapeSnapshot[], shapeId: string, depth = 0): ShapePlacement | null {
  if (depth >= 256) return null;
  const index = shapes.findIndex((shape) => shape.id === shapeId);
  if (index >= 0) return { shape: shapes[index], index, siblings: shapes };
  for (const shape of shapes) {
    const nested = findShapePlacement(shape.children, shapeId, depth + 1);
    if (nested) return nested;
  }
  return null;
}

export function pageById(pages: readonly PageSnapshot[], pageId: string | undefined): PageSnapshot | null {
  return (pageId === undefined ? pages[0] : pages.find((page) => page.id === pageId)) ?? null;
}

function placementIn(pages: readonly PageSnapshot[], selection: VsdxShapeSelection): ShapePlacement | null {
  const page = pages.find((item) => item.id === selection.pageId);
  return page ? findShapePlacement(page.shapes, selection.shapeId) : null;
}

function placementsIn(pages: readonly PageSnapshot[], selection: readonly VsdxShapeSelection[]): Array<{ selection: VsdxShapeSelection; placement: ShapePlacement }> {
  const result: Array<{ selection: VsdxShapeSelection; placement: ShapePlacement }> = [];
  for (const item of selection) {
    const page = pages.find((entry) => entry.id === item.pageId);
    const placement = page ? findShapePlacement(page.shapes, item.shapeId) : null;
    if (placement) result.push({ selection: item, placement });
  }
  return result;
}

function findCell(shape: ShapeSnapshot | null, name: string) {
  return shape?.cells.find((item) => item.locator.section === null && item.locator.row === null && item.locator.cellName === name);
}

export function cellValue(shape: ShapeSnapshot | null, name: string): string | undefined {
  const current = findCell(shape, name);
  return current?.value ?? current?.formula ?? undefined;
}

function cellFormula(shape: ShapeSnapshot | null, name: string): string | undefined {
  const current = findCell(shape, name);
  return current?.formula ?? current?.value ?? undefined;
}

function color(value: string | undefined, fallback: string): string {
  const hex = value?.match(/#[0-9a-f]{6}/i)?.[0];
  if (hex) return hex;
  const rgb = value?.match(/^RGB\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
  return rgb ? `#${rgb.slice(1).map((channel) => Math.min(255, Number(channel)).toString(16).padStart(2, '0')).join('')}` : fallback;
}

function colorFormula(value = '#000000'): string {
  const hex = /^#[0-9a-f]{6}$/i.test(value) ? value.slice(1) : '000000';
  return `RGB(${[0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16)).join(',')})`;
}

export function numberValue(value: string | undefined): number {
  const result = Number(value ?? '0');
  return Number.isFinite(result) ? result : 0;
}

/** Largest LinePattern index Visio documents. 0 clears the stroke, 1 is solid. */
export const LINE_PATTERN_MAX = 23;

/** Selectable dash pattern indexes. Only 0 and 1 have stable Visio-wide meanings. */
export const LINE_PATTERN_VALUES: readonly string[] = Array.from({ length: LINE_PATTERN_MAX + 1 }, (_, index) => String(index));

const LINE_WEIGHT_PATTERN = /^(-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)\s*(in|dl|cm|mm|pt|pica|ft|m)?$/i;

/** Validated line weight, or null when the text is not a positive length. */
export function parseLineWeightInput(raw: string): string | null {
  const match = raw.trim().match(LINE_WEIGHT_PATTERN);
  if (!match) return null;
  const magnitude = Number(match[1]);
  if (!Number.isFinite(magnitude) || magnitude <= 0) return null;
  const unit = (match[2] ?? '').toLowerCase();
  return unit ? `${match[1]} ${unit}` : match[1];
}

/** Validated line pattern index, or null when outside 0..LINE_PATTERN_MAX. */
export function parseLinePatternInput(raw: string): string | null {
  const text = raw.trim();
  if (!/^\d+$/.test(text)) return null;
  const index = Number(text);
  return Number.isSafeInteger(index) && index >= 0 && index <= LINE_PATTERN_MAX ? String(index) : null;
}

/** True when a raw ribbon value would change the stored formula. */
export function isFormulaChange(current: string | undefined, raw: string, parse: (value: string) => string | null): boolean {
  const next = parse(raw);
  if (next === null) return false;
  const baseline = current !== undefined ? parse(current) : null;
  return baseline === null ? raw.trim() !== (current ?? '') : next !== baseline;
}

export function numericCellValue(shape: ShapeSnapshot, name: string, fallback?: number): number {
  const value = cellValue(shape, name);
  if (value === undefined && fallback !== undefined) return fallback;
  const parsed = value?.trim() ? Number(value) : Number.NaN;
  if (!Number.isFinite(parsed)) throw new Error(`Shape cell ${name} has no resolved numeric value.`);
  return parsed;
}

/** True when a ShapeSheet lock cell evaluates to the enabled value 1. */
export function lockCellEnabled(shape: ShapeSnapshot | null, name: string): boolean {
  return Number(cellValue(shape, name)) === 1;
}

/** The engine's verdict per cell for one shape. */
export type ShapeWriteProbes = ReadonlyMap<string, CellWriteProbe>;

/** What the editor gates a control on, probed together per shape. */
export const GATED_CELLS: readonly CellWriteQuery[] = [
  { cellName: 'LockDelete', gesture: 'delete' },
  { cellName: 'PinX', gesture: 'moveX' },
  { cellName: 'PinY', gesture: 'moveY' },
  { cellName: 'Width', gesture: 'resizeWidth' },
  { cellName: 'Height', gesture: 'resizeHeight' },
  { cellName: 'Angle', gesture: 'rotate' },
  { cellName: 'FillForegnd' }, { cellName: 'LineColor' }, { cellName: 'LineWeight' },
  { cellName: 'LinePattern' }, { cellName: 'FlipX' }, { cellName: 'FlipY' },
];

/** True when the engine says a write to this cell would be refused. */
export function isCellWriteBlocked(probes: ShapeWriteProbes | null | undefined, cellName: string): boolean {
  return probes?.get(cellName)?.allowed === false;
}

/** True when the engine says the delete gesture would be refused. */
export function isDeleteBlocked(probes: ShapeWriteProbes | null | undefined): boolean {
  return isCellWriteBlocked(probes, 'LockDelete');
}

/** Cells a handle resize writes; any one refused disables the handles. */
export const HANDLE_RESIZE_CELLS = ['PinX', 'PinY', 'Width', 'Height'] as const;

/** True when the engine says any cell a handle resize writes would be refused. */
export function isHandleResizeBlocked(probes: ShapeWriteProbes | null | undefined): boolean {
  return HANDLE_RESIZE_CELLS.some((cell) => isCellWriteBlocked(probes, cell));
}

/** A shape's own geometry primitive; text boxes share its id, so kind is part of the match. */
function findShapePrimitive(primitives: readonly PagePrimitive[], id: string, depth = 0): ShapePrimitive | null {
  if (depth >= 256) return null;
  for (const primitive of primitives) {
    if (primitive.kind === 'shape' && primitive.id === id) return primitive;
    if (primitive.kind === 'group') {
      const nested = findShapePrimitive(primitive.primitives, id, depth + 1);
      if (nested) return nested;
    }
  }
  return null;
}

const HEX_COLOUR = /^#[0-9a-f]{6}$/i;

/** A gradient stands in for its first stop, the way Visio's fill swatch shows it. */
function paintSwatch(paint: Paint | undefined): string | undefined {
  const colour = paint?.kind === 'solid' ? paint.color : paint?.kind === 'gradient' ? paint.stops[0]?.color : undefined;
  return colour !== undefined && HEX_COLOUR.test(colour) ? colour : undefined;
}

/** Rendered stroke/fill colours for a shape, when the display list resolves one. */
export function frameSwatch(
  frame: PageDisplayList | null | undefined,
  page: PageSnapshot | null,
  shape: ShapeSnapshot | null,
): { fill?: string; line?: string } {
  if (!frame || !page || !shape) return {};
  const primitive = findShapePrimitive(frame.primitives, `${page.sourcePartPath}:${shape.sourceId}`);
  if (!primitive) return {};
  const fill = paintSwatch(primitive.fill);
  const line = primitive.stroke && HEX_COLOUR.test(primitive.stroke.color) ? primitive.stroke.color : undefined;
  return { fill, line };
}

/** Snapshot the selection into an in-app clipboard entry, carrying the whole subtree. */
export function copySelection(handle: DiagramHandle, selection: VsdxShapeSelection): VsdxClipboardEntry {
  const snapshot = handle.snapshot();
  const placement = placementIn(snapshot.pages, selection);
  if (!placement) throw new Error(`vsdx shape ${selection.shapeId} is no longer part of the diagram`);
  const page = snapshot.pages.find((item) => item.id === selection.pageId);
  const pinOffset = page ? ancestorPinOffset(page.shapes, selection.shapeId) : null;
  if (!pinOffset) throw new Error(`vsdx copy is not supported for shape ${selection.shapeId} inside a rotated or flipped group`);
  let text = '';
  try { text = handle.shapeText(selection.pageId, selection.shapeId); }
  catch { text = ''; }
  const textFor = (shape: ShapeSnapshot): string => {
    try { return handle.shapeText(selection.pageId, shape.id); }
    catch { return ''; }
  };
  const glue = subtreeGlueOf(handle, selection.pageId, selection.shapeId);
  return buildClipboardEntry(selection.pageId, placement.shape, text, { textFor, glue, pinOffset });
}

function subtreeGlueOf(handle: DiagramHandle, pageId: string, shapeId: string): VsdxClipboardEntry['glue'] {
  const subtree = (handle as unknown as { subtreeGlue?: (pageId: string, shapeId: string) => Array<{ connectorSource: string; endpoint: string; targetSource: string; toCell: string }> }).subtreeGlue;
  if (typeof subtree !== 'function') return [];
  return subtree.call(handle, pageId, shapeId).map((glue) => ({ connectorSource: glue.connectorSource, endpoint: glue.endpoint, targetSource: glue.targetSource, toCell: glue.toCell }));
}

/** Paste a clipboard entry with a model-space offset as one atomic shape addition. */
export function pasteEntry(handle: DiagramHandle, targetPageId: string, entry: VsdxClipboardEntry, dx: number, dy: number): { receipt: { shapeId: string }; entry: VsdxClipboardEntry } {
  const page = pageById(handle.snapshot().pages, targetPageId);
  if (!page) throw new Error(`vsdx page ${targetPageId} is no longer part of the diagram`);
  if (isTreeEntry(entry)) {
    const addTree = (handle as unknown as { addShapeTree?: (pageId: string, draft: unknown) => { shapeId: string } }).addShapeTree;
    if (typeof addTree !== 'function') throw new Error('vsdx group paste needs a diagram handle with addShapeTree');
    const receipt = addTree.call(handle, page.id, draftTreeForPaste(entry, dx, dy));
    return { receipt, entry: { ...entry, pasteCount: entry.pasteCount + 1 } };
  }
  const draft = draftForPaste(entry, dx, dy);
  const receipt = addShapeWithText(handle, page.id, draft, entry.text);
  return { receipt, entry: { ...entry, pasteCount: entry.pasteCount + 1 } };
}

/** Duplicate a clipboard entry without touching the clipboard, as one atomic addition. */
export function duplicateEntry(handle: DiagramHandle, targetPageId: string, entry: VsdxClipboardEntry, dx: number, dy: number): { shapeId: string } {
  const page = pageById(handle.snapshot().pages, targetPageId);
  if (!page) throw new Error(`vsdx page ${targetPageId} is no longer part of the diagram`);
  if (isTreeEntry(entry)) {
    const addTree = (handle as unknown as { addShapeTree?: (pageId: string, draft: unknown) => { shapeId: string } }).addShapeTree;
    if (typeof addTree !== 'function') throw new Error('vsdx group duplicate needs a diagram handle with addShapeTree');
    return addTree.call(handle, page.id, draftTreeForPaste(entry, dx, dy));
  }
  return addShapeWithText(handle, page.id, draftForPaste(entry, dx, dy), entry.text);
}

export function addShapeWithText(handle: DiagramHandle, pageId: string, draft: FormulaShapeDraft, text: string): { shapeId: string } {
  if (typeof (handle as { addShapeWithText?: unknown }).addShapeWithText === 'function') {
    return (handle as unknown as { addShapeWithText: (pageId: string, draft: FormulaShapeDraft, text: string) => { shapeId: string } }).addShapeWithText(pageId, draft, text);
  }
  const receipt = handle.addShape(pageId, draft);
  if (text) handle.setShapeText(pageId, receipt.shapeId, text);
  return receipt;
}

/** Key for one shape's probes inside a selection-wide map. */
export const probeKey = (pageId: string, shapeId: string): string => `${pageId}${shapeId}`;

/** Asks the engine what it would refuse on one shape. */
export function shapeWriteProbes(handle: DiagramHandle | null, pageId: string, shapeId: string, cells: readonly CellWriteQuery[] = GATED_CELLS): ShapeWriteProbes | null {
  if (typeof handle?.probeCellWrites !== 'function') return null;
  try {
    return new Map(handle.probeCellWrites(pageId, shapeId, cells).map((probe) => [probe.cellName, probe]));
  } catch { return null; }
}

/** The same, for every shape in a selection. */
export function selectionWriteProbes(handle: DiagramHandle | null, selection: readonly VsdxShapeSelection[], cells: readonly CellWriteQuery[] = GATED_CELLS): Map<string, ShapeWriteProbes> {
  const probes = new Map<string, ShapeWriteProbes>();
  for (const item of selection) {
    const shape = shapeWriteProbes(handle, item.pageId, item.shapeId, cells);
    if (shape) probes.set(probeKey(item.pageId, item.shapeId), shape);
  }
  return probes;
}

/** True when the engine says a rotation would be refused. */
export function isRotateBlocked(probes: ShapeWriteProbes | null | undefined): boolean {
  return isCellWriteBlocked(probes, 'Angle');
}

export function createRibbonCommands(
  handle: DiagramHandle | null,
  selection: readonly VsdxShapeSelection[],
  pageId: string | undefined,
  onMutation: () => void,
  onError: (error: unknown) => void,
  onDownload: (bytes: Uint8Array) => void,
  frame?: PageDisplayList | null,
  clipboard: VsdxClipboardEntry | null = null,
  onClipboardChange: (next: VsdxClipboardEntry | null) => void = () => {},
  onSelectShape: (selection: VsdxShapeSelection) => void = () => {},
  pageBreaks?: PageBreakToggle,
  probes: ReadonlyMap<string, ShapeWriteProbes> = selectionWriteProbes(handle, selection),
): RibbonCommands {
  const execute = (operation: (current: DiagramHandle, selected: readonly VsdxShapeSelection[]) => void, needsSelection = false) => () => {
    if (!handle || (needsSelection && selection.length === 0)) return;
    try { operation(handle, selection); onMutation(); } catch (error) { onError(error); }
  };
  const pages = handle ? handle.snapshot().pages : [];
  const placements = placementsIn(pages, selection);
  const first = placements[0];
  const single = placements.length === 1 ? placements[0] : null;
  const shape = first?.placement.shape ?? null;
  const selected = placements.length > 0;
  const probesFor = (entry: { selection: VsdxShapeSelection }) => probes.get(probeKey(entry.selection.pageId, entry.selection.shapeId)) ?? null;
  const activePage = first ? pages.find((page) => page.id === first.selection.pageId) ?? null : null;
  const swatch = frameSwatch(frame, activePage, shape);
  const copyable = Boolean(single && canCopyShape(single.placement.shape) && activePage && ancestorPinOffset(activePage.shapes, single.selection.shapeId));
  const topIndex = single ? single.placement.siblings.length - 1 : 0;
  const livePlacements = (currentHandle: DiagramHandle) => placementsIn(currentHandle.snapshot().pages, selection);
  const formula = (cellName: string, value: string) => execute((currentHandle) => {
    const live = livePlacements(currentHandle);
    if (live.length === 0) return;
    currentHandle.setCellFormulas(live.map(({ selection: item }) => ({ pageId: item.pageId, shapeId: item.shapeId, cellName, formula: value })));
  }, true);
  const reorderTo = (target: (placement: ShapePlacement) => number, allowed: (placement: ShapePlacement) => boolean) => execute((currentHandle) => {
    const live = livePlacements(currentHandle);
    if (live.length !== 1) return;
    const { selection: item, placement } = live[0];
    if (allowed(placement)) currentHandle.reorderShape(item.pageId, item.shapeId, target(placement));
  }, true);
  const setNumeric = (cellName: string, next: (value: number) => string) => execute((currentHandle) => {
    const live = livePlacements(currentHandle);
    if (live.length === 0) return;
    currentHandle.setCellFormulas(live.map(({ selection: item, placement }) => ({ pageId: item.pageId, shapeId: item.shapeId, cellName, formula: next(numericCellValue(placement.shape, cellName, 0)) })));
  }, true);
  const commands = {
    undo: { id: 'undo', enabled: Boolean(handle?.canUndo()), run: execute((currentHandle) => { currentHandle.undo(); }) },
    redo: { id: 'redo', enabled: Boolean(handle?.canRedo()), run: execute((currentHandle) => { currentHandle.redo(); }) },
    delete: { id: 'delete', enabled: selected && placements.every((entry) => !isDeleteBlocked(probesFor(entry))), run: execute((currentHandle) => {
      const live = livePlacements(currentHandle);
      if (live.length === 0) return;
      currentHandle.deleteShapes(live.map(({ selection: item }) => ({ pageId: item.pageId, shapeId: item.shapeId })));
    }, true) },
    fillColor: { id: 'fillColor', enabled: selected && placements.every((entry) => !isCellWriteBlocked(probesFor(entry), 'FillForegnd')), value: swatch.fill ?? color(cellValue(shape, 'FillForegnd'), '#000000'), run: (value?: string) => formula('FillForegnd', colorFormula(value))() },
    lineColor: { id: 'lineColor', enabled: selected && placements.every((entry) => !isCellWriteBlocked(probesFor(entry), 'LineColor')), value: swatch.line ?? color(cellValue(shape, 'LineColor'), '#000000'), run: (value?: string) => formula('LineColor', colorFormula(value))() },
    lineWeight: {
      id: 'lineWeight', enabled: selected && placements.every((entry) => !isCellWriteBlocked(probesFor(entry), 'LineWeight')), value: cellFormula(shape, 'LineWeight'), run: (value?: string) => {
        if (value === undefined) return;
        const next = parseLineWeightInput(value);
        if (next === null || !isFormulaChange(cellFormula(shape, 'LineWeight'), value, parseLineWeightInput)) return;
        formula('LineWeight', next)();
      },
    },
    linePattern: {
      id: 'linePattern', enabled: selected && placements.every((entry) => !isCellWriteBlocked(probesFor(entry), 'LinePattern')), value: cellFormula(shape, 'LinePattern'), run: (value?: string) => {
        if (value === undefined) return;
        const next = parseLinePatternInput(value);
        if (next === null || !isFormulaChange(cellFormula(shape, 'LinePattern'), value, parseLinePatternInput)) return;
        formula('LinePattern', next)();
      },
    },
    bringToFront: { id: 'bringToFront', enabled: single !== null && single.placement.index < topIndex, run: reorderTo((placement) => placement.siblings.length - 1, (placement) => placement.index < placement.siblings.length - 1) },
    bringForward: { id: 'bringForward', enabled: single !== null && single.placement.index < topIndex, run: reorderTo((placement) => placement.index + 1, (placement) => placement.index < placement.siblings.length - 1) },
    sendBackward: { id: 'sendBackward', enabled: single !== null && single.placement.index > 0, run: reorderTo((placement) => placement.index - 1, (placement) => placement.index > 0) },
    sendToBack: { id: 'sendToBack', enabled: single !== null && single.placement.index > 0, run: reorderTo(() => 0, (placement) => placement.index > 0) },
    rotateLeft: { id: 'rotateLeft', enabled: selected && placements.every((entry) => !isRotateBlocked(probesFor(entry))), run: setNumeric('Angle', (value) => String(value - Math.PI / 2)) },
    rotateRight: { id: 'rotateRight', enabled: selected && placements.every((entry) => !isRotateBlocked(probesFor(entry))), run: setNumeric('Angle', (value) => String(value + Math.PI / 2)) },
    flipHorizontal: { id: 'flipHorizontal', enabled: selected && placements.every((entry) => !isCellWriteBlocked(probesFor(entry), 'FlipX')), active: numberValue(cellValue(shape, 'FlipX')) !== 0, run: setNumeric('FlipX', (value) => value === 0 ? '1' : '0') },
    flipVertical: { id: 'flipVertical', enabled: selected && placements.every((entry) => !isCellWriteBlocked(probesFor(entry), 'FlipY')), active: numberValue(cellValue(shape, 'FlipY')) !== 0, run: setNumeric('FlipY', (value) => value === 0 ? '1' : '0') },
    addShape: {
      id: 'addShape',
      enabled: Boolean(pageById(pages, pageId)),
      run: execute((currentHandle) => {
        const page = pageById(currentHandle.snapshot().pages, pageId);
        if (!page) throw new Error(`vsdx page ${pageId ?? ''} is no longer part of the diagram`);
        const rectangle = standardShapeById('rectangle');
        if (!rectangle) throw new Error('vsdx standard rectangle shape is unavailable');
        currentHandle.addShape(page.id, rectangle.draft(1, 1, rectangle.defaultSize.width, rectangle.defaultSize.height));
      }),
    },
    cut: {
      id: 'cut',
      enabled: copyable,
      run: () => {
        if (!handle || !single) return;
        try {
          onClipboardChange(copySelection(handle, single.selection));
          handle.deleteShape(single.selection.pageId, single.selection.shapeId);
          onMutation();
        } catch (error) { onError(error); }
      },
    },
    copy: {
      id: 'copy',
      enabled: copyable,
      run: () => {
        if (!handle || !single) return;
        try { onClipboardChange(copySelection(handle, single.selection)); } catch (error) { onError(error); }
      },
    },
    paste: {
      id: 'paste',
      enabled: Boolean(handle && clipboard && pageById(pages, pageId ?? selection[0]?.pageId ?? clipboard.pageId)),
      run: () => {
        if (!handle || !clipboard) return;
        try {
          const target = pageId ?? selection[0]?.pageId ?? clipboard.pageId;
          const step = clipboard.pasteCount + 1;
          const { receipt, entry } = pasteEntry(handle, target, clipboard, PASTE_OFFSET.x * step, PASTE_OFFSET.y * step);
          onClipboardChange(entry);
          onSelectShape({ pageId: target, shapeId: receipt.shapeId, hit: { kind: 'shape', shapeId: receipt.shapeId } });
          onMutation();
        } catch (error) { onError(error); }
      },
    },
    duplicate: {
      id: 'duplicate',
      enabled: copyable,
      run: () => {
        if (!handle || !single) return;
        try {
          const entry = copySelection(handle, single.selection);
          const receipt = duplicateEntry(handle, single.selection.pageId, entry, DUPLICATE_OFFSET.x, DUPLICATE_OFFSET.y);
          onSelectShape({ pageId: single.selection.pageId, shapeId: receipt.shapeId, hit: { kind: 'shape', shapeId: receipt.shapeId } });
          onMutation();
        } catch (error) { onError(error); }
      },
    },
    download: { id: 'download', enabled: Boolean(handle), run: () => { if (!handle) return; try { onDownload(handle.save()); } catch (error) { onError(error); } } },
    pageBreaks: { id: 'pageBreaks', enabled: Boolean(frame && pageBreaks), active: Boolean(pageBreaks?.shown), run: () => pageBreaks?.toggle() },
  } as RibbonCommands;
  return commands;
}

export function RibbonCommandsProvider({ handle, snapshot, pageId, selection, frame, clipboard = null, onClipboardChange = () => {}, onSelectShape = () => {}, onMutation, onError, onDownload, pageBreaks, probes, children }: RibbonCommandsProviderProps) {
  const commands = useMemo(() => createRibbonCommands(handle, selection, pageId, onMutation, onError, onDownload, frame, clipboard, onClipboardChange, onSelectShape, pageBreaks, probes), [handle, snapshot, pageId, selection, frame, clipboard, onClipboardChange, onSelectShape, onMutation, onError, onDownload, pageBreaks, probes]);
  return createElement(RibbonCommandsContext.Provider, { value: commands }, children);
}

export function useRibbonCommands(): RibbonCommands {
  const context = useContext(RibbonCommandsContext);
  if (!context) throw new Error('useRibbonCommands must be used within a <RibbonCommandsProvider>');
  return context;
}
