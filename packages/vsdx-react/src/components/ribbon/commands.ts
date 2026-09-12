import { createContext, createElement, useContext, useMemo } from 'react';
import type { ReactNode } from 'react';
import type { DiagramHandle, DiagramSnapshot, PageSnapshot, ShapeSnapshot } from '@betteroffice/vsdx';
import type { VsdxShapeSelection } from '../../VsdxEditor';
import { standardShapeById } from '../shapes/shapeLibrary';

export type RibbonCommandId =
  | 'undo' | 'redo' | 'delete' | 'fillColor' | 'lineColor' | 'lineWeight' | 'linePattern'
  | 'bringToFront' | 'bringForward' | 'sendBackward' | 'sendToBack'
  | 'rotateLeft' | 'rotateRight' | 'flipHorizontal' | 'flipVertical' | 'addShape' | 'download';

export interface RibbonCommand { id: RibbonCommandId; run: (value?: string) => void; enabled: boolean; active?: boolean; value?: string; }
export type RibbonCommands = Record<RibbonCommandId, RibbonCommand>;

export const RibbonCommandsContext = createContext<RibbonCommands | null>(null);

export interface RibbonCommandsProviderProps {
  handle: DiagramHandle | null;
  snapshot: DiagramSnapshot | null;
  pageId?: string;
  selection: VsdxShapeSelection | null;
  onMutation: () => void;
  onError: (error: unknown) => void;
  onDownload: (bytes: Uint8Array) => void;
  children: ReactNode;
}

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

function placementIn(pages: readonly PageSnapshot[], selection: VsdxShapeSelection | null): ShapePlacement | null {
  if (!selection) return null;
  const page = pages.find((item) => item.id === selection.pageId);
  return page ? findShapePlacement(page.shapes, selection.shapeId) : null;
}

function findCell(shape: ShapeSnapshot | null, name: string) {
  return shape?.cells.find((item) => item.locator.cellName === name || item.name === name);
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
  const result = Number.parseFloat(value ?? '0');
  return Number.isFinite(result) ? result : 0;
}

export function createRibbonCommands(
  handle: DiagramHandle | null,
  selection: VsdxShapeSelection | null,
  pageId: string | undefined,
  onMutation: () => void,
  onError: (error: unknown) => void,
  onDownload: (bytes: Uint8Array) => void
): RibbonCommands {
  const execute = (operation: (current: DiagramHandle, selected: VsdxShapeSelection | null) => void, needsSelection = false) => () => {
    if (!handle || (needsSelection && !selection)) return;
    try { operation(handle, selection); onMutation(); } catch (error) { onError(error); }
  };
  const pages = handle ? handle.snapshot().pages : [];
  const current = placementIn(pages, selection);
  const shape = current?.shape ?? null;
  const selected = Boolean(current && selection);
  const topIndex = current ? current.siblings.length - 1 : 0;
  const livePlacement = (currentHandle: DiagramHandle, currentSelection: VsdxShapeSelection | null) => placementIn(currentHandle.snapshot().pages, currentSelection);
  const formula = (cellName: string, value: string) => execute((currentHandle, currentSelection) => {
    currentHandle.setCellFormula(currentSelection!.pageId, currentSelection!.shapeId, { cellName }, value);
  }, true);
  const reorderTo = (target: (placement: ShapePlacement) => number, allowed: (placement: ShapePlacement) => boolean) => execute((currentHandle, currentSelection) => {
    const placement = livePlacement(currentHandle, currentSelection);
    if (placement && allowed(placement)) currentHandle.reorderShape(currentSelection!.pageId, currentSelection!.shapeId, target(placement));
  }, true);
  const setNumeric = (cellName: string, next: (value: number) => string) => execute((currentHandle, currentSelection) => {
    const placement = livePlacement(currentHandle, currentSelection);
    if (!placement) return;
    currentHandle.setCellFormula(currentSelection!.pageId, currentSelection!.shapeId, { cellName }, next(numberValue(cellValue(placement.shape, cellName))));
  }, true);
  const commands = {
    undo: { id: 'undo', enabled: Boolean(handle?.canUndo()), run: execute((currentHandle) => { currentHandle.undo(); }) },
    redo: { id: 'redo', enabled: Boolean(handle?.canRedo()), run: execute((currentHandle) => { currentHandle.redo(); }) },
    delete: { id: 'delete', enabled: selected, run: execute((currentHandle, currentSelection) => { currentHandle.deleteShape(currentSelection!.pageId, currentSelection!.shapeId); }, true) },
    fillColor: { id: 'fillColor', enabled: selected, value: color(cellValue(shape, 'FillForegnd'), '#000000'), run: (value?: string) => formula('FillForegnd', colorFormula(value))() },
    lineColor: { id: 'lineColor', enabled: selected, value: color(cellValue(shape, 'LineColor'), '#000000'), run: (value?: string) => formula('LineColor', colorFormula(value))() },
    lineWeight: { id: 'lineWeight', enabled: selected, value: cellFormula(shape, 'LineWeight'), run: (value?: string) => { if (value) formula('LineWeight', value)(); } },
    linePattern: { id: 'linePattern', enabled: selected, value: cellFormula(shape, 'LinePattern'), run: (value?: string) => { if (value) formula('LinePattern', value)(); } },
    bringToFront: { id: 'bringToFront', enabled: selected && current!.index < topIndex, run: reorderTo((placement) => placement.siblings.length - 1, (placement) => placement.index < placement.siblings.length - 1) },
    bringForward: { id: 'bringForward', enabled: selected && current!.index < topIndex, run: reorderTo((placement) => placement.index + 1, (placement) => placement.index < placement.siblings.length - 1) },
    sendBackward: { id: 'sendBackward', enabled: selected && current!.index > 0, run: reorderTo((placement) => placement.index - 1, (placement) => placement.index > 0) },
    sendToBack: { id: 'sendToBack', enabled: selected && current!.index > 0, run: reorderTo(() => 0, (placement) => placement.index > 0) },
    rotateLeft: { id: 'rotateLeft', enabled: selected, run: setNumeric('Angle', (value) => String(value - Math.PI / 2)) },
    rotateRight: { id: 'rotateRight', enabled: selected, run: setNumeric('Angle', (value) => String(value + Math.PI / 2)) },
    flipHorizontal: { id: 'flipHorizontal', enabled: selected, active: numberValue(cellValue(shape, 'FlipX')) !== 0, run: setNumeric('FlipX', (value) => value === 0 ? '1' : '0') },
    flipVertical: { id: 'flipVertical', enabled: selected, active: numberValue(cellValue(shape, 'FlipY')) !== 0, run: setNumeric('FlipY', (value) => value === 0 ? '1' : '0') },
    addShape: {
      id: 'addShape',
      enabled: Boolean(pageById(pages, pageId)),
      run: execute((currentHandle) => {
        const page = pageById(currentHandle.snapshot().pages, pageId);
        if (!page) throw new Error(`vsdx page ${pageId ?? ''} is no longer part of the diagram`);
        const rectangle = standardShapeById('rectangle');
        if (!rectangle) throw new Error('vsdx standard rectangle shape is unavailable');
        currentHandle.addShape(page.id, rectangle.draft(1, 1, 1, 1));
      }),
    },
    download: { id: 'download', enabled: Boolean(handle), run: () => { if (!handle) return; try { onDownload(handle.save()); } catch (error) { onError(error); } } },
  } as RibbonCommands;
  return commands;
}

export function RibbonCommandsProvider({ handle, snapshot, pageId, selection, onMutation, onError, onDownload, children }: RibbonCommandsProviderProps) {
  const commands = useMemo(() => createRibbonCommands(handle, selection, pageId, onMutation, onError, onDownload), [handle, snapshot, pageId, selection, onMutation, onError, onDownload]);
  return createElement(RibbonCommandsContext.Provider, { value: commands }, children);
}

export function useRibbonCommands(): RibbonCommands {
  const context = useContext(RibbonCommandsContext);
  if (!context) throw new Error('useRibbonCommands must be used within a <RibbonCommandsProvider>');
  return context;
}
