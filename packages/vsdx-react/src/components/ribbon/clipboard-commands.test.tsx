import { expect, mock, test } from 'bun:test';
import type { DiagramHandle, DiagramSnapshot } from '@betteroffice/vsdx';
import { createRibbonCommands } from './commands';
import type { VsdxClipboardEntry } from './clipboard';

function snapshot(): DiagramSnapshot {
  return {
    pages: [{
      id: 'page',
      sourcePartPath: 'page',
      name: 'Page',
      shapes: [{
        id: 'one',
        sourceId: 1,
        name: 'Rect',
        children: [],
        cells: [
          { locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: 'PinX' }, name: 'PinX', formula: '1', value: '1' },
          { locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: 'PinY' }, name: 'PinY', formula: '2', value: '2' },
          { locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: 'FillForegnd' }, name: 'FillForegnd', formula: 'RGB(1,2,3)', value: null },
        ],
      }],
    }],
  };
}

function handle(state: DiagramSnapshot) {
  const value = {
    snapshot: () => state,
    canUndo: () => false,
    canRedo: () => false,
    undo: mock(() => ({})),
    redo: mock(() => ({})),
    deleteShape: mock(() => ({})),
    shapeText: mock(() => 'hello'),
    setShapeText: mock(() => ({})),
    addShape: mock(() => ({ shapeId: 'new' })),
    addShapeWithText: mock((pageId: string, _draft: unknown, text: string) => ({ pageId, shapeId: 'new', text })),
    addShapeTree: mock((pageId: string, _draft: unknown) => ({ pageId, shapeId: 'new-tree' })),
    subtreeGlue: mock(() => []),
  };
  return value as unknown as DiagramHandle & typeof value;
}

const selected = { pageId: 'page', shapeId: 'one', hit: { kind: 'shape' as const, shapeId: 'one' } };

test('copy stores cells and text without mutating', () => {
  const diagram = handle(snapshot());
  let clipboard: unknown = null;
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {}, undefined, null, (next) => { clipboard = next; });
  expect(commands.copy.enabled).toBe(true);
  commands.copy.run();
  expect((clipboard as { text: string }).text).toBe('hello');
  expect((clipboard as { cells: unknown[] }).cells).toHaveLength(3);
  expect(diagram.deleteShape).not.toHaveBeenCalled();
  expect(diagram.addShapeWithText).not.toHaveBeenCalled();
});

test('cut is copy plus a single delete', () => {
  const diagram = handle(snapshot());
  let clipboard: unknown = null;
  const refresh = mock(() => {});
  const commands = createRibbonCommands(diagram, [selected], 'page', refresh, () => {}, () => {}, undefined, null, (next) => { clipboard = next; });
  commands.cut.run();
  expect(clipboard).not.toBeNull();
  expect(diagram.deleteShape).toHaveBeenCalledTimes(1);
  expect(diagram.deleteShape).toHaveBeenCalledWith('page', 'one');
  expect(refresh).toHaveBeenCalledTimes(1);
});

test('paste offsets, selects the copy, and advances the paste count', () => {
  const diagram = handle(snapshot());
  let clipboard: VsdxClipboardEntry | null = null;
  const seen: unknown[] = [];
  const errors: unknown[] = [];
  const first = createRibbonCommands(diagram, [selected], 'page', () => {}, (error) => errors.push(error), () => {}, undefined, null, (next) => { clipboard = next; });
  first.copy.run();
  expect(errors).toEqual([]);
  const refresh = mock(() => {});
  const stored: VsdxClipboardEntry | null = clipboard;
  const second = createRibbonCommands(diagram, [selected], 'page', refresh, (error) => errors.push(error), () => {}, undefined, stored, (next) => { clipboard = next; }, (next) => seen.push(next));
  expect(second.paste.enabled).toBe(true);
  second.paste.run();
  expect(errors).toEqual([]);
  const calls = (diagram.addShapeWithText as unknown as { mock: { calls: unknown[][] } }).mock.calls;
  const draft = calls[0] as [string, { cells: Array<{ locator: { cellName: string }; formula?: string }> }, string];
  expect(draft[0]).toBe('page');
  expect(draft[2]).toBe('hello');
  expect(draft[1].cells.find((cell) => cell.locator.cellName === 'PinX')?.formula).toBe('1.25');
  expect(draft[1].cells.find((cell) => cell.locator.cellName === 'PinY')?.formula).toBe('1.75');
  expect((clipboard as VsdxClipboardEntry | null)?.pasteCount).toBe(1);
  expect(seen).toEqual([{ pageId: 'page', shapeId: 'new', hit: { kind: 'shape', shapeId: 'new' } }]);
  expect(refresh).toHaveBeenCalledTimes(1);
});

test('duplicate lands offset up-left without touching the clipboard', () => {
  const diagram = handle(snapshot());
  let clipboard: unknown = 'untouched';
  const seen: unknown[] = [];
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {}, undefined, null, (next) => { clipboard = next; }, (next) => seen.push(next));
  expect(commands.duplicate.enabled).toBe(true);
  commands.duplicate.run();
  expect(clipboard).toBe('untouched');
  const dupCalls = (diagram.addShapeWithText as unknown as { mock: { calls: unknown[][] } }).mock.calls;
  const draft = (dupCalls[0][1] as { cells: Array<{ locator: { cellName: string }; formula?: string }> });
  expect(draft.cells.find((cell) => cell.locator.cellName === 'PinX')?.formula).toBe('0.75');
  expect(draft.cells.find((cell) => cell.locator.cellName === 'PinY')?.formula).toBe('2.25');
  expect(seen).toHaveLength(1);
});

test('paste stays disabled without a clipboard entry', () => {
  const diagram = handle(snapshot());
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {}, undefined, null, () => {});
  expect(commands.paste.enabled).toBe(false);
  commands.paste.run();
  expect(diagram.addShapeWithText).not.toHaveBeenCalled();
});

test('enables cut, copy and duplicate for groups and pastes the whole tree', () => {
  const state = snapshot();
  const child = { ...state.pages[0].shapes[0], id: 'inner' };
  state.pages[0].shapes = [{ ...state.pages[0].shapes[0], id: 'group', children: [child] }];
  const diagram = handle(state);
  const errors: unknown[] = [];
  let clipboard: VsdxClipboardEntry | null = null;
  const grouped = { pageId: 'page', shapeId: 'group', hit: { kind: 'shape' as const, shapeId: 'group' } };
  const commands = createRibbonCommands(diagram, [grouped], 'page', () => {}, (error) => errors.push(error), () => {}, undefined, null, (next) => { clipboard = next; });
  expect(commands.copy.enabled).toBe(true);
  expect(commands.cut.enabled).toBe(true);
  expect(commands.duplicate.enabled).toBe(true);
  commands.copy.run();
  expect(errors).toEqual([]);
  expect((clipboard as VsdxClipboardEntry | null)?.children).toHaveLength(1);
  const stored = clipboard as VsdxClipboardEntry | null;
  const refresh = mock(() => {});
  const seen: unknown[] = [];
  const second = createRibbonCommands(diagram, [grouped], 'page', refresh, (error) => errors.push(error), () => {}, undefined, stored, (next) => { clipboard = next; }, (next) => seen.push(next));
  second.paste.run();
  expect(errors).toEqual([]);
  expect(diagram.addShapeTree).toHaveBeenCalledTimes(1);
  expect(seen).toEqual([{ pageId: 'page', shapeId: 'new-tree', hit: { kind: 'shape', shapeId: 'new-tree' } }]);
});

test('disables cut, copy and duplicate only for unportable content', () => {
  const state = snapshot();
  const child = { ...state.pages[0].shapes[0], id: 'inner' };
  state.pages[0].shapes = [{ ...state.pages[0].shapes[0], id: 'group', copyRefusal: 'embedded media', children: [child] }];
  const diagram = handle(state);
  const errors: unknown[] = [];
  let clipboard: unknown = 'untouched';
  const grouped = { pageId: 'page', shapeId: 'group', hit: { kind: 'shape' as const, shapeId: 'group' } };
  const commands = createRibbonCommands(diagram, [grouped], 'page', () => {}, (error) => errors.push(error), () => {}, undefined, null, (next) => { clipboard = next; });
  expect(commands.copy.enabled).toBe(false);
  expect(commands.cut.enabled).toBe(false);
  expect(commands.duplicate.enabled).toBe(false);
  commands.copy.run();
  expect(clipboard).toBe('untouched');
  expect(errors).toHaveLength(1);
  commands.cut.run();
  expect(diagram.deleteShape).not.toHaveBeenCalled();
  commands.duplicate.run();
  expect(diagram.addShapeWithText).not.toHaveBeenCalled();
  expect(diagram.addShapeTree).not.toHaveBeenCalled();
});
