import { expect, mock, test } from 'bun:test';
import type { DiagramHandle, DiagramSnapshot } from '@betteroffice/vsdx';
import { createRibbonCommands, findShapePlacement, numericCellValue } from './commands';

function snapshot(cells: Record<string, string> = {}): DiagramSnapshot {
  return { pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: ['one', 'two', 'three'].map((id) => ({ id, sourceId: 1, name: id, children: [], cells: Object.entries(cells).map(([name, value]) => ({ locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: name }, name, formula: value, value })) })) }] };
}

function handle(state: DiagramSnapshot, history = { undo: true, redo: false }) {
  const setCellFormula = mock((pageId: string, shapeId: string, locator: { cellName: string }, formula: string) => {
    const target = state.pages.find((page) => page.id === pageId)?.shapes.find((shape) => shape.id === shapeId);
    const targetCell = target?.cells.find((item) => item.locator.cellName === locator.cellName);
    if (targetCell) { targetCell.formula = formula; targetCell.value = formula; }
    return {};
  });
  const value = {
    snapshot: () => state, canUndo: () => history.undo, canRedo: () => history.redo,
    undo: mock(() => ({})), redo: mock(() => ({})), deleteShape: mock(() => ({})), setCellFormula, reorderShape: mock(() => ({})), addShape: mock(() => ({})), save: mock(() => new Uint8Array()),
  };
  return value as unknown as DiagramHandle & typeof value;
}

const selected = { pageId: 'page', shapeId: 'two', hit: { kind: 'shape' as const, shapeId: 'two' } };

test('exposes history from the handle and refreshes after mutations', () => {
  const state = snapshot(); const historyState = { undo: true, redo: false }; const diagram = handle(state, historyState); const refresh = mock(() => {});
  const commands = createRibbonCommands(diagram, null, 'page', refresh, () => {}, () => {});
  expect(commands.undo.enabled).toBe(true); expect(commands.redo.enabled).toBe(false); expect(commands.addShape.enabled).toBe(true); expect(commands.download.enabled).toBe(true); expect(commands.delete.enabled).toBe(false);
  commands.undo.run(); commands.addShape.run();
  expect(diagram.undo).toHaveBeenCalledTimes(1); expect(diagram.addShape).toHaveBeenCalledTimes(1); expect(refresh).toHaveBeenCalledTimes(2);
  historyState.undo = false; historyState.redo = true;
  const refreshed = createRibbonCommands(diagram, null, 'page', refresh, () => {}, () => {});
  expect(refreshed.undo.enabled).toBe(false); expect(refreshed.redo.enabled).toBe(true);
});

test('uses exact z-order bounds and ShapeSheet formulas', () => {
  const state = snapshot({ Angle: '0', FlipX: '0', FillForegnd: '#112233', LineColor: '#445566', LineWeight: '0.01 in', LinePattern: '4' }); const diagram = handle(state);
  const commands = createRibbonCommands(diagram, selected, 'page', () => {}, () => {}, () => {});
  commands.bringToFront.run(); commands.sendToBack.run(); commands.fillColor.run('#abcdef'); commands.lineColor.run('#fedcba'); commands.rotateRight.run(); commands.rotateRight.run(); commands.flipHorizontal.run();
  expect(diagram.reorderShape).toHaveBeenNthCalledWith(1, 'page', 'two', 2); expect(diagram.reorderShape).toHaveBeenNthCalledWith(2, 'page', 'two', 0);
  expect(diagram.setCellFormula).toHaveBeenCalledWith('page', 'two', { cellName: 'FillForegnd' }, 'RGB(171,205,239)'); expect(diagram.setCellFormula).toHaveBeenCalledWith('page', 'two', { cellName: 'LineColor' }, 'RGB(254,220,186)');
  expect(diagram.setCellFormula).toHaveBeenCalledWith('page', 'two', { cellName: 'Angle' }, String(Math.PI));
  expect(diagram.setCellFormula).toHaveBeenCalledWith('page', 'two', { cellName: 'FlipX' }, '1');
  expect(commands.fillColor.value).toBe('#112233'); expect(commands.lineColor.value).toBe('#445566'); expect(commands.lineWeight.value).toBe('0.01 in'); expect(commands.linePattern.value).toBe('4');
});

test('locks and guards disable the operations the mutation policy would refuse', () => {
  const state = snapshot({ LockDelete: '1', Angle: 'GUARD(0)', FlipX: 'GUARD(0)', FlipY: '0' });
  const diagram = handle(state);
  const commands = createRibbonCommands(diagram, selected, 'page', () => {}, () => {}, () => {});
  expect(commands.delete.enabled).toBe(false);
  expect(commands.rotateLeft.enabled).toBe(false);
  expect(commands.rotateRight.enabled).toBe(false);
  expect(commands.flipHorizontal.enabled).toBe(false);
  expect(commands.flipVertical.enabled).toBe(true);
  expect(commands.bringForward.enabled).toBe(true);
  expect(commands.sendBackward.enabled).toBe(true);
});

test('does not reorder forward past the topmost shape', () => {
  const state = snapshot(); const diagram = handle(state); const commands = createRibbonCommands(diagram, { ...selected, shapeId: 'three', hit: { kind: 'shape', shapeId: 'three' } }, 'page', () => {}, () => {}, () => {});
  expect(commands.bringForward.enabled).toBe(false); commands.bringForward.run(); expect(diagram.reorderShape).not.toHaveBeenCalled();
});

function cellsOf(pairs: Record<string, { formula?: string; value?: string }>) {
  return Object.entries(pairs).map(([name, entry]) => ({ locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: name }, name, formula: entry.formula ?? null, value: entry.value ?? null }));
}

function groupedSnapshot(): DiagramSnapshot {
  const child = (id: string) => ({ id, sourceId: 2, name: id, children: [], cells: cellsOf({ FillForegnd: { formula: '"#010203"', value: '#010203' } }) });
  return {
    pages: [{
      id: 'page',
      sourcePartPath: 'page',
      name: 'Page',
      shapes: [
        { id: 'loose', sourceId: 1, name: 'loose', children: [], cells: [] },
        { id: 'group', sourceId: 1, name: 'group', children: [child('inner-a'), child('inner-b'), child('inner-c')], cells: [] },
      ],
    }],
  };
}

test('finds a shape nested inside a group and reports its own siblings', () => {
  const placement = findShapePlacement(groupedSnapshot().pages[0].shapes, 'inner-b');
  expect(placement?.shape.id).toBe('inner-b');
  expect(placement?.index).toBe(1);
  expect(placement?.siblings).toHaveLength(3);
  expect(findShapePlacement(groupedSnapshot().pages[0].shapes, 'absent')).toBeNull();
});

test('enables and reorders group members against their own sibling order', () => {
  const state = groupedSnapshot();
  const diagram = handle(state);
  const nested = { pageId: 'page', shapeId: 'inner-a', hit: { kind: 'shape' as const, shapeId: 'inner-a' } };
  const commands = createRibbonCommands(diagram, nested, 'page', () => {}, () => {}, () => {});
  expect(commands.delete.enabled).toBe(true);
  expect(commands.fillColor.enabled).toBe(true);
  expect(commands.fillColor.value).toBe('#010203');
  expect(commands.sendBackward.enabled).toBe(false);
  expect(commands.bringToFront.enabled).toBe(true);
  commands.bringToFront.run();
  expect(diagram.reorderShape).toHaveBeenNthCalledWith(1, 'page', 'inner-a', 2);
});

test('offers the stored formula rather than the cached value for formula controls', () => {
  const state: DiagramSnapshot = {
    pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: [{ id: 'two', sourceId: 1, name: 'two', children: [], cells: cellsOf({ LineWeight: { formula: 'ThePage!LineWeight', value: '0.01 in' }, LinePattern: { formula: 'Sheet.5!LinePattern', value: '4' } }) }] }],
  };
  const diagram = handle(state);
  const commands = createRibbonCommands(diagram, selected, 'page', () => {}, () => {}, () => {});
  expect(commands.lineWeight.value).toBe('ThePage!LineWeight');
  expect(commands.linePattern.value).toBe('Sheet.5!LinePattern');
});

test('adds a rectangle carrying geometry rows instead of a bodiless shape', () => {
  const diagram = handle(snapshot());
  const commands = createRibbonCommands(diagram, null, 'page', () => {}, () => {}, () => {});
  commands.addShape.run();
  const draft = (diagram.addShape as unknown as { mock: { calls: unknown[][] } }).mock.calls[0][1] as { cells: Array<{ locator: { section?: string }; name: string; formula: string }> };
  expect(draft.cells.some((cell) => cell.locator.section === 'Geometry')).toBe(true);
  expect(Number(draft.cells.find((cell) => cell.name === 'Width')?.formula)).toBeCloseTo(4 / 3, 10);
});

test('refuses to add a shape onto a page that is no longer present', () => {
  const diagram = handle(snapshot());
  const errors: unknown[] = [];
  const commands = createRibbonCommands(diagram, null, 'missing-page', () => {}, (error) => errors.push(error), () => {});
  expect(commands.addShape.enabled).toBe(false);
  commands.addShape.run();
  expect(diagram.addShape).not.toHaveBeenCalled();
  expect(errors).toHaveLength(1);
});


test('does not mistake a prefix of an unresolved formula for a numeric angle', () => {
  const state = snapshot({ Angle: '2*ThePage!Angle' });
  state.pages[0].shapes[1].cells[0].value = null;
  const diagram = handle(state);
  const errors: unknown[] = [];
  const commands = createRibbonCommands(diagram, selected, 'page', () => {}, (error) => errors.push(error), () => {});
  commands.rotateRight.run();
  expect(diagram.setCellFormula).not.toHaveBeenCalled();
  expect(errors[0]).toEqual(new Error('Shape cell Angle has no resolved numeric value.'));
});

test('uses a shape root cell without confusing a same-named User cell', () => {
  const state = snapshot({ PinX: '4' });
  const shape = state.pages[0].shapes[1];
  shape.cells.unshift({ ...shape.cells[0], value: '99', formula: '99', locator: { ...shape.cells[0].locator, section: 'User', row: { index: 0 } } });
  expect(numericCellValue(shape, 'PinX')).toBe(4);
  shape.cells.pop();
  expect(() => numericCellValue(shape, 'PinX')).toThrow('Shape cell PinX has no resolved numeric value.');
});
