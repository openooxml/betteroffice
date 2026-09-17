import { beforeAll, expect, mock, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { DiagramHandle, DiagramSnapshot } from '@betteroffice/vsdx';
import { initWasm, openDiagram } from '@betteroffice/vsdx';
import { LINE_PATTERN_VALUES, createRibbonCommands, findShapePlacement, frameSwatch, isFormulaChange, isRotateBlocked, numericCellValue, parseLinePatternInput, parseLineWeightInput } from './commands';

function snapshot(cells: Record<string, string> = {}): DiagramSnapshot {
  return { pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: ['one', 'two', 'three'].map((id) => ({ id, sourceId: 1, name: id, children: [], cells: Object.entries(cells).map(([name, value]) => ({ locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: name }, name, formula: value, value })) })) }] };
}

function handle(state: DiagramSnapshot, history = { undo: true, redo: false }) {
  const applyWrite = (pageId: string, shapeId: string, cellName: string, formula: string) => {
    const target = state.pages.find((page) => page.id === pageId)?.shapes.find((shape) => shape.id === shapeId);
    const targetCell = target?.cells.find((item) => item.locator.cellName === cellName);
    if (targetCell) { targetCell.formula = formula; targetCell.value = formula; }
  };
  const setCellFormula = mock((pageId: string, shapeId: string, locator: { cellName: string }, formula: string) => {
    applyWrite(pageId, shapeId, locator.cellName, formula);
    return {};
  });
  const setCellFormulas = mock((writes: ReadonlyArray<{ pageId: string; shapeId: string; cellName: string; formula: string }>) => {
    for (const write of writes) applyWrite(write.pageId, write.shapeId, write.cellName, write.formula);
    return writes.map(() => ({}));
  });
  const deleteShapes = mock((deletes: ReadonlyArray<{ pageId: string; shapeId: string }>) => {
    return deletes.map(() => ({}));
  });
  const value = {
    snapshot: () => state, canUndo: () => history.undo, canRedo: () => history.redo,
    undo: mock(() => ({})), redo: mock(() => ({})), deleteShape: mock(() => ({})), deleteShapes, setCellFormula, setCellFormulas, reorderShape: mock(() => ({})), addShape: mock(() => ({})), save: mock(() => new Uint8Array()),
  };
  return value as unknown as DiagramHandle & typeof value;
}

const selected = { pageId: 'page', shapeId: 'two', hit: { kind: 'shape' as const, shapeId: 'two' } };

const repoRoot = resolve(import.meta.dir, '../../../../..');
let guardFixture: Uint8Array;

beforeAll(async () => {
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(import.meta.dir, '../../../../vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(repoRoot, 'crates/vsdx-parse/tests/fixtures/guard-format.vsdx')),
  ]);
  await initWasm(wasm);
  guardFixture = fixture;
});

test('exposes history from the handle and refreshes after mutations', () => {
  const state = snapshot(); const historyState = { undo: true, redo: false }; const diagram = handle(state, historyState); const refresh = mock(() => {});
  const commands = createRibbonCommands(diagram, [], 'page', refresh, () => {}, () => {});
  expect(commands.undo.enabled).toBe(true); expect(commands.redo.enabled).toBe(false); expect(commands.addShape.enabled).toBe(true); expect(commands.download.enabled).toBe(true); expect(commands.delete.enabled).toBe(false);
  commands.undo.run(); commands.addShape.run();
  expect(diagram.undo).toHaveBeenCalledTimes(1); expect(diagram.addShape).toHaveBeenCalledTimes(1); expect(refresh).toHaveBeenCalledTimes(2);
  historyState.undo = false; historyState.redo = true;
  const refreshed = createRibbonCommands(diagram, [], 'page', refresh, () => {}, () => {});
  expect(refreshed.undo.enabled).toBe(false); expect(refreshed.redo.enabled).toBe(true);
});

test('uses exact z-order bounds and ShapeSheet formulas', () => {
  const state = snapshot({ Angle: '0', FlipX: '0', FillForegnd: '#112233', LineColor: '#445566', LineWeight: '0.01 in', LinePattern: '4' }); const diagram = handle(state);
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {});
  commands.bringToFront.run(); commands.sendToBack.run(); commands.fillColor.run('#abcdef'); commands.lineColor.run('#fedcba'); commands.rotateRight.run(); commands.rotateRight.run(); commands.flipHorizontal.run();
  expect(diagram.reorderShape).toHaveBeenNthCalledWith(1, 'page', 'two', 2); expect(diagram.reorderShape).toHaveBeenNthCalledWith(2, 'page', 'two', 0);
  expect(diagram.setCellFormulas).toHaveBeenCalledWith([{ pageId: 'page', shapeId: 'two', cellName: 'FillForegnd', formula: 'RGB(171,205,239)' }]);
  expect(diagram.setCellFormulas).toHaveBeenCalledWith([{ pageId: 'page', shapeId: 'two', cellName: 'LineColor', formula: 'RGB(254,220,186)' }]);
  expect(diagram.setCellFormulas).toHaveBeenCalledWith([{ pageId: 'page', shapeId: 'two', cellName: 'Angle', formula: String(Math.PI) }]);
  expect(diagram.setCellFormulas).toHaveBeenCalledWith([{ pageId: 'page', shapeId: 'two', cellName: 'FlipX', formula: '1' }]);
  expect(commands.fillColor.value).toBe('#112233'); expect(commands.lineColor.value).toBe('#445566'); expect(commands.lineWeight.value).toBe('0.01 in'); expect(commands.linePattern.value).toBe('4');
});

test('locks and guards disable the operations the mutation policy would refuse', () => {
  const state = snapshot({ LockDelete: '1', Angle: 'GUARD(0)', FlipX: 'GUARD(0)', FlipY: '0' });
  const diagram = handle(state);
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {});
  expect(commands.delete.enabled).toBe(false);
  expect(commands.rotateLeft.enabled).toBe(false);
  expect(commands.rotateRight.enabled).toBe(false);
  expect(commands.flipHorizontal.enabled).toBe(false);
  expect(commands.flipVertical.enabled).toBe(true);
  expect(commands.bringForward.enabled).toBe(true);
  expect(commands.sendBackward.enabled).toBe(true);
});

test('a GUARD on a colour cell disables its picker instead of refusing on pick', () => {
  const state = snapshot({ FillForegnd: 'GUARD(RGB(255,0,0))', LineColor: 'RGB(0,0,255)' });
  const diagram = handle(state);
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {});
  expect(commands.fillColor.enabled).toBe(false);
  expect(commands.lineColor.enabled).toBe(true);
});

test('a GUARD substring inside a reference name disables nothing', () => {
  const state = snapshot({ LockDelete: 'User.GuardDelete', Angle: 'User.GuardAngle', FlipX: 'User.GuardFlip', FlipY: 'User.GuardFlip', FillForegnd: 'User.GuardFill', LineColor: 'User.GuardLine', LineWeight: 'User.GuardWeight', LinePattern: 'User.GuardPattern' });
  const diagram = handle(state);
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {});
  expect(commands.delete.enabled).toBe(true);
  expect(commands.rotateLeft.enabled).toBe(true);
  expect(commands.rotateRight.enabled).toBe(true);
  expect(commands.flipHorizontal.enabled).toBe(true);
  expect(commands.flipVertical.enabled).toBe(true);
  expect(commands.fillColor.enabled).toBe(true);
  expect(commands.lineColor.enabled).toBe(true);
  expect(commands.lineWeight.enabled).toBe(true);
  expect(commands.linePattern.enabled).toBe(true);
});

test('format commands follow the same guard predicate as rotate and flip', () => {
  const guarded = snapshot({ FillForegnd: 'GUARD(RGB(1,2,3))', LineColor: '=GUARD(RGB(4,5,6))', LineWeight: 'GUARD(0.01 in)', LinePattern: 'guard(1)', Angle: '0' });
  const blocked = createRibbonCommands(handle(guarded), [selected], 'page', () => {}, () => {}, () => {});
  expect(blocked.fillColor.enabled).toBe(false);
  expect(blocked.lineColor.enabled).toBe(false);
  expect(blocked.lineWeight.enabled).toBe(false);
  expect(blocked.linePattern.enabled).toBe(false);
  expect(blocked.rotateLeft.enabled).toBe(true);
  const plain = snapshot({ FillForegnd: 'RGB(1,2,3)', LineColor: 'RGB(4,5,6)', LineWeight: '0.01 in', LinePattern: '1', Angle: '0' });
  const open = createRibbonCommands(handle(plain), [selected], 'page', () => {}, () => {}, () => {});
  expect(open.fillColor.enabled).toBe(true);
  expect(open.lineColor.enabled).toBe(true);
  expect(open.lineWeight.enabled).toBe(true);
  expect(open.linePattern.enabled).toBe(true);
});

test('a SETATREF redirect to a guarded cell disables the control', () => {
  const redirected = snapshot({ FillForegnd: 'RGB(1,2,3)', LineColor: 'SETATREF(LineTarget)', LineTarget: 'GUARD(RGB(4,5,6))' });
  const blocked = createRibbonCommands(handle(redirected), [selected], 'page', () => {}, () => {}, () => {});
  expect(blocked.lineColor.enabled).toBe(false);
  expect(blocked.fillColor.enabled).toBe(true);
  const chained = snapshot({ LineColor: 'SETATREF(A)', A: 'SETATREF(B)', B: 'GUARD(RGB(4,5,6))' });
  expect(createRibbonCommands(handle(chained), [selected], 'page', () => {}, () => {}, () => {}).lineColor.enabled).toBe(false);
  const open = snapshot({ LineColor: 'SETATREF(LineTarget)', LineTarget: 'RGB(4,5,6)' });
  expect(createRibbonCommands(handle(open), [selected], 'page', () => {}, () => {}, () => {}).lineColor.enabled).toBe(true);
  const openChain = snapshot({ LineColor: '=SETATREF(A)', A: 'SETATREF(B)', B: 'RGB(4,5,6)' });
  expect(createRibbonCommands(handle(openChain), [selected], 'page', () => {}, () => {}, () => {}).lineColor.enabled).toBe(true);
});

test('an unresolvable SETATREF redirect disables the control the policy would refuse', () => {
  for (const cells of [
    { LineColor: 'SETATREF(Missing)' },
    { LineColor: 'SETATREF(Sheet.2!LineColor)' },
    { LineColor: 'SETATREF(LineColor)' },
    { LineColor: 'SETATREF(LineTarget)+1' },
    { LineColor: 'IF(1,SETATREF(LineTarget),0)', LineTarget: 'RGB(4,5,6)' },
  ] as Record<string, string>[]) {
    const commands = createRibbonCommands(handle(snapshot(cells)), [selected], 'page', () => {}, () => {}, () => {});
    expect(commands.lineColor.enabled).toBe(false);
  }
});

test('the guard fixture disables guarded controls end to end', () => {
  const diagram = openDiagram(guardFixture, { clientId: 7701 });
  try {
    const state = diagram.snapshot();
    const pageId = state.pages[0].id;
    const [guarded, redirected, plain] = state.pages[0].shapes.map((shape) => shape.id);
    const select = (shapeId: string) => ({ pageId, shapeId, hit: { kind: 'shape' as const, shapeId } });
    const guardedCommands = createRibbonCommands(diagram, [select(guarded)], pageId, () => {}, () => {}, () => {});
    expect(guardedCommands.fillColor.enabled).toBe(false);
    expect(guardedCommands.rotateLeft.enabled).toBe(false);
    expect(guardedCommands.rotateRight.enabled).toBe(false);
    expect(guardedCommands.delete.enabled).toBe(false);
    expect(guardedCommands.lineWeight.enabled).toBe(true);
    const redirectedCommands = createRibbonCommands(diagram, [select(redirected)], pageId, () => {}, () => {}, () => {});
    expect(redirectedCommands.lineColor.enabled).toBe(false);
    expect(redirectedCommands.fillColor.enabled).toBe(true);
    const plainCommands = createRibbonCommands(diagram, [select(plain)], pageId, () => {}, () => {}, () => {});
    expect(plainCommands.fillColor.enabled).toBe(true);
    expect(plainCommands.lineColor.enabled).toBe(true);
    expect(plainCommands.lineWeight.enabled).toBe(true);
    expect(plainCommands.linePattern.enabled).toBe(true);
    expect(plainCommands.rotateLeft.enabled).toBe(true);
    expect(plainCommands.delete.enabled).toBe(true);
    expect(() => diagram.setCellFormula(pageId, guarded, { cellName: 'FillForegnd' }, 'RGB(1,2,3)')).toThrow('GUARD protects the requested cell');
    expect(() => diagram.setCellFormula(pageId, redirected, { cellName: 'LineColor' }, 'RGB(1,2,3)')).toThrow('GUARD protects the requested cell');
  } finally {
    diagram.dispose();
  }
});

test('does not reorder forward past the topmost shape', () => {
  const state = snapshot(); const diagram = handle(state); const commands = createRibbonCommands(diagram, [{ ...selected, shapeId: 'three', hit: { kind: 'shape', shapeId: 'three' } }], 'page', () => {}, () => {}, () => {});
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
  const commands = createRibbonCommands(diagram, [nested], 'page', () => {}, () => {}, () => {});
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
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {});
  expect(commands.lineWeight.value).toBe('ThePage!LineWeight');
  expect(commands.linePattern.value).toBe('Sheet.5!LinePattern');
});

test('adds a rectangle carrying geometry rows instead of a bodiless shape', () => {
  const diagram = handle(snapshot());
  const commands = createRibbonCommands(diagram, [], 'page', () => {}, () => {}, () => {});
  commands.addShape.run();
  const draft = (diagram.addShape as unknown as { mock: { calls: unknown[][] } }).mock.calls[0][1] as { cells: Array<{ locator: { section?: string }; name: string; formula: string }> };
  expect(draft.cells.some((cell) => cell.locator.section === 'Geometry')).toBe(true);
  expect(Number(draft.cells.find((cell) => cell.name === 'Width')?.formula)).toBeCloseTo(4 / 3, 10);
});

test('refuses to add a shape onto a page that is no longer present', () => {
  const diagram = handle(snapshot());
  const errors: unknown[] = [];
  const commands = createRibbonCommands(diagram, [], 'missing-page', () => {}, (error) => errors.push(error), () => {});
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
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, (error) => errors.push(error), () => {});
  commands.rotateRight.run();
  expect(diagram.setCellFormulas).not.toHaveBeenCalled();
  expect(errors[0]).toEqual(new Error('Shape cell Angle has no resolved numeric value.'));
});

test('rejects non-positive and non-numeric line weights', () => {
  expect(parseLineWeightInput('-5')).toBeNull();
  expect(parseLineWeightInput('0')).toBeNull();
  expect(parseLineWeightInput('abc')).toBeNull();
  expect(parseLineWeightInput('')).toBeNull();
  expect(parseLineWeightInput('1e999')).toBeNull();
  expect(parseLineWeightInput('+5')).toBeNull();
  expect(parseLineWeightInput('0.018')).toBe('0.018');
  expect(parseLineWeightInput('0.01 in')).toBe('0.01 in');
  expect(parseLineWeightInput('12pt')).toBe('12 pt');
});

test('bounds line patterns to the documented 0..23 range', () => {
  expect(LINE_PATTERN_VALUES).toHaveLength(24);
  expect(parseLinePatternInput('999')).toBeNull();
  expect(parseLinePatternInput('abc')).toBeNull();
  expect(parseLinePatternInput('-1')).toBeNull();
  expect(parseLinePatternInput('4')).toBe('4');
  expect(parseLinePatternInput('0')).toBe('0');
  expect(parseLinePatternInput('23')).toBe('23');
});

test('invalid or unchanged ribbon values never reach the undo stack', () => {
  const state = snapshot({ LineWeight: '0.01', LinePattern: '1' });
  const diagram = handle(state);
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {});
  commands.lineWeight.run('-5');
  commands.lineWeight.run('abc');
  commands.lineWeight.run('0.01');
  commands.linePattern.run('999');
  commands.linePattern.run('abc');
  commands.linePattern.run('1');
  expect(diagram.setCellFormulas).not.toHaveBeenCalled();
  commands.lineWeight.run('0.05');
  commands.linePattern.run('4');
  expect(diagram.setCellFormulas).toHaveBeenCalledTimes(2);
  expect(isFormulaChange('0.01', '0.01', parseLineWeightInput)).toBe(false);
  expect(isFormulaChange('0.01 in', '0.01in', parseLineWeightInput)).toBe(false);
});

test('prefers rendered display-list colours over unresolved palette indexes', () => {
  const frame = {
    contractVersion: 4, width: 8, height: 8,
    paintTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    primitives: [{ kind: 'shape', id: 'page:2', zOrder: 0, path: [], fill: { kind: 'solid', color: '#0000FF' }, stroke: { color: '#00FF00', width: 1 } }],
  } as unknown as Parameters<typeof frameSwatch>[0];
  const state = snapshot({ FillForegnd: '5', LineColor: '7' });
  state.pages[0].shapes[1].sourceId = 2;
  state.pages[0].sourcePartPath = 'page';
  const diagram = handle(state);
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {}, frame);
  expect(commands.fillColor.value).toBe('#0000FF');
  expect(commands.lineColor.value).toBe('#00FF00');
  const withoutFrame = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {});
  expect(withoutFrame.fillColor.value).toBe('#000000');
});

test('a gradient fill shows its first stop instead of falling back to black', () => {
  const frame = {
    contractVersion: 5, width: 8, height: 8,
    paintTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    primitives: [{
      kind: 'shape', id: 'page:2', zOrder: 0, path: [],
      fill: { kind: 'gradient', angleDeg: 90, stops: [{ position: 0, color: '#123456' }, { position: 1, color: '#ABCDEF' }] },
    }],
  } as unknown as Parameters<typeof frameSwatch>[0];
  const state = snapshot({ FillForegnd: '5' });
  state.pages[0].shapes[1].sourceId = 2;
  state.pages[0].sourcePartPath = 'page';
  const commands = createRibbonCommands(handle(state), [selected], 'page', () => {}, () => {}, () => {}, frame);
  expect(commands.fillColor.value).toBe('#123456');
  expect(frameSwatch(frame, state.pages[0], state.pages[0].shapes[1]).line).toBeUndefined();
});

test('uses a shape root cell without confusing a same-named User cell', () => {
  const state = snapshot({ PinX: '4' });
  const shape = state.pages[0].shapes[1];
  shape.cells.unshift({ ...shape.cells[0], value: '99', formula: '99', locator: { ...shape.cells[0].locator, section: 'User', row: { index: 0 } } });
  expect(numericCellValue(shape, 'PinX')).toBe(4);
  shape.cells.pop();
  expect(() => numericCellValue(shape, 'PinX')).toThrow('Shape cell PinX has no resolved numeric value.');
});

test('deletes every selected shape while z-order stays single-selection', () => {
  const state = snapshot();
  const diagram = handle(state);
  const both = [{ ...selected, shapeId: 'one', hit: { kind: 'shape' as const, shapeId: 'one' } }, { ...selected, shapeId: 'two', hit: { kind: 'shape' as const, shapeId: 'two' } }];
  const commands = createRibbonCommands(diagram, both, 'page', () => {}, () => {}, () => {});
  expect(commands.delete.enabled).toBe(true);
  expect(commands.fillColor.enabled).toBe(true);
  expect(commands.bringToFront.enabled).toBe(false);
  expect(commands.sendBackward.enabled).toBe(false);
  commands.delete.run();
  expect(diagram.deleteShapes).toHaveBeenCalledTimes(1);
  expect(diagram.deleteShapes).toHaveBeenCalledWith([{ pageId: 'page', shapeId: 'one' }, { pageId: 'page', shapeId: 'two' }]);
});

test('a locked member disables delete for the whole selection', () => {
  const state = snapshot({ LockDelete: '1' });
  const diagram = handle(state);
  const both = [{ ...selected, shapeId: 'one', hit: { kind: 'shape' as const, shapeId: 'one' } }, { ...selected, shapeId: 'two', hit: { kind: 'shape' as const, shapeId: 'two' } }];
  const commands = createRibbonCommands(diagram, both, 'page', () => {}, () => {}, () => {});
  expect(commands.delete.enabled).toBe(false);
});

test('a rotation lock disables the rotate commands without touching the others', () => {
  const state = snapshot({ Angle: '0', LockRotate: '1' });
  const diagram = handle(state);
  const commands = createRibbonCommands(diagram, [selected], 'page', () => {}, () => {}, () => {});
  expect(commands.rotateLeft.enabled).toBe(false);
  expect(commands.rotateRight.enabled).toBe(false);
  expect(commands.flipHorizontal.enabled).toBe(true);
  expect(commands.delete.enabled).toBe(true);
  expect(isRotateBlocked(state.pages[0].shapes[1])).toBe(true);
  expect(isRotateBlocked(snapshot({ Angle: '0' }).pages[0].shapes[1])).toBe(false);
});
