import { expect, test } from 'bun:test';
import type { ShapeSnapshot } from '@betteroffice/vsdx';
import { DUPLICATE_OFFSET, PASTE_OFFSET, ancestorPinOffset, buildClipboardEntry, canCopyShape, draftForPaste, draftTreeForPaste, resolvedNumeric, toFormula } from './clipboard';

function shape(cells: Array<{ name: string; formula?: string | null; value?: string | null; section?: string | null; row?: { index: number } | { name: string } | null }>): ShapeSnapshot {
  return {
    id: 'shape:1',
    sourceId: 1,
    name: 'Rect',
    children: [],
    cells: cells.map((cell) => ({
      locator: { sheet: { page: 1 }, shapeId: 1, section: cell.section ?? null, sectionIndex: null, row: cell.row as never, cellName: cell.name },
      name: cell.name,
      formula: cell.formula ?? null,
      value: cell.value ?? null,
    })),
  };
}

test('copies every cell formula and reads the resolved pin', () => {
  const source = shape([
    { name: 'PinX', formula: '1', value: '1' },
    { name: 'PinY', formula: '2', value: '2' },
    { name: 'FillForegnd', formula: 'RGB(10,20,30)', value: null },
  ]);
  const entry = buildClipboardEntry('page:1', source, 'hello');
  expect(entry.text).toBe('hello');
  expect(entry.pinX).toBe(1);
  expect(entry.pinY).toBe(2);
  expect(entry.cells).toHaveLength(3);
  expect(entry.cells.find((cell) => cell.name === 'FillForegnd')?.formula).toBe('RGB(10,20,30)');
  expect(entry.cells.find((cell) => cell.name === 'PinX')).toMatchObject({ formula: '1', value: '1' });
});

test('paste offsets PinX and PinY without touching other formulas', () => {
  const source = shape([
    { name: 'PinX', formula: '4', value: '4' },
    { name: 'PinY', formula: '5', value: '5' },
    { name: 'Width', formula: 'Width', value: '1' },
  ]);
  const entry = buildClipboardEntry('page:1', source, '');
  const draft = draftForPaste(entry, PASTE_OFFSET.x, PASTE_OFFSET.y);
  const pinX = draft.cells.find((cell) => cell.locator.cellName === 'PinX');
  const pinY = draft.cells.find((cell) => cell.locator.cellName === 'PinY');
  expect(pinX?.formula).toBe('4.25');
  expect(pinX?.value).toBeUndefined();
  expect(pinY?.formula).toBe('4.75');
  expect(pinY?.value).toBeUndefined();
  expect(draft.cells.find((cell) => cell.locator.cellName === 'Width')?.formula).toBe('Width');
});

test('paste carries cached values so formula-less cells survive', () => {
  const source = shape([
    { name: 'PinX', formula: '1', value: '1' },
    { name: 'PinY', formula: '2', value: '2' },
    { name: 'Width', formula: null, value: '3.5' },
    { name: 'FillForegnd', formula: null, value: '1' },
  ]);
  const entry = buildClipboardEntry('page:1', source, '');
  const draft = draftForPaste(entry, PASTE_OFFSET.x, PASTE_OFFSET.y);
  const width = draft.cells.find((cell) => cell.locator.cellName === 'Width');
  expect(width?.value).toBe('3.5');
  expect(width?.formula).toBeUndefined();
  const fill = draft.cells.find((cell) => cell.locator.cellName === 'FillForegnd');
  expect(fill?.value).toBe('1');
  expect(fill?.formula).toBeUndefined();
});

test('duplicate offset lands up-left on screen', () => {
  expect(DUPLICATE_OFFSET.x).toBeLessThan(0);
  expect(DUPLICATE_OFFSET.y).toBeGreaterThan(0);
  expect(PASTE_OFFSET.x).toBeGreaterThan(0);
  expect(PASTE_OFFSET.y).toBeLessThan(0);
});

test('adds missing Pin cells so the copy never lands exactly on top', () => {
  const source = shape([{ name: 'Width', formula: '1', value: '1' }]);
  const entry = buildClipboardEntry('page:1', source, '');
  expect(entry.pinX).toBeNull();
  const draft = draftForPaste(entry, 0.25, -0.25);
  expect(draft.cells.find((cell) => cell.locator.cellName === 'PinX')?.formula).toBe('0.25');
  expect(draft.cells.find((cell) => cell.locator.cellName === 'PinY')?.formula).toBe('-0.25');
});

test('resolves non-literal pins through their cached value', () => {
  const source = shape([{ name: 'PinX', formula: 'Width/2', value: '3.5' }]);
  expect(resolvedNumeric(source, 'PinX')).toBe(3.5);
  expect(resolvedNumeric(source, 'Missing')).toBeNull();
  expect(toFormula(-0)).toBe('0');
});

test('copies groups with their children and offsets only the root', () => {
  const child = shape([{ name: 'PinX', formula: '1', value: '1' }]);
  const group: ShapeSnapshot = { id: 'group:1', sourceId: 1, name: 'Group', children: [{ ...child, id: 'child:1' }], cells: [{ locator: { sheet: { page: 1 }, shapeId: 1, section: null, sectionIndex: null, row: null, cellName: 'PinX' }, name: 'PinX', formula: '4', value: '4' }] };
  expect(canCopyShape(child)).toBe(true);
  expect(canCopyShape(group)).toBe(true);
  const entry = buildClipboardEntry('page:1', group, 'group text', { textFor: () => 'child text' });
  expect(entry.children).toHaveLength(1);
  expect(entry.children[0].text).toBe('child text');
  expect(entry.children[0].sourceShapeId).toBe('child:1');
  const tree = draftTreeForPaste(entry, PASTE_OFFSET.x, PASTE_OFFSET.y);
  expect(tree.text).toBe('group text');
  expect(tree.children?.[0]?.text).toBe('child text');
  expect(tree.cells.find((cell) => cell.locator.cellName === 'PinX')?.formula).toBe('4.25');
  expect(tree.children?.[0]?.cells.find((cell) => cell.locator.cellName === 'PinX')?.formula).toBe('1');
});

test('refuses only unportable content with its reason', () => {
  const child = shape([{ name: 'PinX', formula: '1', value: '1' }]);
  const blocked: ShapeSnapshot = { id: 'group:1', sourceId: 1, name: 'Group', copyRefusal: 'embedded media', children: [child], cells: [] };
  expect(canCopyShape(blocked)).toBe(false);
  expect(() => buildClipboardEntry('page:1', blocked, '')).toThrow('embedded media');
});

function group(id: string, cells: Array<{ name: string; value: string }>, children: ShapeSnapshot[]): ShapeSnapshot {
  return {
    id,
    sourceId: 10,
    name: 'Group',
    children,
    cells: cells.map((cell) => ({ locator: { sheet: { page: 1 }, shapeId: 10, section: null, sectionIndex: null, row: null, cellName: cell.name }, name: cell.name, formula: cell.value, value: cell.value })),
  };
}

test('a nested shape copies at the page position it is drawn at', () => {
  const child = { ...shape([{ name: 'PinX', formula: '1', value: '1' }, { name: 'PinY', formula: '2', value: '2' }]), id: 'child:1' };
  const parent = group('group:1', [{ name: 'PinX', value: '10' }, { name: 'PinY', value: '4' }, { name: 'LocPinX', value: '0.5' }, { name: 'LocPinY', value: '0.5' }], [child]);
  const offset = ancestorPinOffset([parent], 'child:1');
  expect(offset).toEqual({ dx: 9.5, dy: 3.5 });
  const entry = buildClipboardEntry('page:1', child, '', { pinOffset: offset ?? undefined });
  expect(entry.pinX).toBe(10.5);
  expect(entry.pinY).toBe(5.5);
  expect(ancestorPinOffset([parent], 'group:1')).toEqual({ dx: 0, dy: 0 });
});

test('a rotated or flipped ancestor refuses the copy instead of misplacing it', () => {
  const child = { ...shape([{ name: 'PinX', formula: '1', value: '1' }]), id: 'child:1' };
  const rotated = group('group:1', [{ name: 'PinX', value: '10' }, { name: 'PinY', value: '10' }, { name: 'Angle', value: String(Math.PI / 2) }], [child]);
  expect(ancestorPinOffset([rotated], 'child:1')).toBeNull();
  const flipped = group('group:1', [{ name: 'PinX', value: '10' }, { name: 'FlipX', value: '1' }], [child]);
  expect(ancestorPinOffset([flipped], 'child:1')).toBeNull();
  expect(ancestorPinOffset([rotated], 'absent')).toBeNull();
});
