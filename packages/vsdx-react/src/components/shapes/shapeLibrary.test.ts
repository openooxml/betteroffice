import { expect, test } from 'bun:test';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { polygonVertices, previewPathForVertices, standardShapes } from './shapeLibrary';

const t = createT(en);

function geometry(shapeId: string) {
  return standardShapes.find((shape) => shape.id === shapeId)!.draft(2, 3, 4, 5).cells.filter((cell) => !['Angle', 'FlipX', 'FlipY', 'FillPattern', 'FillForegnd', 'LinePattern', 'LineColor', 'LineWeight'].includes(cell.locator.cellName));
}

function defaultSize(shapeId: string) {
  const cells = standardShapes.find((shape) => shape.id === shapeId)!.draft(2, 3, 1, 1).cells;
  const width = Number(cells.find((cell) => cell.name === 'Width')?.formula);
  const heightFormula = cells.find((cell) => cell.name === 'Height')?.formula ?? '';
  const ratio = heightFormula.startsWith('Width/') ? Number(heightFormula.slice('Width/'.length)) : 0;
  const height = ratio > 0 ? width / ratio : heightFormula === 'Width' ? width : Number(heightFormula);
  return { width, height };
}

test('produces finite, complete formula-only drafts', () => {
  for (const shape of standardShapes) {
    for (const cell of shape.draft(Number.NaN, Number.POSITIVE_INFINITY, Number.NaN, Number.NEGATIVE_INFINITY).cells) {
      expect(cell.formula).toBeTruthy();
      expect(cell.formula).not.toMatch(/(?:nan|infinity)/i);
    }
  }
});

test('encodes the rectangle geometry cell by cell', () => {
  expect(geometry('rectangle')).toEqual([
    { locator: { cellName: 'PinX' }, name: 'PinX', formula: '2' },
    { locator: { cellName: 'PinY' }, name: 'PinY', formula: '3' },
    { locator: { cellName: 'Width' }, name: 'Width', formula: '4' },
    { locator: { cellName: 'Height' }, name: 'Height', formula: 'Width/1.333333333333' },
    { locator: { cellName: 'LocPinX' }, name: 'LocPinX', formula: 'Width*0.5' },
    { locator: { cellName: 'LocPinY' }, name: 'LocPinY', formula: 'Height*0.5' },
    { locator: { section: 'Geometry', rowIndex: 0, rowType: 'MoveTo', cellName: 'X' }, name: 'X', formula: 'Width*0' },
    { locator: { section: 'Geometry', rowIndex: 0, rowType: 'MoveTo', cellName: 'Y' }, name: 'Y', formula: 'Height*0' },
    { locator: { section: 'Geometry', rowIndex: 1, rowType: 'LineTo', cellName: 'X' }, name: 'X', formula: 'Width*1' },
    { locator: { section: 'Geometry', rowIndex: 1, rowType: 'LineTo', cellName: 'Y' }, name: 'Y', formula: 'Height*0' },
    { locator: { section: 'Geometry', rowIndex: 2, rowType: 'LineTo', cellName: 'X' }, name: 'X', formula: 'Width*1' },
    { locator: { section: 'Geometry', rowIndex: 2, rowType: 'LineTo', cellName: 'Y' }, name: 'Y', formula: 'Height*1' },
    { locator: { section: 'Geometry', rowIndex: 3, rowType: 'LineTo', cellName: 'X' }, name: 'X', formula: 'Width*0' },
    { locator: { section: 'Geometry', rowIndex: 3, rowType: 'LineTo', cellName: 'Y' }, name: 'Y', formula: 'Height*1' },
    { locator: { section: 'Geometry', rowIndex: 4, rowType: 'Close', cellName: 'NoShow' }, name: 'NoShow', formula: '0' },
  ]);
});

test('encodes the computed hexagon cell by cell', () => {
  const expected = polygonVertices.hexagon.flatMap(([x, y], index) => [
    { locator: { section: 'Geometry', rowIndex: index, rowType: index === 0 ? 'MoveTo' : 'LineTo', cellName: 'X' }, name: 'X', formula: `Width*${x}` },
    { locator: { section: 'Geometry', rowIndex: index, rowType: index === 0 ? 'MoveTo' : 'LineTo', cellName: 'Y' }, name: 'Y', formula: `Height*${y}` },
  ]);
  expect(geometry('hexagon').slice(6)).toEqual([
    ...expected,
    { locator: { section: 'Geometry', rowIndex: 6, rowType: 'Close', cellName: 'NoShow' }, name: 'NoShow', formula: '0' },
  ]);
});

test('constrains a square to equal dimensions', () => {
  const cells = standardShapes.find((shape) => shape.id === 'square')!.draft(2, 3, 4, 9).cells;
  expect(cells.find((cell) => cell.name === 'Width')?.formula).toBe('4');
  expect(cells.find((cell) => cell.name === 'Height')?.formula).toBe('Width');
});

test('derives every polygon preview and geometry from shared vertices', () => {
  for (const [id, vertices] of Object.entries(polygonVertices)) {
    const shape = standardShapes.find((candidate) => candidate.id === id)!;
    const cells = shape.draft(0, 0, 1, 1).cells.filter((cell) => cell.name === 'X' || cell.name === 'Y').slice(0, vertices.length * 2);
    if (id !== 'rectangle') expect(shape.preview.startsWith(previewPathForVertices(vertices))).toBe(true);
    expect(cells.map((cell) => cell.formula)).toEqual(vertices.flatMap(([x, y]) => [`Width*${x}`, `Height*${y}`]));
  }
});

test('renders the rectangle preview wider than tall', () => {
  expect(standardShapes.find((shape) => shape.id === 'rectangle')?.preview).toBe('M 0 0.875 L 1 0.875 L 1 0.125 L 0 0.125 Z');
});

test('gives every shape a distinct preview path', () => {
  const previews = standardShapes.map((shape) => shape.preview);
  expect(new Set(previews).size).toBe(standardShapes.length);
});

test('inserts the ellipse wider than tall and the circle square', () => {
  const ellipse = defaultSize('ellipse');
  expect(ellipse.width / ellipse.height).toBeCloseTo(1.5, 10);
  const circle = defaultSize('circle');
  expect(circle.width).toBe(circle.height);
});

test('gives every master a one-inch-tall default box at the ratio its draft writes', () => {
  for (const shape of standardShapes) {
    const drafted = defaultSize(shape.id);
    expect(shape.defaultSize.height).toBe(1);
    expect(shape.defaultSize.width / shape.defaultSize.height).toBeCloseTo(drafted.width / drafted.height, 10);
  }
  expect(standardShapes.find((shape) => shape.id === 'ellipse')?.defaultSize).toEqual({ width: 1.5, height: 1 });
  expect(standardShapes.find((shape) => shape.id === 'rectangle')?.defaultSize.width).toBeCloseTo(4 / 3, 10);
  expect(standardShapes.find((shape) => shape.id === 'circle')?.defaultSize).toEqual({ width: 1, height: 1 });
});

test('follows the Visio gallery order', () => {
  expect(standardShapes.map((shape) => shape.id)).toEqual([
    'rectangle', 'square', 'circle', 'ellipse', 'rightTriangle', 'triangle', 'rotatedTriangle',
    'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'cylinder', 'parallelogram',
    'trapezoid', 'diamond', 'cross', 'chevron', 'cube', 'teardrop', 'semicircle', 'halfEllipse',
    'cone', 'invertedCone', 'pyramid', 'pointedOval', 'funnel',
    'star4', 'star5', 'star6', 'star7', 'star16',
  ]);
});

test('resolves every shape name key', () => {
  for (const shape of standardShapes) expect(t(shape.nameKey)).not.toBe(shape.nameKey);
});
