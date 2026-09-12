import { expect, test } from 'bun:test';
import { polygonVertices, previewPathForVertices, standardShapes } from './shapeLibrary';

function geometry(shapeId: string) {
  return standardShapes.find((shape) => shape.id === shapeId)!.draft(2, 3, 4, 5).cells.filter((cell) => !['Angle', 'FlipX', 'FlipY', 'FillPattern', 'FillForegnd', 'LinePattern', 'LineColor', 'LineWeight'].includes(cell.locator.cellName));
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
    { locator: { cellName: 'Height' }, name: 'Height', formula: '5' },
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
    expect(shape.preview.startsWith(previewPathForVertices(vertices))).toBe(true);
    expect(cells.map((cell) => cell.formula)).toEqual(vertices.flatMap(([x, y]) => [`Width*${x}`, `Height*${y}`]));
  }
});
