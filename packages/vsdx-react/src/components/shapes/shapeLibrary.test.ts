import { expect, test } from 'bun:test';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { arrowShapes, arrowVertices, polygonVertices, previewPathForVertices, shapeStencils, standardShapes, stencilShapeById } from './shapeLibrary';

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
  for (const shape of [...standardShapes, ...arrowShapes]) {
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
    { locator: { section: 'Connection', rowIndex: 0, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: 'Width*0.5' },
    { locator: { section: 'Connection', rowIndex: 0, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: 'Height*1' },
    { locator: { section: 'Connection', rowIndex: 1, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: 'Width*1' },
    { locator: { section: 'Connection', rowIndex: 1, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: 'Height*0.5' },
    { locator: { section: 'Connection', rowIndex: 2, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: 'Width*0.5' },
    { locator: { section: 'Connection', rowIndex: 2, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: 'Height*0' },
    { locator: { section: 'Connection', rowIndex: 3, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: 'Width*0' },
    { locator: { section: 'Connection', rowIndex: 3, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: 'Height*0.5' },
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
    { locator: { section: 'Connection', rowIndex: 0, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: 'Width*0.5' },
    { locator: { section: 'Connection', rowIndex: 0, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: 'Height*1' },
    { locator: { section: 'Connection', rowIndex: 1, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: 'Width*1' },
    { locator: { section: 'Connection', rowIndex: 1, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: 'Height*0.5' },
    { locator: { section: 'Connection', rowIndex: 2, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: 'Width*0.5' },
    { locator: { section: 'Connection', rowIndex: 2, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: 'Height*0' },
    { locator: { section: 'Connection', rowIndex: 3, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: 'Width*0' },
    { locator: { section: 'Connection', rowIndex: 3, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: 'Height*0.5' },
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
    expect(shape.preview.startsWith(previewPathForVertices(vertices, shape.defaultSize.width / shape.defaultSize.height))).toBe(true);
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

test('builds the cube as a 4:3 box whose inner edges meet at the front corner', () => {
  const cube = defaultSize('cube');
  expect(cube.width / cube.height).toBeCloseTo(4 / 3, 10);
  const corners = standardShapes.find((shape) => shape.id === 'cube')!.draft(0, 0, 1, 1).cells
    .filter((cell) => cell.locator.section === 'Geometry' && (cell.name === 'X' || cell.name === 'Y'))
    .map((cell) => cell.formula);
  expect(corners).toEqual([
    'Width*0', 'Height*0',
    'Width*0.75', 'Height*0',
    'Width*1', 'Height*0.25',
    'Width*1', 'Height*1',
    'Width*0.25', 'Height*1',
    'Width*0', 'Height*0.75',
    'Width*0.75', 'Height*0',
    'Width*0.75', 'Height*0.75',
    'Width*0', 'Height*0.75',
    'Width*0.75', 'Height*0.75',
    'Width*1', 'Height*1',
  ]);
});

function subpathAreas(shapeId: string): number[] {
  const cells = shapeStencils.flatMap((stencil) => stencil.shapes).find((shape) => shape.id === shapeId)!.draft(0, 0, 1, 1).cells
    .filter((cell) => cell.locator.section === 'Geometry' && (cell.name === 'X' || cell.name === 'Y'));
  const areas: number[] = [];
  let points: Array<[number, number]> = [];
  const close = () => {
    if (points.length > 2) areas.push(points.reduce((sum, [x, y], index) => {
      const [nextX, nextY] = points[(index + 1) % points.length];
      return sum + x * nextY - nextX * y;
    }, 0));
    points = [];
  };
  for (let index = 0; index < cells.length; index += 2) {
    if (cells[index].locator.rowType === 'MoveTo') close();
    points.push([Number(cells[index].formula!.split('*')[1]), Number(cells[index + 1].formula!.split('*')[1])]);
  }
  close();
  return areas;
}

test('winds the cube inner edges with its outline so a filled cube has no hole', () => {
  const [outline, ...inner] = subpathAreas('cube');
  expect(inner.length).toBeGreaterThan(0);
  for (const area of inner) expect(Math.sign(area)).toBe(Math.sign(outline));
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

test('derives every arrow polygon preview and geometry from shared vertices', () => {
  for (const [id, vertices] of Object.entries(arrowVertices)) {
    const shape = arrowShapes.find((candidate) => candidate.id === id)!;
    const cells = shape.draft(0, 0, 1, 1).cells.filter((cell) => cell.name === 'X' || cell.name === 'Y').slice(0, vertices.length * 2);
    expect(shape.preview.startsWith(previewPathForVertices(vertices, shape.defaultSize.width / shape.defaultSize.height))).toBe(true);
    expect(cells.map((cell) => cell.formula)).toEqual(vertices.flatMap(([x, y]) => [`Width*${x}`, `Height*${y}`]));
  }
});

test('scales every arrow coordinate off Width and Height', () => {
  for (const shape of arrowShapes) {
    for (const cell of shape.draft(0, 0, 1, 1).cells.filter((cell) => cell.locator.section === 'Geometry')) {
      if (cell.name === 'X') expect(cell.formula).toMatch(/^Width\*-?\d/);
      if (cell.name === 'Y') expect(cell.formula).toMatch(/^Height\*-?\d/);
    }
  }
});

test('gives every stencil shape a one-inch box, a distinct preview and a resolved name', () => {
  const shapes = shapeStencils.flatMap((stencil) => stencil.shapes);
  expect(new Set(shapes.map((shape) => shape.preview)).size).toBe(shapes.length);
  expect(new Set(shapes.map((shape) => shape.id)).size).toBe(shapes.length);
  for (const shape of shapes) {
    expect(shape.defaultSize.height).toBe(1);
    expect(t(shape.nameKey)).not.toBe(shape.nameKey);
  }
});

test('exposes two stencils covering every shape', () => {
  expect(shapeStencils.map((stencil) => stencil.id)).toEqual(['standard', 'arrows']);
  expect(shapeStencils[0].shapes).toEqual(standardShapes);
  expect(shapeStencils[1].shapes).toEqual(arrowShapes);
  expect(arrowShapes).toHaveLength(28);
});

test('follows the Visio arrow stencil order', () => {
  expect(arrowShapes.map((shape) => shape.id)).toEqual([
    'arrowRight', 'arrowLeft', 'arrowUp', 'arrowDown', 'arrowDoubleHorizontal', 'arrowDoubleVertical',
    'curvedArrowRight', 'curvedArrowLeft', 'curvedArrowUp', 'curvedArrowDown',
    'lineArrowRight', 'lineArrowLeft', 'lineArrowUp', 'lineArrowDown',
    'lineHorizontal', 'lineVertical', 'lineDiagonal', 'lineElbow',
    'bentArrow', 'uTurnArrow', 'sharpBent', 'stripedArrow', 'notched', 'blockArrow',
    'circularArrow', 'quadArrow', 'leftRightUp', 'arcedLine',
  ]);
});

test('resolves a dropped tile from either stencil', () => {
  expect(stencilShapeById('rectangle')?.id).toBe('rectangle');
  expect(stencilShapeById('curvedArrowRight')?.id).toBe('curvedArrowRight');
  expect(stencilShapeById('not a shape')).toBeUndefined();
});

test('gives every arrow arc a control point off its chord', () => {
  for (const shape of arrowShapes) {
    const cells = shape.draft(0, 0, 1, 1).cells.filter((cell) => cell.locator.section === 'Geometry');
    const at = (rowIndex: number, name: string) => Number((cells.find((cell) => cell.locator.rowIndex === rowIndex && cell.name === name)?.formula ?? '').split('*')[1]);
    let previous: [number, number] | undefined;
    for (const rowIndex of [...new Set(cells.map((cell) => cell.locator.rowIndex as number))].sort((left, right) => left - right)) {
      const rowType = cells.find((cell) => cell.locator.rowIndex === rowIndex)?.locator.rowType;
      const end: [number, number] = [at(rowIndex, 'X'), at(rowIndex, 'Y')];
      if (rowType === 'EllipticalArcTo') {
        const control: [number, number] = [at(rowIndex, 'A'), at(rowIndex, 'B')];
        expect(Number.isFinite(control[0]) && Number.isFinite(control[1])).toBe(true);
        expect(cells.some((cell) => cell.locator.rowIndex === rowIndex && cell.name === 'D')).toBe(true);
        const [startX, startY] = previous!;
        expect(Math.abs((end[0] - startX) * (control[1] - startY) - (end[1] - startY) * (control[0] - startX))).toBeGreaterThan(1e-6);
      }
      if (rowType !== 'Close') previous = end;
    }
  }
});

test('winds every filled subpath with its outline so no stencil shape has a hole', () => {
  for (const shape of shapeStencils.flatMap((stencil) => stencil.shapes)) {
    const [outline, ...inner] = subpathAreas(shape.id).filter((area) => Math.abs(area) > 1e-9);
    for (const area of inner) expect({ id: shape.id, sign: Math.sign(area) }).toEqual({ id: shape.id, sign: Math.sign(outline) });
  }
});
