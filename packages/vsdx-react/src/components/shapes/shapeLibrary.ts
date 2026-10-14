import type { FormulaShapeDraft } from '@betteroffice/vsdx';
import type { TranslationKey } from '@betteroffice/vsdx-i18n';

type Point = readonly [number, number];
type GeometryRowType = 'MoveTo' | 'LineTo' | 'EllipticalArcTo' | 'Close';

interface GeometryLocator {
  section: 'Geometry';
  rowIndex: number;
  rowType: GeometryRowType;
  cellName: string;
}

interface GeometryRow {
  type: GeometryRowType;
  end?: Point;
  through?: Point;
  radii?: Point;
  sweep?: number;
}

interface GeometryPath {
  rows: readonly GeometryRow[];
  preview: string;
}

export interface StandardShape {
  id: string;
  nameKey: TranslationKey;
  label?: string;
  preview: string;
  defaultSize: { width: number; height: number };
  draft: (x: number, y: number, width: number, height: number) => FormulaShapeDraft;
}

export interface ShapeStencil {
  id: string;
  nameKey: TranslationKey;
  shapes: readonly StandardShape[];
}

/** Height of every default insert box, in inches; the width follows the master's aspect ratio. */
const DEFAULT_SHAPE_HEIGHT_IN = 1;

export function shapeLabel(shape: StandardShape, t: (key: TranslationKey) => string): string {
  return shape.label ?? t(shape.nameKey);
}

export const polygonVertices: Readonly<Record<string, readonly Point[]>> = {
  rectangle: [[0, 0], [1, 0], [1, 1], [0, 1]],
  square: [[0, 0], [1, 0], [1, 1], [0, 1]],
  rightTriangle: [[0, 0], [1, 0], [0, 1]],
  triangle: [[0, 0], [1, 0], [0.5, 1]],
  rotatedTriangle: [[0, 1], [1, 1], [0.5, 0]],
  funnel: [[0, 1], [1, 1], [0.62, 0.45], [0.62, 0], [0.38, 0], [0.38, 0.45]],
  pyramid: [[0.5, 0.95], [0.9, 0.45], [0.75, 0.15], [0.15, 0.15]],
  star4: starVertices(4, 0.35),
  star5: starVertices(5, 0.382),
  star6: starVertices(6, 0.4),
  star7: starVertices(7, 0.45),
  star16: starVertices(16, 0.75),
  pentagon: regularPolygon(5),
  hexagon: regularPolygon(6),
  heptagon: regularPolygon(7),
  octagon: regularPolygon(8),
  decagon: regularPolygon(10),
  diamond: [[0.5, 1], [1, 0.5], [0.5, 0], [0, 0.5]],
  cross: [[0.3, 1], [0.7, 1], [0.7, 0.7], [1, 0.7], [1, 0.3], [0.7, 0.3], [0.7, 0], [0.3, 0], [0.3, 0.3], [0, 0.3], [0, 0.7], [0.3, 0.7]],
  chevron: [[0, 0.2], [0.38, 0.2], [0.62, 0], [1, 0.5], [0.62, 1], [0.38, 0.8], [0, 0.8], [0.35, 0.5]],
  parallelogram: [[0.2, 0], [1, 0], [0.8, 1], [0, 1]],
  trapezoid: [[0, 0], [1, 0], [0.8, 1], [0.2, 1]],
  cube: [[0, 0], [0.75, 0], [1, 0.25], [1, 1], [0.25, 1], [0, 0.75]],
};

function regularPolygon(sides: number): readonly Point[] {
  return Array.from({ length: sides }, (_, index) => {
    const angle = Math.PI / 2 + (Math.PI * 2 * index) / sides;
    return [cleanNumber(0.5 + 0.5 * Math.cos(angle)), cleanNumber(0.5 + 0.5 * Math.sin(angle))] as const;
  });
}

function starVertices(points: number, innerRatio: number): readonly Point[] {
  return Array.from({ length: points * 2 }, (_, index) => {
    const angle = Math.PI / 2 + (Math.PI * index) / points;
    const radius = index % 2 === 0 ? 0.5 : 0.5 * innerRatio;
    return [cleanNumber(0.5 + radius * Math.cos(angle)), cleanNumber(0.5 + radius * Math.sin(angle))] as const;
  });
}

function cleanNumber(value: number): number {
  return Number.isFinite(value) ? Number(value.toFixed(12)) : 0;
}

function numberFormula(value: number): string {
  const safe = cleanNumber(value);
  return Object.is(safe, -0) ? '0' : String(safe);
}

function dimension(value: number, fallback: number): number {
  return Math.max(0.01, Math.abs(Number.isFinite(value) ? value : fallback));
}

function coordinate(value: number): number {
  return Number.isFinite(value) ? value : 0;
}

function pointFormulas([x, y]: Point): [string, string] {
  return [`Width*${numberFormula(x)}`, `Height*${numberFormula(y)}`];
}

function polygonPath(vertices: readonly Point[], extraRows: readonly GeometryRow[] = []): GeometryPath {
  const rows: GeometryRow[] = [
    { type: 'MoveTo', end: vertices[0] },
    ...vertices.slice(1).map((end) => ({ type: 'LineTo' as const, end })),
    { type: 'Close' },
    ...extraRows,
  ];
  return { rows, preview: `${previewPathForVertices(vertices)}${svgRows(extraRows)}` };
}

export function previewPathForVertices(vertices: readonly Point[], aspectRatio = 1): string {
  return `${vertices.map(([x, y], index) => `${index === 0 ? 'M' : 'L'} ${numberFormula(x)} ${numberFormula(0.5 + (0.5 - y) / aspectRatio)}`).join(' ')} Z`;
}

function svgRows(rows: readonly GeometryRow[], aspectRatio = 1): string {
  const y = (value: number) => numberFormula(0.5 + (0.5 - value) / aspectRatio);
  return rows.map((row) => {
    if (row.type === 'Close') return ' Z';
    if (!row.end) return '';
    const end = `${numberFormula(row.end[0])} ${y(row.end[1])}`;
    if (row.type === 'EllipticalArcTo' && row.radii && row.sweep !== undefined) {
      return ` A ${numberFormula(row.radii[0])} ${numberFormula(row.radii[1] / aspectRatio)} 0 ${Math.abs(row.sweep) > Math.PI ? 1 : 0} ${row.sweep < 0 ? 1 : 0} ${end}`;
    }
    return ` ${row.type === 'MoveTo' ? 'M' : 'L'} ${end}`;
  }).join('');
}

function arc(center: Point, radii: Point, start: number, sweep: number): GeometryRow {
  const point = (angle: number): Point => [cleanNumber(center[0] + radii[0] * Math.cos(angle)), cleanNumber(center[1] + radii[1] * Math.sin(angle))];
  return { type: 'EllipticalArcTo', end: point(start + sweep), through: point(start + sweep / 2), radii, sweep };
}

function curvedPath(rows: readonly GeometryRow[]): GeometryPath {
  return { rows, preview: svgRows(rows).trim() };
}

function geometryCells(path: GeometryPath): FormulaShapeDraft['cells'] {
  return path.rows.flatMap((row, rowIndex) => {
    const locator = (cellName: string): GeometryLocator => ({ section: 'Geometry', rowIndex, rowType: row.type, cellName });
    if (row.type === 'Close') return [{ locator: locator('NoShow'), name: 'NoShow', formula: '0' }];
    if (!row.end) return [];
    const [x, y] = pointFormulas(row.end);
    const cells = [
      { locator: locator('X'), name: 'X', formula: x },
      { locator: locator('Y'), name: 'Y', formula: y },
    ];
    if (row.type !== 'EllipticalArcTo' || !row.through || !row.radii) return cells;
    const [a, b] = pointFormulas(row.through);
    return [
      ...cells,
      { locator: locator('A'), name: 'A', formula: a },
      { locator: locator('B'), name: 'B', formula: b },
      { locator: locator('C'), name: 'C', formula: '0' },
      { locator: locator('D'), name: 'D', formula: `Width*${numberFormula(row.radii[0])}/(Height*${numberFormula(row.radii[1])})` },
    ];
  });
}

function draftFor(id: string, path: GeometryPath, square: boolean, aspectRatio = 0) {
  return (x: number, y: number, width: number, height: number): FormulaShapeDraft => {
    const safeWidth = dimension(width, 1);
    const safeHeight = square ? safeWidth : aspectRatio > 0 ? safeWidth / aspectRatio : dimension(height, 1);
    return {
      name: id,
      cells: [
        { locator: { cellName: 'PinX' }, name: 'PinX', formula: numberFormula(coordinate(x)) },
        { locator: { cellName: 'PinY' }, name: 'PinY', formula: numberFormula(coordinate(y)) },
        { locator: { cellName: 'Width' }, name: 'Width', formula: numberFormula(safeWidth) },
        { locator: { cellName: 'Height' }, name: 'Height', formula: square ? 'Width' : aspectRatio > 0 ? `Width/${numberFormula(aspectRatio)}` : numberFormula(safeHeight) },
        { locator: { cellName: 'LocPinX' }, name: 'LocPinX', formula: 'Width*0.5' },
        { locator: { cellName: 'LocPinY' }, name: 'LocPinY', formula: 'Height*0.5' },
        ...geometryCells(path),
        ...Object.entries({ Angle: '0', FlipX: '0', FlipY: '0', FillPattern: '1', FillForegnd: 'RGB(255,255,255)', LinePattern: '1', LineColor: 'RGB(23,32,51)', LineWeight: '0.01' }).map(([name, formula]) => ({ locator: { cellName: name }, name, formula })),
        ...connectionCells(),
      ],
    };
  };
}

/** Connection rows in the order the connector overlay reads them: north, east, south, west. */
const connectionSides: readonly Point[] = [[0.5, 1], [1, 0.5], [0.5, 0], [0, 0.5]];

function connectionCells(): FormulaShapeDraft['cells'] {
  return connectionSides.flatMap(([x, y], rowIndex) => {
    const [xFormula, yFormula] = pointFormulas([x, y]);
    return [
      { locator: { section: 'Connection', rowIndex, rowType: 'Connection', cellName: 'X' }, name: 'X', formula: xFormula },
      { locator: { section: 'Connection', rowIndex, rowType: 'Connection', cellName: 'Y' }, name: 'Y', formula: yFormula },
    ];
  });
}

function shapeFrom(id: string, path: GeometryPath, square = false, aspectRatio = 0): StandardShape {
  return {
    id,
    nameKey: `shapesPanel.shape.${id}` as TranslationKey,
    preview: aspectRatio ? svgRows(path.rows, aspectRatio).trim() : path.preview,
    defaultSize: { width: aspectRatio > 0 ? DEFAULT_SHAPE_HEIGHT_IN * aspectRatio : DEFAULT_SHAPE_HEIGHT_IN, height: DEFAULT_SHAPE_HEIGHT_IN },
    draft: draftFor(id, path, square, aspectRatio),
  };
}

function polygonShape(id: keyof typeof polygonVertices, extraRows: readonly GeometryRow[] = [], square = false, aspectRatio = 0): StandardShape {
  return shapeFrom(id, polygonPath(polygonVertices[id], extraRows), square, aspectRatio);
}

const circlePath = curvedPath([
  { type: 'MoveTo', end: [1, 0.5] },
  arc([0.5, 0.5], [0.5, 0.5], 0, Math.PI),
  arc([0.5, 0.5], [0.5, 0.5], Math.PI, Math.PI),
  { type: 'Close' },
]);

const cylinderPath = curvedPath([
  { type: 'MoveTo', end: [0, 0.72] },
  arc([0.5, 0.72], [0.5, 0.28], Math.PI, -Math.PI),
  { type: 'LineTo', end: [1, 0.28] },
  arc([0.5, 0.28], [0.5, 0.28], 0, -Math.PI),
  { type: 'LineTo', end: [0, 0.72] },
  { type: 'Close' },
  { type: 'MoveTo', end: [0, 0.72] },
  arc([0.5, 0.72], [0.5, 0.28], Math.PI, Math.PI),
]);

const semicirclePath = curvedPath([
  { type: 'MoveTo', end: [0, 0] },
  { type: 'LineTo', end: [1, 0] },
  arc([0.5, 0], [0.5, 0.5], 0, Math.PI),
  { type: 'Close' },
]);

const halfEllipsePath = curvedPath([
  { type: 'MoveTo', end: [0, 0] },
  { type: 'LineTo', end: [1, 0] },
  arc([0.5, 0], [0.5, 0.35], 0, Math.PI),
  { type: 'Close' },
]);

const teardropRadius = 0.32;
const teardropTangent = Math.asin(teardropRadius / (1 - teardropRadius));
const teardropPath = curvedPath([
  { type: 'MoveTo', end: [0.5, 1] },
  { type: 'LineTo', end: [0.5 + teardropRadius * Math.cos(teardropTangent), teardropRadius * (1 + Math.sin(teardropTangent))] },
  arc([0.5, teardropRadius], [teardropRadius, teardropRadius], teardropTangent, -Math.PI - 2 * teardropTangent),
  { type: 'LineTo', end: [0.5, 1] },
  { type: 'Close' },
]);

const ovalRadius = 0.5125;
const ovalAngle = Math.atan2(0.5, 0.1125);
const pointedOvalPath = curvedPath([
  { type: 'MoveTo', end: [0.5, 1] },
  arc([0.3875, 0.5], [ovalRadius, ovalRadius], ovalAngle, -2 * ovalAngle),
  arc([0.6125, 0.5], [ovalRadius, ovalRadius], Math.PI + ovalAngle, -2 * ovalAngle),
  { type: 'Close' },
]);

function conePath(inverted: boolean): GeometryPath {
  const base = inverted ? 0.78 : 0.22;
  return curvedPath([
    { type: 'MoveTo', end: [0.5, inverted ? 0 : 1] },
    { type: 'LineTo', end: [0, base] },
    arc([0.5, base], [0.5, 0.14], Math.PI, Math.PI),
    { type: 'Close' },
    { type: 'MoveTo', end: [0, base] },
    arc([0.5, base], [0.5, 0.14], Math.PI, -Math.PI),
  ]);
}
const uprightConePath = conePath(false);
const invertedConePath = conePath(true);

const cubeEdges: readonly GeometryRow[] = [
  { type: 'MoveTo', end: [0.75, 0] },
  { type: 'LineTo', end: [0.75, 0.75] },
  { type: 'LineTo', end: [0, 0.75] },
  { type: 'MoveTo', end: [0.75, 0.75] },
  { type: 'LineTo', end: [1, 1] },
];

const pyramidEdges: readonly GeometryRow[] = [
  { type: 'MoveTo', end: [0.15, 0.15] },
  { type: 'LineTo', end: [0.3, 0.45] },
  { type: 'LineTo', end: [0.9, 0.45] },
  { type: 'MoveTo', end: [0.5, 0.95] },
  { type: 'LineTo', end: [0.75, 0.15] },
  { type: 'MoveTo', end: [0.5, 0.95] },
  { type: 'LineTo', end: [0.3, 0.45] },
];

export const standardShapes: readonly StandardShape[] = [
  polygonShape('rectangle', [], false, 4 / 3),
  polygonShape('square', [], true),
  shapeFrom('circle', circlePath, true),
  shapeFrom('ellipse', circlePath, false, 1.5),
  polygonShape('rightTriangle'),
  polygonShape('triangle'),
  polygonShape('rotatedTriangle'),
  polygonShape('pentagon'),
  polygonShape('hexagon'),
  polygonShape('heptagon'),
  polygonShape('octagon'),
  polygonShape('decagon'),
  shapeFrom('cylinder', cylinderPath),
  polygonShape('parallelogram'),
  polygonShape('trapezoid'),
  polygonShape('diamond'),
  polygonShape('cross'),
  polygonShape('chevron'),
  polygonShape('cube', cubeEdges, false, 4 / 3),
  shapeFrom('teardrop', teardropPath),
  shapeFrom('semicircle', semicirclePath),
  shapeFrom('halfEllipse', halfEllipsePath),
  shapeFrom('cone', uprightConePath),
  shapeFrom('invertedCone', invertedConePath),
  polygonShape('pyramid', pyramidEdges),
  shapeFrom('pointedOval', pointedOvalPath),
  polygonShape('funnel'),
  polygonShape('star4'),
  polygonShape('star5'),
  polygonShape('star6'),
  polygonShape('star7'),
  polygonShape('star16'),
];

export const arrowVertices: Readonly<Record<string, readonly Point[]>> = {
  arrowRight: [[0, 0.35], [0.6, 0.35], [0.6, 0.15], [1, 0.5], [0.6, 0.85], [0.6, 0.65], [0, 0.65]],
  arrowLeft: [[1, 0.35], [0.4, 0.35], [0.4, 0.15], [0, 0.5], [0.4, 0.85], [0.4, 0.65], [1, 0.65]],
  arrowUp: [[0.35, 0], [0.35, 0.6], [0.15, 0.6], [0.5, 1], [0.85, 0.6], [0.65, 0.6], [0.65, 0]],
  arrowDown: [[0.35, 1], [0.35, 0.4], [0.15, 0.4], [0.5, 0], [0.85, 0.4], [0.65, 0.4], [0.65, 1]],
  arrowDoubleHorizontal: [[0, 0.5], [0.2, 0.15], [0.2, 0.35], [0.8, 0.35], [0.8, 0.15], [1, 0.5], [0.8, 0.85], [0.8, 0.65], [0.2, 0.65], [0.2, 0.85]],
  arrowDoubleVertical: [[0.5, 0], [0.15, 0.2], [0.35, 0.2], [0.35, 0.8], [0.15, 0.8], [0.5, 1], [0.85, 0.8], [0.65, 0.8], [0.65, 0.2], [0.85, 0.2]],
  sharpBent: [[0, 0.25], [0.7, 0.25], [0.7, 0.62], [0.85, 0.62], [0.6, 1], [0.35, 0.62], [0.5, 0.62], [0.5, 0.45], [0, 0.45]],
  stripedArrow: [[0.45, 0.15], [1, 0.5], [0.45, 0.85]],
  notched: [[0.25, 0.35], [0.6, 0.35], [0.6, 0.15], [1, 0.5], [0.6, 0.85], [0.6, 0.65], [0.25, 0.65], [0.4, 0.5]],
  blockArrow: [[0, 0.3], [0.55, 0.3], [0.55, 0.1], [1, 0.5], [0.55, 0.9], [0.55, 0.7], [0, 0.7]],
  quadArrow: [[0.5, 1], [0.7, 0.76], [0.6, 0.76], [0.6, 0.6], [0.76, 0.6], [0.76, 0.7], [1, 0.5], [0.76, 0.3], [0.76, 0.4], [0.6, 0.4], [0.6, 0.24], [0.7, 0.24], [0.5, 0], [0.3, 0.24], [0.4, 0.24], [0.4, 0.4], [0.24, 0.4], [0.24, 0.3], [0, 0.5], [0.24, 0.7], [0.24, 0.6], [0.4, 0.6], [0.4, 0.76], [0.3, 0.76]],
  leftRightUp: [[0.38, 0], [0.38, 0.38], [0.12, 0.38], [0, 0.5], [0.12, 0.62], [0.38, 0.62], [0.38, 0.7], [0.3, 0.7], [0.5, 1], [0.7, 0.7], [0.62, 0.7], [0.62, 0.62], [0.88, 0.62], [1, 0.5], [0.88, 0.38], [0.62, 0.38], [0.62, 0]],
};

function arrowShape(id: keyof typeof arrowVertices, extraRows: readonly GeometryRow[] = []): StandardShape {
  return shapeFrom(id, polygonPath(arrowVertices[id], extraRows));
}

function arcPoint(center: Point, radii: Point, angle: number): Point {
  return [cleanNumber(center[0] + radii[0] * Math.cos(angle)), cleanNumber(center[1] + radii[1] * Math.sin(angle))];
}

const curvedArrowRadii: Point = [0.32, 0.32];
const curvedArrowSweep = (140 * Math.PI) / 180;

function curvedArrowPath(center: Point, start: number, sweep: number, head: readonly [Point, Point, Point]): GeometryPath {
  return curvedPath([
    { type: 'MoveTo', end: arcPoint(center, curvedArrowRadii, start) },
    arc(center, curvedArrowRadii, start, sweep),
    { type: 'MoveTo', end: head[0] },
    { type: 'LineTo', end: head[1] },
    { type: 'LineTo', end: head[2] },
    { type: 'Close' },
  ]);
}

const degrees = (value: number) => (value * Math.PI) / 180;

const curvedArrowRightPath = curvedArrowPath([0.35, 0.2], degrees(160), -curvedArrowSweep, [[0.9, 0.31], [0.65, 0.43], [0.65, 0.19]]);
const curvedArrowLeftPath = curvedArrowPath([0.65, 0.2], degrees(20), curvedArrowSweep, [[0.1, 0.31], [0.35, 0.19], [0.35, 0.43]]);
const curvedArrowUpPath = curvedArrowPath([0.2, 0.35], degrees(-70), curvedArrowSweep, [[0.31, 0.9], [0.19, 0.65], [0.43, 0.65]]);
const curvedArrowDownPath = curvedArrowPath([0.2, 0.65], degrees(70), -curvedArrowSweep, [[0.31, 0.1], [0.43, 0.35], [0.19, 0.35]]);

const lineArrowRightPath = curvedPath([
  { type: 'MoveTo', end: [0, 0.5] },
  { type: 'LineTo', end: [0.76, 0.5] },
  { type: 'MoveTo', end: [1, 0.5] },
  { type: 'LineTo', end: [0.76, 0.62] },
  { type: 'LineTo', end: [0.76, 0.38] },
  { type: 'Close' },
]);

const lineArrowLeftPath = curvedPath([
  { type: 'MoveTo', end: [1, 0.5] },
  { type: 'LineTo', end: [0.24, 0.5] },
  { type: 'MoveTo', end: [0, 0.5] },
  { type: 'LineTo', end: [0.24, 0.38] },
  { type: 'LineTo', end: [0.24, 0.62] },
  { type: 'Close' },
]);

const lineArrowUpPath = curvedPath([
  { type: 'MoveTo', end: [0.5, 0] },
  { type: 'LineTo', end: [0.5, 0.76] },
  { type: 'MoveTo', end: [0.5, 1] },
  { type: 'LineTo', end: [0.38, 0.76] },
  { type: 'LineTo', end: [0.62, 0.76] },
  { type: 'Close' },
]);

const lineArrowDownPath = curvedPath([
  { type: 'MoveTo', end: [0.5, 1] },
  { type: 'LineTo', end: [0.5, 0.24] },
  { type: 'MoveTo', end: [0.5, 0] },
  { type: 'LineTo', end: [0.62, 0.24] },
  { type: 'LineTo', end: [0.38, 0.24] },
  { type: 'Close' },
]);

const lineHorizontalPath = curvedPath([
  { type: 'MoveTo', end: [0, 0.5] },
  { type: 'LineTo', end: [1, 0.5] },
]);

const lineVerticalPath = curvedPath([
  { type: 'MoveTo', end: [0.5, 0] },
  { type: 'LineTo', end: [0.5, 1] },
]);

const lineDiagonalPath = curvedPath([
  { type: 'MoveTo', end: [0.05, 0.05] },
  { type: 'LineTo', end: [0.95, 0.95] },
]);

const lineElbowPath = curvedPath([
  { type: 'MoveTo', end: [0.05, 0.7] },
  { type: 'LineTo', end: [0.55, 0.7] },
  { type: 'LineTo', end: [0.55, 0.3] },
]);

const stripedArrowStripes: readonly GeometryRow[] = [
  { type: 'MoveTo', end: [0, 0.32] },
  { type: 'LineTo', end: [0.42, 0.32] },
  { type: 'MoveTo', end: [0, 0.5] },
  { type: 'LineTo', end: [0.42, 0.5] },
  { type: 'MoveTo', end: [0, 0.68] },
  { type: 'LineTo', end: [0.42, 0.68] },
];

const bentArrowPath = curvedPath([
  { type: 'MoveTo', end: [0, 0.25] },
  { type: 'LineTo', end: [0.5, 0.25] },
  arc([0.5, 0.45], [0.2, 0.2], degrees(-90), degrees(90)),
  { type: 'LineTo', end: [0.7, 0.62] },
  { type: 'LineTo', end: [0.85, 0.62] },
  { type: 'LineTo', end: [0.6, 1] },
  { type: 'LineTo', end: [0.35, 0.62] },
  { type: 'LineTo', end: [0.5, 0.62] },
  { type: 'LineTo', end: [0.5, 0.45] },
  { type: 'LineTo', end: [0, 0.45] },
  { type: 'Close' },
]);

const uTurnArrowPath = curvedPath([
  { type: 'MoveTo', end: [0, 0.1] },
  { type: 'LineTo', end: [0.6, 0.1] },
  arc([0.6, 0.5], [0.4, 0.4], degrees(-90), degrees(90)),
  arc([0.6, 0.5], [0.4, 0.4], 0, degrees(90)),
  { type: 'LineTo', end: [0.3, 0.9] },
  { type: 'LineTo', end: [0.3, 0.98] },
  { type: 'LineTo', end: [0, 0.8] },
  { type: 'LineTo', end: [0.3, 0.62] },
  { type: 'LineTo', end: [0.3, 0.7] },
  { type: 'LineTo', end: [0.6, 0.7] },
  arc([0.6, 0.5], [0.2, 0.2], degrees(90), degrees(-90)),
  arc([0.6, 0.5], [0.2, 0.2], 0, degrees(-90)),
  { type: 'LineTo', end: [0, 0.3] },
  { type: 'Close' },
]);

const circularArrowCenter: Point = [0.5, 0.52];
const circularArrowPath = curvedPath([
  { type: 'MoveTo', end: arcPoint(circularArrowCenter, [0.44, 0.44], degrees(-60)) },
  arc(circularArrowCenter, [0.44, 0.44], degrees(-60), degrees(100)),
  arc(circularArrowCenter, [0.44, 0.44], degrees(40), degrees(100)),
  arc(circularArrowCenter, [0.44, 0.44], degrees(140), degrees(100)),
  { type: 'LineTo', end: [0.409903810568, 0.063948822335] },
  { type: 'LineTo', end: arcPoint(circularArrowCenter, [0.24, 0.24], degrees(-120)) },
  arc(circularArrowCenter, [0.24, 0.24], degrees(-120), degrees(-100)),
  arc(circularArrowCenter, [0.24, 0.24], degrees(140), degrees(-100)),
  arc(circularArrowCenter, [0.24, 0.24], degrees(40), degrees(-100)),
  { type: 'Close' },
]);

const arcedLinePath = curvedPath([
  { type: 'MoveTo', end: [0.2, 0.9] },
  arc([0.2, 0.3], [0.6, 0.6], degrees(90), degrees(-90)),
  { type: 'MoveTo', end: [0.8, 0.16] },
  { type: 'LineTo', end: [0.72, 0.3] },
  { type: 'LineTo', end: [0.88, 0.3] },
  { type: 'Close' },
]);

export const arrowShapes: readonly StandardShape[] = [
  arrowShape('arrowRight'),
  arrowShape('arrowLeft'),
  arrowShape('arrowUp'),
  arrowShape('arrowDown'),
  arrowShape('arrowDoubleHorizontal'),
  arrowShape('arrowDoubleVertical'),
  shapeFrom('curvedArrowRight', curvedArrowRightPath),
  shapeFrom('curvedArrowLeft', curvedArrowLeftPath),
  shapeFrom('curvedArrowUp', curvedArrowUpPath),
  shapeFrom('curvedArrowDown', curvedArrowDownPath),
  shapeFrom('lineArrowRight', lineArrowRightPath),
  shapeFrom('lineArrowLeft', lineArrowLeftPath),
  shapeFrom('lineArrowUp', lineArrowUpPath),
  shapeFrom('lineArrowDown', lineArrowDownPath),
  shapeFrom('lineHorizontal', lineHorizontalPath),
  shapeFrom('lineVertical', lineVerticalPath),
  shapeFrom('lineDiagonal', lineDiagonalPath),
  shapeFrom('lineElbow', lineElbowPath),
  shapeFrom('bentArrow', bentArrowPath),
  shapeFrom('uTurnArrow', uTurnArrowPath),
  arrowShape('sharpBent'),
  arrowShape('stripedArrow', stripedArrowStripes),
  arrowShape('notched'),
  arrowShape('blockArrow'),
  shapeFrom('circularArrow', circularArrowPath),
  arrowShape('quadArrow'),
  arrowShape('leftRightUp'),
  shapeFrom('arcedLine', arcedLinePath),
];

export const shapeStencils: readonly ShapeStencil[] = [
  { id: 'standard', nameKey: 'shapesPanel.standardShapes', shapes: standardShapes },
  { id: 'arrows', nameKey: 'shapesPanel.arrowShapes', shapes: arrowShapes },
];

export function standardShapeById(id: string): StandardShape | undefined {
  return standardShapes.find((shape) => shape.id === id);
}

export function stencilShapeById(id: string): StandardShape | undefined {
  for (const stencil of shapeStencils) {
    const shape = stencil.shapes.find((candidate) => candidate.id === id);
    if (shape) return shape;
  }
  return undefined;
}
