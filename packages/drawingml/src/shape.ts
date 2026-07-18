/**
 * DrawingML shape primitives — preset shape types, fill/outline/text-body
 * types, and the pure a:-namespace parsers for gradients, line ends,
 * transforms (a:xfrm), and preset geometry.
 */

import type { ColorValue } from './color';
import { parseColorElement } from './drawing';
import {
  findByFullName,
  findChildrenByLocalName,
  getAttribute,
  getChildElements,
  getLocalName,
  parseNumericAttribute,
  type XmlLike,
} from './xml';

/**
 * Shape types
 *
 * @public
 */
export type ShapeType =
  // Basic shapes
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'rtTriangle'
  | 'parallelogram'
  | 'trapezoid'
  | 'pentagon'
  | 'hexagon'
  | 'heptagon'
  | 'octagon'
  | 'decagon'
  | 'dodecagon'
  | 'star4'
  | 'star5'
  | 'star6'
  | 'star7'
  | 'star8'
  | 'star10'
  | 'star12'
  | 'star16'
  | 'star24'
  | 'star32'
  // Lines and connectors
  | 'line'
  | 'straightConnector1'
  | 'bentConnector2'
  | 'bentConnector3'
  | 'bentConnector4'
  | 'bentConnector5'
  | 'curvedConnector2'
  | 'curvedConnector3'
  | 'curvedConnector4'
  | 'curvedConnector5'
  // Arrows
  | 'rightArrow'
  | 'leftArrow'
  | 'upArrow'
  | 'downArrow'
  | 'leftRightArrow'
  | 'upDownArrow'
  | 'quadArrow'
  | 'leftRightUpArrow'
  | 'bentArrow'
  | 'uturnArrow'
  | 'leftUpArrow'
  | 'bentUpArrow'
  | 'curvedRightArrow'
  | 'curvedLeftArrow'
  | 'curvedUpArrow'
  | 'curvedDownArrow'
  | 'stripedRightArrow'
  | 'notchedRightArrow'
  | 'homePlate'
  | 'chevron'
  | 'rightArrowCallout'
  | 'downArrowCallout'
  | 'leftArrowCallout'
  | 'upArrowCallout'
  | 'leftRightArrowCallout'
  | 'quadArrowCallout'
  | 'circularArrow'
  // Flowchart
  | 'flowChartProcess'
  | 'flowChartAlternateProcess'
  | 'flowChartDecision'
  | 'flowChartInputOutput'
  | 'flowChartPredefinedProcess'
  | 'flowChartInternalStorage'
  | 'flowChartDocument'
  | 'flowChartMultidocument'
  | 'flowChartTerminator'
  | 'flowChartPreparation'
  | 'flowChartManualInput'
  | 'flowChartManualOperation'
  | 'flowChartConnector'
  | 'flowChartOffpageConnector'
  | 'flowChartPunchedCard'
  | 'flowChartPunchedTape'
  | 'flowChartSummingJunction'
  | 'flowChartOr'
  | 'flowChartCollate'
  | 'flowChartSort'
  | 'flowChartExtract'
  | 'flowChartMerge'
  | 'flowChartOnlineStorage'
  | 'flowChartDelay'
  | 'flowChartMagneticTape'
  | 'flowChartMagneticDisk'
  | 'flowChartMagneticDrum'
  | 'flowChartDisplay'
  // Callouts
  | 'wedgeRectCallout'
  | 'wedgeRoundRectCallout'
  | 'wedgeEllipseCallout'
  | 'cloudCallout'
  | 'borderCallout1'
  | 'borderCallout2'
  | 'borderCallout3'
  | 'accentCallout1'
  | 'accentCallout2'
  | 'accentCallout3'
  | 'callout1'
  | 'callout2'
  | 'callout3'
  | 'accentBorderCallout1'
  | 'accentBorderCallout2'
  | 'accentBorderCallout3'
  // Other
  | 'actionButtonBlank'
  | 'actionButtonHome'
  | 'actionButtonHelp'
  | 'actionButtonInformation'
  | 'actionButtonBackPrevious'
  | 'actionButtonForwardNext'
  | 'actionButtonBeginning'
  | 'actionButtonEnd'
  | 'actionButtonReturn'
  | 'actionButtonDocument'
  | 'actionButtonSound'
  | 'actionButtonMovie'
  | 'irregularSeal1'
  | 'irregularSeal2'
  | 'frame'
  | 'halfFrame'
  | 'corner'
  | 'diagStripe'
  | 'chord'
  | 'arc'
  | 'bracketPair'
  | 'bracePair'
  | 'leftBracket'
  | 'rightBracket'
  | 'leftBrace'
  | 'rightBrace'
  | 'can'
  | 'cube'
  | 'bevel'
  | 'donut'
  | 'noSmoking'
  | 'blockArc'
  | 'foldedCorner'
  | 'smileyFace'
  | 'heart'
  | 'lightningBolt'
  | 'sun'
  | 'moon'
  | 'cloud'
  | 'snip1Rect'
  | 'snip2SameRect'
  | 'snip2DiagRect'
  | 'snipRoundRect'
  | 'round1Rect'
  | 'round2SameRect'
  | 'round2DiagRect'
  | 'plaque'
  | 'teardrop'
  | 'mathPlus'
  | 'mathMinus'
  | 'mathMultiply'
  | 'mathDivide'
  | 'mathEqual'
  | 'mathNotEqual'
  | 'gear6'
  | 'gear9'
  | 'funnel'
  | 'pieWedge'
  | 'pie'
  | 'leftCircularArrow'
  | 'leftRightCircularArrow'
  | 'swooshArrow'
  | 'textBox';

/**
 * Shape fill type
 *
 * @public
 */
export interface ShapeFill {
  type: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture';
  /** Solid fill color */
  color?: ColorValue;
  /** Gradient stops for gradient fill */
  gradient?: {
    type: 'linear' | 'radial' | 'rectangular' | 'path';
    angle?: number;
    stops: Array<{
      position: number; // 0-100000
      color: ColorValue;
    }>;
  };
}

/**
 * Shape outline/stroke
 *
 * @public
 */
export interface ShapeOutline {
  /** Line width in EMUs */
  width?: number;
  /** Line color */
  color?: ColorValue;
  /** Line style */
  style?:
    | 'solid'
    | 'dot'
    | 'dash'
    | 'lgDash'
    | 'dashDot'
    | 'lgDashDot'
    | 'lgDashDotDot'
    | 'sysDot'
    | 'sysDash'
    | 'sysDashDot'
    | 'sysDashDotDot';
  /** Line cap */
  cap?: 'flat' | 'round' | 'square';
  /** Line join */
  join?: 'bevel' | 'miter' | 'round';
  /** Head arrow */
  headEnd?: {
    type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
    width?: 'sm' | 'med' | 'lg';
    length?: 'sm' | 'med' | 'lg';
  };
  /** Tail arrow */
  tailEnd?: {
    type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
    width?: 'sm' | 'med' | 'lg';
    length?: 'sm' | 'med' | 'lg';
  };
}

/**
 * Text body inside a shape.
 *
 * `TContent` is the host's paragraph/content type — DrawingML itself does
 * not prescribe what a paragraph looks like (docx uses `w:p`, pptx `a:p`).
 *
 * @public
 */
export interface ShapeTextBody<TContent = unknown> {
  /** Text direction */
  vertical?: boolean;
  /** Rotation */
  rotation?: number;
  /** Anchor/vertical alignment */
  anchor?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justified';
  /** Anchor center */
  anchorCenter?: boolean;
  /** Auto fit */
  autoFit?: 'none' | 'normal' | 'shape';
  /** Text margins */
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  /** Paragraphs inside the shape */
  content: TContent[];
}

/**
 * Width/height pair in EMUs (structural match for host image-size types)
 *
 * @public
 */
export interface Size2D {
  width: number;
  height: number;
}

/**
 * Rotation/flip transform (structural match for host transform types)
 *
 * @public
 */
export interface Transform2D {
  /** Rotation in degrees */
  rotation?: number;
  /** Flip horizontal */
  flipH?: boolean;
  /** Flip vertical */
  flipV?: boolean;
}

/**
 * Normalized DrawingML preset-geometry path command.
 *
 * Coordinates are shape-local fractions in the `[0, 1]` box. Consumers scale
 * `x/y` and control points by the shape's resolved width/height.
 *
 * @public
 */
export type PresetGeometryPathCommand =
  | { type: 'move'; x: number; y: number }
  | { type: 'line'; x: number; y: number }
  | { type: 'quad'; cpx: number; cpy: number; x: number; y: number }
  | {
      type: 'cubic';
      cp1x: number;
      cp1y: number;
      cp2x: number;
      cp2y: number;
      x: number;
      y: number;
    }
  | { type: 'close' };

type RawGeometryPathCommand =
  | { type: 'move'; x: number; y: number }
  | { type: 'line'; x: number; y: number }
  | { type: 'quad'; cpx: number; cpy: number; x: number; y: number }
  | {
      type: 'cubic';
      cp1x: number;
      cp1y: number;
      cp2x: number;
      cp2y: number;
      x: number;
      y: number;
    }
  | { type: 'close' };

type RawPoint = { x: number; y: number };

/** Basic preset geometries supported by the v1 upstream shape pipeline. */
export type BasicPresetGeometry = 'rect' | 'roundRect' | 'ellipse' | 'line' | 'triangle';

const ELLIPSE_KAPPA = 0.5522847498307936;
const ROUND_RECT_RADIUS = 0.16666666666666666;
const ANGLE_UNITS_PER_DEGREE = 60000;
const MAX_CUSTOM_PATH_COMMANDS = 2048;
const MAX_CUSTOM_GUIDES = 512;
const MAX_ABS_CUSTOM_PATH_NUMBER = 1_000_000_000;
const MAX_ABS_NORMALIZED_PATH_NUMBER = 10_000;

type GuideValues = Map<string, number>;

function closePath(): PresetGeometryPathCommand {
  return { type: 'close' };
}

/**
 * Return true when the preset has a built-in normalized path in this package.
 *
 * @public
 */
export function isSupportedBasicPresetGeometry(shapeType: string): boolean {
  return presetGeometryToPath(shapeType) !== null;
}

function polygonPath(points: RawPoint[]): PresetGeometryPathCommand[] {
  if (points.length === 0) return [];
  return [
    { type: 'move', x: points[0].x, y: points[0].y },
    ...points.slice(1).map(
      (point): PresetGeometryPathCommand => ({
        type: 'line',
        x: point.x,
        y: point.y,
      })
    ),
    closePath(),
  ];
}

function regularPolygonPath(sides: number, rotation = -Math.PI / 2): PresetGeometryPathCommand[] {
  const points = Array.from({ length: sides }, (_, index) => {
    const angle = rotation + (index * Math.PI * 2) / sides;
    return { x: 0.5 + Math.cos(angle) * 0.5, y: 0.5 + Math.sin(angle) * 0.5 };
  });
  return polygonPath(points);
}

function starPath(points: number, innerRadius = 0.45): PresetGeometryPathCommand[] {
  const vertices = Array.from({ length: points * 2 }, (_, index) => {
    const angle = -Math.PI / 2 + (index * Math.PI) / points;
    const radius = index % 2 === 0 ? 0.5 : 0.5 * innerRadius;
    return { x: 0.5 + Math.cos(angle) * radius, y: 0.5 + Math.sin(angle) * radius };
  });
  return polygonPath(vertices);
}

function clampFraction(value: number | undefined, fallback: number): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.max(0, Math.min(1, value > 1 ? value / 100000 : value));
}

function arrowPath(direction: 'right' | 'left' | 'up' | 'down', adjustment?: number) {
  const head = clampFraction(adjustment, 0.5);
  const right = polygonPath([
    { x: 0, y: 0.25 },
    { x: 1 - head, y: 0.25 },
    { x: 1 - head, y: 0 },
    { x: 1, y: 0.5 },
    { x: 1 - head, y: 1 },
    { x: 1 - head, y: 0.75 },
    { x: 0, y: 0.75 },
  ]);
  const transform = (point: RawPoint): RawPoint => {
    switch (direction) {
      case 'left':
        return { x: 1 - point.x, y: point.y };
      case 'up':
        return { x: point.y, y: 1 - point.x };
      case 'down':
        return { x: point.y, y: point.x };
      case 'right':
        return point;
    }
  };
  return right.map((command): PresetGeometryPathCommand => {
    if (command.type !== 'move' && command.type !== 'line') return command;
    const point = transform(command);
    return { type: command.type, x: point.x, y: point.y };
  });
}

function bentConnectorPath(segments: number, adjustment?: number): PresetGeometryPathCommand[] {
  const bend = clampFraction(adjustment, 0.5);
  if (segments <= 2) {
    return [
      { type: 'move', x: 0, y: 0 },
      { type: 'line', x: bend, y: 0 },
      { type: 'line', x: bend, y: 1 },
      { type: 'line', x: 1, y: 1 },
    ];
  }
  const commands: PresetGeometryPathCommand[] = [{ type: 'move', x: 0, y: 0 }];
  for (let index = 1; index < segments; index++) {
    const fraction = index / segments;
    commands.push(
      index % 2 === 1
        ? { type: 'line', x: index === 1 ? bend : fraction, y: (index - 1) / segments }
        : { type: 'line', x: (index - 1) / segments, y: fraction }
    );
  }
  commands.push({ type: 'line', x: 1, y: 1 });
  return commands;
}

function curvedConnectorPath(segments: number): PresetGeometryPathCommand[] {
  if (segments <= 2) {
    return [
      { type: 'move', x: 0, y: 0 },
      { type: 'cubic', cp1x: 0.5, cp1y: 0, cp2x: 0.5, cp2y: 1, x: 1, y: 1 },
    ];
  }
  const commands: PresetGeometryPathCommand[] = [{ type: 'move', x: 0, y: 0 }];
  for (let index = 0; index < segments - 1; index++) {
    const start = index / (segments - 1);
    const end = (index + 1) / (segments - 1);
    commands.push({
      type: 'cubic',
      cp1x: start + (end - start) * 0.5,
      cp1y: start,
      cp2x: start + (end - start) * 0.5,
      cp2y: end,
      x: end,
      y: end,
    });
  }
  return commands;
}

/**
 * Generate a normalized path for the DrawingML preset geometries currently
 * supported by the DOCX layout upstream. Unsupported presets return `null`
 * rather than guessing at Word's shape guide formulas.
 *
 * @public
 */
export function presetGeometryToPath(
  shapeType: string,
  adjustments: Record<string, number> = {}
): PresetGeometryPathCommand[] | null {
  switch (shapeType) {
    case 'rect':
      return [
        { type: 'move', x: 0, y: 0 },
        { type: 'line', x: 1, y: 0 },
        { type: 'line', x: 1, y: 1 },
        { type: 'line', x: 0, y: 1 },
        closePath(),
      ];

    case 'roundRect': {
      const r = ROUND_RECT_RADIUS;
      return [
        { type: 'move', x: r, y: 0 },
        { type: 'line', x: 1 - r, y: 0 },
        { type: 'quad', cpx: 1, cpy: 0, x: 1, y: r },
        { type: 'line', x: 1, y: 1 - r },
        { type: 'quad', cpx: 1, cpy: 1, x: 1 - r, y: 1 },
        { type: 'line', x: r, y: 1 },
        { type: 'quad', cpx: 0, cpy: 1, x: 0, y: 1 - r },
        { type: 'line', x: 0, y: r },
        { type: 'quad', cpx: 0, cpy: 0, x: r, y: 0 },
        closePath(),
      ];
    }

    case 'ellipse':
      return [
        { type: 'move', x: 1, y: 0.5 },
        {
          type: 'cubic',
          cp1x: 1,
          cp1y: 0.5 + ELLIPSE_KAPPA / 2,
          cp2x: 0.5 + ELLIPSE_KAPPA / 2,
          cp2y: 1,
          x: 0.5,
          y: 1,
        },
        {
          type: 'cubic',
          cp1x: 0.5 - ELLIPSE_KAPPA / 2,
          cp1y: 1,
          cp2x: 0,
          cp2y: 0.5 + ELLIPSE_KAPPA / 2,
          x: 0,
          y: 0.5,
        },
        {
          type: 'cubic',
          cp1x: 0,
          cp1y: 0.5 - ELLIPSE_KAPPA / 2,
          cp2x: 0.5 - ELLIPSE_KAPPA / 2,
          cp2y: 0,
          x: 0.5,
          y: 0,
        },
        {
          type: 'cubic',
          cp1x: 0.5 + ELLIPSE_KAPPA / 2,
          cp1y: 0,
          cp2x: 1,
          cp2y: 0.5 - ELLIPSE_KAPPA / 2,
          x: 1,
          y: 0.5,
        },
        closePath(),
      ];

    case 'line':
    case 'straightConnector1':
      return [
        { type: 'move', x: 0, y: 0 },
        { type: 'line', x: 1, y: 1 },
      ];

    case 'triangle':
    case 'isosTriangle':
      return [
        { type: 'move', x: 0.5, y: 0 },
        { type: 'line', x: 1, y: 1 },
        { type: 'line', x: 0, y: 1 },
        closePath(),
      ];

    case 'rtTriangle':
      return polygonPath([
        { x: 0, y: 0 },
        { x: 1, y: 1 },
        { x: 0, y: 1 },
      ]);

    case 'diamond':
    case 'flowChartDecision':
      return polygonPath([
        { x: 0.5, y: 0 },
        { x: 1, y: 0.5 },
        { x: 0.5, y: 1 },
        { x: 0, y: 0.5 },
      ]);

    case 'parallelogram': {
      const inset = clampFraction(adjustments.adj, 0.25);
      return polygonPath([
        { x: inset, y: 0 },
        { x: 1, y: 0 },
        { x: 1 - inset, y: 1 },
        { x: 0, y: 1 },
      ]);
    }

    case 'trapezoid': {
      const inset = clampFraction(adjustments.adj, 0.2);
      return polygonPath([
        { x: inset, y: 0 },
        { x: 1 - inset, y: 0 },
        { x: 1, y: 1 },
        { x: 0, y: 1 },
      ]);
    }

    case 'pentagon':
    case 'flowChartOffpageConnector':
      return regularPolygonPath(5);
    case 'hexagon':
      return regularPolygonPath(6);
    case 'heptagon':
      return regularPolygonPath(7);
    case 'octagon':
      return regularPolygonPath(8);
    case 'decagon':
      return regularPolygonPath(10);
    case 'dodecagon':
      return regularPolygonPath(12);
    case 'star4':
      return starPath(4);
    case 'star5':
      return starPath(5);
    case 'star6':
      return starPath(6);
    case 'star7':
      return starPath(7);
    case 'star8':
      return starPath(8);
    case 'star10':
      return starPath(10);
    case 'star12':
      return starPath(12);
    case 'star16':
      return starPath(16);
    case 'star24':
      return starPath(24);
    case 'star32':
      return starPath(32);

    case 'bentConnector2':
      return bentConnectorPath(2, adjustments.adj1);
    case 'bentConnector3':
      return bentConnectorPath(3, adjustments.adj1);
    case 'bentConnector4':
      return bentConnectorPath(4, adjustments.adj1);
    case 'bentConnector5':
      return bentConnectorPath(5, adjustments.adj1);
    case 'curvedConnector2':
      return curvedConnectorPath(2);
    case 'curvedConnector3':
      return curvedConnectorPath(3);
    case 'curvedConnector4':
      return curvedConnectorPath(4);
    case 'curvedConnector5':
      return curvedConnectorPath(5);

    case 'rightArrow':
      return arrowPath('right', adjustments.adj2);
    case 'leftArrow':
      return arrowPath('left', adjustments.adj2);
    case 'upArrow':
      return arrowPath('up', adjustments.adj2);
    case 'downArrow':
      return arrowPath('down', adjustments.adj2);
    case 'leftRightArrow':
      return polygonPath([
        { x: 0, y: 0.5 },
        { x: 0.25, y: 0 },
        { x: 0.25, y: 0.25 },
        { x: 0.75, y: 0.25 },
        { x: 0.75, y: 0 },
        { x: 1, y: 0.5 },
        { x: 0.75, y: 1 },
        { x: 0.75, y: 0.75 },
        { x: 0.25, y: 0.75 },
        { x: 0.25, y: 1 },
      ]);
    case 'upDownArrow':
      return polygonPath([
        { x: 0.5, y: 0 },
        { x: 1, y: 0.25 },
        { x: 0.75, y: 0.25 },
        { x: 0.75, y: 0.75 },
        { x: 1, y: 0.75 },
        { x: 0.5, y: 1 },
        { x: 0, y: 0.75 },
        { x: 0.25, y: 0.75 },
        { x: 0.25, y: 0.25 },
        { x: 0, y: 0.25 },
      ]);
    case 'chevron':
      return polygonPath([
        { x: 0, y: 0 },
        { x: 0.65, y: 0 },
        { x: 1, y: 0.5 },
        { x: 0.65, y: 1 },
        { x: 0, y: 1 },
        { x: 0.35, y: 0.5 },
      ]);
    case 'homePlate':
      return polygonPath([
        { x: 0, y: 0 },
        { x: 0.75, y: 0 },
        { x: 1, y: 0.5 },
        { x: 0.75, y: 1 },
        { x: 0, y: 1 },
      ]);

    case 'flowChartProcess':
    case 'flowChartAlternateProcess':
    case 'flowChartPredefinedProcess':
    case 'flowChartInternalStorage':
    case 'flowChartPreparation':
    case 'flowChartManualOperation':
    case 'flowChartMagneticTape':
    case 'flowChartMagneticDisk':
    case 'flowChartMagneticDrum':
    case 'flowChartDisplay':
    case 'textBox':
      return presetGeometryToPath('rect');
    case 'flowChartConnector':
      return presetGeometryToPath('ellipse');
    case 'flowChartInputOutput':
    case 'flowChartManualInput':
      return presetGeometryToPath('parallelogram', adjustments);
    case 'flowChartTerminator':
      return presetGeometryToPath('roundRect');

    default:
      return null;
  }
}

function parseFinitePathNumber(value: string | null): number | undefined {
  if (value === null) return undefined;
  const n = Number(value);
  if (!Number.isFinite(n) || Math.abs(n) > MAX_ABS_CUSTOM_PATH_NUMBER) return undefined;
  return n;
}

function parseFinitePathAttribute(
  element: XmlLike | null | undefined,
  name: string
): number | undefined {
  return parseFinitePathNumber(getAttribute(element, null, name));
}

function finiteGuideValue(value: number): number | undefined {
  return Number.isFinite(value) && Math.abs(value) <= MAX_ABS_CUSTOM_PATH_NUMBER
    ? value
    : undefined;
}

function guideTokenValue(token: string | undefined, values: GuideValues): number | undefined {
  if (!token) return undefined;
  const numeric = parseFinitePathNumber(token);
  if (numeric !== undefined) return numeric;
  return values.get(token);
}

function angleToRadians(angle: number): number {
  return (angle / ANGLE_UNITS_PER_DEGREE) * (Math.PI / 180);
}

/** Evaluate one ECMA-376 DrawingML guide formula with bounded finite arithmetic. */
function evaluateGuideFormula(formula: string | null, values: GuideValues): number | undefined {
  if (!formula) return undefined;
  const [operator, ...tokens] = formula.trim().split(/\s+/u);
  const args = tokens.map((token) => guideTokenValue(token, values));
  if (args.some((value) => value === undefined)) return undefined;
  const numbers = args as number[];
  const [x = 0, y = 0, z = 0] = numbers;
  let result: number;

  switch (operator) {
    case 'val':
      result = x;
      break;
    case '*/':
      if (z === 0) return undefined;
      result = (x * y) / z;
      break;
    case '+-':
      result = x + y - z;
      break;
    case '+/':
      if (z === 0) return undefined;
      result = (x + y) / z;
      break;
    case '?:':
      result = x > 0 ? y : z;
      break;
    case 'abs':
      result = Math.abs(x);
      break;
    case 'at2':
      result = (Math.atan2(y, x) * 180 * ANGLE_UNITS_PER_DEGREE) / Math.PI;
      break;
    case 'cat2':
      result = x * Math.cos(Math.atan2(z, y));
      break;
    case 'cos':
      result = x * Math.cos(angleToRadians(y));
      break;
    case 'max':
      result = Math.max(x, y);
      break;
    case 'min':
      result = Math.min(x, y);
      break;
    case 'mod':
      result = Math.hypot(x, y, z);
      break;
    case 'pin':
      result = Math.min(Math.max(y, x), z);
      break;
    case 'sat2':
      result = x * Math.sin(Math.atan2(z, y));
      break;
    case 'sin':
      result = x * Math.sin(angleToRadians(y));
      break;
    case 'sqrt':
      if (x < 0) return undefined;
      result = Math.sqrt(x);
      break;
    case 'tan':
      result = x * Math.tan(angleToRadians(y));
      break;
    default:
      return undefined;
  }
  return finiteGuideValue(result);
}

function standardGuideValues(width: number, height: number): GuideValues {
  const w = width > 0 ? width : 1;
  const h = height > 0 ? height : 1;
  const ss = Math.min(w, h);
  const ls = Math.max(w, h);
  const values: GuideValues = new Map([
    ['l', 0],
    ['t', 0],
    ['r', w],
    ['b', h],
    ['w', w],
    ['h', h],
    ['hc', w / 2],
    ['vc', h / 2],
    ['ss', ss],
    ['ls', ls],
    ['cd2', 10800000],
    ['cd4', 5400000],
    ['cd8', 2700000],
    ['3cd4', 16200000],
    ['3cd8', 8100000],
    ['5cd8', 13500000],
    ['7cd8', 18900000],
  ]);
  for (const divisor of [2, 3, 4, 5, 6, 8, 10, 12, 16, 20, 32]) {
    values.set(`wd${divisor}`, w / divisor);
    values.set(`hd${divisor}`, h / divisor);
    values.set(`ssd${divisor}`, ss / divisor);
    values.set(`lsd${divisor}`, ls / divisor);
  }
  return values;
}

function applyGuideList(parent: XmlLike | null, listName: string, values: GuideValues): void {
  const list = getFirstLocalChild(parent, listName);
  if (!list) return;
  const guides = findChildrenByLocalName(list, 'gd').slice(0, MAX_CUSTOM_GUIDES);
  const pending = [...guides];
  for (let pass = 0; pass < guides.length && pending.length > 0; pass++) {
    for (let index = pending.length - 1; index >= 0; index--) {
      const guide = pending[index];
      const name = getAttribute(guide, null, 'name');
      const value = evaluateGuideFormula(getAttribute(guide, null, 'fmla'), values);
      if (!name || value === undefined) continue;
      values.set(name, value);
      pending.splice(index, 1);
    }
  }
}

function buildCustomGeometryGuides(
  custGeom: XmlLike | null | undefined,
  width: number,
  height: number
): GuideValues {
  const values = standardGuideValues(width, height);
  applyGuideList(custGeom ?? null, 'avLst', values);
  applyGuideList(custGeom ?? null, 'gdLst', values);
  return values;
}

function parseGuidePathAttribute(
  element: XmlLike | null | undefined,
  name: string,
  values: GuideValues
): number | undefined {
  return guideTokenValue(getAttribute(element, null, name) ?? undefined, values);
}

function getFirstLocalChild(parent: XmlLike | null | undefined, localName: string): XmlLike | null {
  return (
    getChildElements(parent).find((child) => getLocalName(child.name || '') === localName) ?? null
  );
}

function parsePathPoint(pt: XmlLike | null | undefined, values?: GuideValues): RawPoint | null {
  const x = values ? parseGuidePathAttribute(pt, 'x', values) : parseFinitePathAttribute(pt, 'x');
  const y = values ? parseGuidePathAttribute(pt, 'y', values) : parseFinitePathAttribute(pt, 'y');
  if (x === undefined || y === undefined) return null;
  return { x, y };
}

function pathPoints(command: XmlLike, values?: GuideValues): RawPoint[] {
  const points: RawPoint[] = [];
  for (const pt of findChildrenByLocalName(command, 'pt')) {
    const point = parsePathPoint(pt, values);
    if (point) points.push(point);
  }
  return points;
}

function angleUnitsToRadians(value: number): number {
  return (value / ANGLE_UNITS_PER_DEGREE) * (Math.PI / 180);
}

function arcToCubics(
  current: RawPoint,
  wR: number,
  hR: number,
  stAng: number,
  swAng: number
): RawGeometryPathCommand[] {
  const rx = Math.abs(wR);
  const ry = Math.abs(hR);
  if (rx <= 0 || ry <= 0 || swAng === 0) return [];

  const start = angleUnitsToRadians(stAng);
  const sweep = angleUnitsToRadians(swAng);
  const center = {
    x: current.x - rx * Math.cos(start),
    y: current.y - ry * Math.sin(start),
  };
  const segmentCount = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2)));
  const segmentSweep = sweep / segmentCount;
  const cubics: RawGeometryPathCommand[] = [];
  let p0 = current;

  for (let i = 0; i < segmentCount; i++) {
    const t0 = start + segmentSweep * i;
    const t1 = t0 + segmentSweep;
    const alpha = (4 / 3) * Math.tan(segmentSweep / 4);
    const p1 = {
      x: center.x + rx * Math.cos(t1),
      y: center.y + ry * Math.sin(t1),
    };
    const d0 = { x: -rx * Math.sin(t0), y: ry * Math.cos(t0) };
    const d1 = { x: -rx * Math.sin(t1), y: ry * Math.cos(t1) };

    cubics.push({
      type: 'cubic',
      cp1x: p0.x + alpha * d0.x,
      cp1y: p0.y + alpha * d0.y,
      cp2x: p1.x - alpha * d1.x,
      cp2y: p1.y - alpha * d1.y,
      x: p1.x,
      y: p1.y,
    });
    p0 = p1;
  }

  return cubics;
}

function collectRawCommandPoints(command: RawGeometryPathCommand): RawPoint[] {
  switch (command.type) {
    case 'move':
    case 'line':
      return [{ x: command.x, y: command.y }];
    case 'quad':
      return [
        { x: command.cpx, y: command.cpy },
        { x: command.x, y: command.y },
      ];
    case 'cubic':
      return [
        { x: command.cp1x, y: command.cp1y },
        { x: command.cp2x, y: command.cp2y },
        { x: command.x, y: command.y },
      ];
    case 'close':
      return [];
  }
}

function inferPositiveExtent(points: RawPoint[], axis: 'x' | 'y'): number {
  const max = points.reduce((acc, point) => Math.max(acc, point[axis]), 0);
  return max > 0 ? max : 1;
}

function normalizePathNumber(value: number, denominator: number): number {
  const normalized = value / denominator;
  if (!Number.isFinite(normalized)) return 0;
  return Math.max(
    -MAX_ABS_NORMALIZED_PATH_NUMBER,
    Math.min(MAX_ABS_NORMALIZED_PATH_NUMBER, normalized)
  );
}

function normalizeRawPath(
  commands: RawGeometryPathCommand[],
  pathWidth: number | undefined,
  pathHeight: number | undefined
): PresetGeometryPathCommand[] {
  const points = commands.flatMap(collectRawCommandPoints);
  const width = pathWidth && pathWidth > 0 ? pathWidth : inferPositiveExtent(points, 'x');
  const height = pathHeight && pathHeight > 0 ? pathHeight : inferPositiveExtent(points, 'y');
  const x = (value: number) => normalizePathNumber(value, width);
  const y = (value: number) => normalizePathNumber(value, height);

  return commands.map((command) => {
    switch (command.type) {
      case 'move':
        return { type: 'move', x: x(command.x), y: y(command.y) };
      case 'line':
        return { type: 'line', x: x(command.x), y: y(command.y) };
      case 'quad':
        return {
          type: 'quad',
          cpx: x(command.cpx),
          cpy: y(command.cpy),
          x: x(command.x),
          y: y(command.y),
        };
      case 'cubic':
        return {
          type: 'cubic',
          cp1x: x(command.cp1x),
          cp1y: y(command.cp1y),
          cp2x: x(command.cp2x),
          cp2y: y(command.cp2y),
          x: x(command.x),
          y: y(command.y),
        };
      case 'close':
        return closePath();
    }
  });
}

function parseCustomPathElement(
  path: XmlLike,
  remainingBudget: number,
  geometryValues: GuideValues
): PresetGeometryPathCommand[] {
  const rawCommands: RawGeometryPathCommand[] = [];
  let current: RawPoint | null = null;
  const pathWidth = parseGuidePathAttribute(path, 'w', geometryValues);
  const pathHeight = parseGuidePathAttribute(path, 'h', geometryValues);
  const values = new Map(geometryValues);
  if (pathWidth !== undefined && pathWidth > 0) {
    values.set('w', pathWidth);
    values.set('r', pathWidth);
    values.set('hc', pathWidth / 2);
  }
  if (pathHeight !== undefined && pathHeight > 0) {
    values.set('h', pathHeight);
    values.set('b', pathHeight);
    values.set('vc', pathHeight / 2);
  }

  for (const child of getChildElements(path)) {
    if (rawCommands.length >= remainingBudget) break;

    switch (getLocalName(child.name || '')) {
      case 'moveTo': {
        const pt = parsePathPoint(getFirstLocalChild(child, 'pt'), values);
        if (!pt) break;
        rawCommands.push({ type: 'move', x: pt.x, y: pt.y });
        current = pt;
        break;
      }

      case 'lnTo': {
        const pt = parsePathPoint(getFirstLocalChild(child, 'pt'), values);
        if (!pt) break;
        rawCommands.push({ type: 'line', x: pt.x, y: pt.y });
        current = pt;
        break;
      }

      case 'quadBezTo': {
        const points = pathPoints(child, values);
        if (points.length < 2) break;
        const [cp, end] = points;
        rawCommands.push({ type: 'quad', cpx: cp.x, cpy: cp.y, x: end.x, y: end.y });
        current = end;
        break;
      }

      case 'cubicBezTo': {
        const points = pathPoints(child, values);
        if (points.length < 3) break;
        const [cp1, cp2, end] = points;
        rawCommands.push({
          type: 'cubic',
          cp1x: cp1.x,
          cp1y: cp1.y,
          cp2x: cp2.x,
          cp2y: cp2.y,
          x: end.x,
          y: end.y,
        });
        current = end;
        break;
      }

      case 'arcTo': {
        if (!current) break;
        const wR = parseGuidePathAttribute(child, 'wR', values);
        const hR = parseGuidePathAttribute(child, 'hR', values);
        const stAng = parseGuidePathAttribute(child, 'stAng', values);
        const swAng = parseGuidePathAttribute(child, 'swAng', values);
        if (wR === undefined || hR === undefined || stAng === undefined || swAng === undefined) {
          break;
        }
        const cubics: RawGeometryPathCommand[] = arcToCubics(current, wR, hR, stAng, swAng).slice(
          0,
          remainingBudget - rawCommands.length
        );
        rawCommands.push(...cubics);
        const last: RawGeometryPathCommand | undefined = cubics[cubics.length - 1];
        if (last && last.type === 'cubic') current = { x: last.x, y: last.y };
        break;
      }

      case 'close':
        rawCommands.push(closePath());
        break;
    }
  }

  if (rawCommands.length === 0) return [];
  return normalizeRawPath(rawCommands, pathWidth, pathHeight);
}

/**
 * Parse DrawingML custom geometry (`a:custGeom`) into the same normalized path
 * command format used by supported preset geometries. `a:arcTo` segments are
 * flattened into cubic Bezier commands so downstream shape renderers do not
 * need a separate arc primitive.
 *
 * @public
 */
export function parseCustomGeometryPath(
  custGeom: XmlLike | null | undefined
): PresetGeometryPathCommand[] | null {
  const pathLst = getFirstLocalChild(custGeom, 'pathLst');
  if (!pathLst) return null;

  const geometryPath: PresetGeometryPathCommand[] = [];
  for (const path of findChildrenByLocalName(pathLst, 'path')) {
    if (geometryPath.length >= MAX_CUSTOM_PATH_COMMANDS) break;
    const pathWidth = parseFinitePathAttribute(path, 'w') ?? 1;
    const pathHeight = parseFinitePathAttribute(path, 'h') ?? 1;
    const values = buildCustomGeometryGuides(custGeom, pathWidth, pathHeight);
    geometryPath.push(
      ...parseCustomPathElement(path, MAX_CUSTOM_PATH_COMMANDS - geometryPath.length, values)
    );
  }

  return geometryPath.length > 0 ? geometryPath : null;
}

/**
 * Resolve a preset geometry and its authored `a:avLst` adjustment guides into
 * normalized commands. Unknown presets return `null` and remain explicitly
 * unsupported downstream.
 *
 * @public
 */
export function parsePresetGeometryPath(
  spPr: XmlLike | null | undefined
): PresetGeometryPathCommand[] | null {
  const preset = spPr ? findByFullName(spPr, 'a:prstGeom') : null;
  if (!preset) return null;
  const shapeType = getAttribute(preset, null, 'prst');
  if (!shapeType) return null;

  const values = standardGuideValues(100000, 100000);
  applyGuideList(preset, 'avLst', values);
  const adjustments: Record<string, number> = {};
  const adjustmentList = getFirstLocalChild(preset, 'avLst');
  for (const guide of findChildrenByLocalName(adjustmentList, 'gd').slice(0, MAX_CUSTOM_GUIDES)) {
    const name = getAttribute(guide, null, 'name');
    if (!name) continue;
    const value = values.get(name);
    if (value !== undefined) adjustments[name] = value;
  }
  return presetGeometryToPath(shapeType, adjustments);
}

/**
 * Convert rotation value (1/60000 of a degree) to degrees
 *
 * @public
 */
export function rotToDegrees(rot: string | null | undefined): number | undefined {
  if (!rot) return undefined;
  const val = parseInt(rot, 10);
  if (isNaN(val)) return undefined;
  return val / 60000;
}

/**
 * Parse gradient fill (a:gradFill)
 *
 * @public
 */
export function parseGradientFill(gradFill: XmlLike): ShapeFill {
  const children = getChildElements(gradFill);

  // Determine gradient type
  let gradientType: 'linear' | 'radial' | 'rectangular' | 'path' = 'linear';
  let angle: number | undefined;

  // Check for linear gradient
  const lin = children.find((el) => el.name === 'a:lin');
  if (lin) {
    gradientType = 'linear';
    const ang = getAttribute(lin, null, 'ang');
    if (ang) {
      // Angle is in 60000ths of a degree
      angle = parseInt(ang, 10) / 60000;
    }
  }

  // Check for path gradient (radial)
  const path = children.find((el) => el.name === 'a:path');
  if (path) {
    const pathType = getAttribute(path, null, 'path');
    if (pathType === 'circle') {
      gradientType = 'radial';
    } else if (pathType === 'rect') {
      gradientType = 'rectangular';
    } else {
      gradientType = 'path';
    }
  }

  // Parse gradient stops
  const gsLst = children.find((el) => el.name === 'a:gsLst');
  const stops: Array<{ position: number; color: ColorValue }> = [];

  if (gsLst) {
    const gsElements = findChildrenByLocalName(gsLst, 'gs');
    for (const gs of gsElements) {
      const pos = getAttribute(gs, null, 'pos');
      const position = pos ? parseInt(pos, 10) : 0;
      const color = parseColorElement(gs);
      if (color) {
        stops.push({ position, color });
      }
    }
  }

  return {
    type: 'gradient',
    gradient: {
      type: gradientType,
      angle,
      stops,
    },
  };
}

/**
 * Parse line end (arrow head/tail)
 *
 * @public
 */
export function parseLineEnd(element: XmlLike): NonNullable<ShapeOutline['headEnd']> {
  const type = getAttribute(element, null, 'type') ?? 'none';
  const w = getAttribute(element, null, 'w') as 'sm' | 'med' | 'lg' | undefined;
  const len = getAttribute(element, null, 'len') as 'sm' | 'med' | 'lg' | undefined;

  type LineEndType = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
  const typeMap: Record<string, LineEndType> = {
    none: 'none',
    triangle: 'triangle',
    stealth: 'stealth',
    diamond: 'diamond',
    oval: 'oval',
    arrow: 'arrow',
  };

  return {
    type: typeMap[type] ?? 'none',
    width: w,
    length: len,
  };
}

/**
 * Parse transform from a:xfrm element
 *
 * @public
 */
export function parseTransform(xfrm: XmlLike | null): {
  size: Size2D;
  transform?: Transform2D;
  offset?: { x: number; y: number };
} {
  if (!xfrm) {
    return { size: { width: 0, height: 0 } };
  }

  // Get extent (size)
  const ext = findByFullName(xfrm, 'a:ext');
  const cx = parseNumericAttribute(ext, null, 'cx') ?? 0;
  const cy = parseNumericAttribute(ext, null, 'cy') ?? 0;

  const size: Size2D = { width: cx, height: cy };

  // Get offset
  const off = findByFullName(xfrm, 'a:off');
  let offset: { x: number; y: number } | undefined;
  if (off) {
    const x = parseNumericAttribute(off, null, 'x') ?? 0;
    const y = parseNumericAttribute(off, null, 'y') ?? 0;
    offset = { x, y };
  }

  // Get transform properties
  const rot = getAttribute(xfrm, null, 'rot');
  const flipH = getAttribute(xfrm, null, 'flipH') === '1';
  const flipV = getAttribute(xfrm, null, 'flipV') === '1';

  const rotation = rotToDegrees(rot);

  let transform: Transform2D | undefined;
  if (rotation !== undefined || flipH || flipV) {
    transform = {};
    if (rotation !== undefined) transform.rotation = rotation;
    if (flipH) transform.flipH = true;
    if (flipV) transform.flipV = true;
  }

  return { size, transform, offset };
}

/**
 * Parse preset geometry to get shape type
 *
 * @public
 */
export function parseShapeType(spPr: XmlLike | null): ShapeType {
  if (!spPr) {
    return 'rect';
  }

  // Check for preset geometry
  const prstGeom = findByFullName(spPr, 'a:prstGeom');
  if (prstGeom) {
    const prst = getAttribute(prstGeom, null, 'prst');
    if (prst) {
      return prst as ShapeType;
    }
  }

  // Custom geometry paths are carried separately; keep a safe preset fallback
  // for consumers that still only understand shapeType.
  const custGeom = findByFullName(spPr, 'a:custGeom');
  if (custGeom) {
    return 'rect';
  }

  return 'rect';
}
