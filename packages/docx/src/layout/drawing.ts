import type { Chart } from '../types/content/chart';
import type { Image } from '../types/content/image';
import type { Shape } from '../types/content/shape';
import type { ColorValue, ThemeColorSlot } from '../types/colors';

export interface ShapeFillPaint {
  kind?: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture' | 'theme';
  color?: ColorValue;
  themeColor?: ThemeColorSlot;
  gradientType?: 'linear' | 'path' | 'radial' | 'rectangular';
  angle?: number;
  stops?: Array<{ position?: number; color?: ColorValue }>;
  pathShape?: 'circle' | 'rect' | 'shape';
  focusRect?: { left?: number; top?: number; right?: number; bottom?: number };
  rotateWithShape?: boolean;
  patternPreset?: string;
  foregroundColor?: ColorValue;
  backgroundColor?: ColorValue;
  picture?: Image;
  themeRefIndex?: number;
  srcRect?: { left?: number; top?: number; right?: number; bottom?: number };
  fillMode?: 'stretch' | 'tile';
  tile?: {
    offsetX?: number;
    offsetY?: number;
    scaleX?: number;
    scaleY?: number;
    alignment?: string;
    flip?: 'none' | 'x' | 'y' | 'xy';
  };
  stretchRect?: { left?: number; top?: number; right?: number; bottom?: number };
  pictureOpacity?: number;
}

export interface ShapeStrokePaint {
  fill?: ShapeFillPaint;
  width?: number;
  dash?: string;
  customDash?: number[];
  compound?: 'single' | 'double' | 'thickThin' | 'thinThick' | 'triple';
  alignment?: 'center' | 'inset';
  cap?: 'flat' | 'round' | 'square';
  join?: 'bevel' | 'miter' | 'round';
  miterLimit?: number;
  headEnd?: { type?: string; width?: string; length?: string };
  tailEnd?: { type?: string; width?: string; length?: string };
  themeRefIndex?: number;
}

export interface ShapeEffect {
  kind?: 'shadow' | 'glow' | 'reflection' | 'softEdge' | 'blur' | 'unknown';
  color?: ColorValue;
  opacity?: number;
  blurRadius?: number;
  distance?: number;
  direction?: number;
  size?: number;
  rawName?: string;
}

export interface ShapeTextBodyProperties {
  vertical?:
    | 'horizontal'
    | 'vertical'
    | 'vertical270'
    | 'wordArtVertical'
    | 'eastAsianVertical'
    | 'mongolianVertical';
  rotation?: number;
  upright?: boolean;
  anchor?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justified';
  anchorCenter?: boolean;
  columns?: number;
  columnSpacing?: number;
  wrap?: 'square' | 'none';
  horizontalOverflow?: 'clip' | 'overflow';
  verticalOverflow?: 'clip' | 'ellipsis' | 'overflow';
  margins?: { top?: number; bottom?: number; left?: number; right?: number };
  autoFit?: 'none' | 'normal' | 'shape';
  fontScale?: number;
  lineSpacingReduction?: number;
  fromWordArt?: boolean;
  presetTextWarp?: string;
}

export interface DrawingSceneNode {
  kind?: 'shape' | 'group' | 'canvas' | 'picture' | 'chart' | 'graphicFrame' | 'contentPart';
  id?: string;
  name?: string;
  shape?: Shape;
  image?: Image;
  chart?: Chart;
  children?: DrawingSceneNode[];
  transform?: {
    offsetX?: number;
    offsetY?: number;
    width?: number;
    height?: number;
    childOffsetX?: number;
    childOffsetY?: number;
    childWidth?: number;
    childHeight?: number;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
  };
  fill?: ShapeFillPaint;
  effects?: ShapeEffect[];
  relationshipId?: string;
}

export interface DrawingScene {
  version?: number;
  root?: DrawingSceneNode;
  title?: string;
  description?: string;
  decorative?: boolean;
  hidden?: boolean;
}
