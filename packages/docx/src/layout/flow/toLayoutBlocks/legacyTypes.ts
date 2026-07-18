import type { PresetGeometryPathCommand } from '@betteroffice/drawingml';

import type {
  NumberFormat,
  ParagraphFormatting,
  ParagraphPropertyChange,
  SectionProperties,
  TextFormatting,
  ThemeColorSlot,
  UnderlineStyle,
} from '../../../types/document';
import type { ShapeFillPaint } from '../../../types/content/shape';
import type { RevisionInfo } from '../../../types/content/trackedChange';

/**
 * Attribute shapes accepted by the legacy tree-to-layout adapter.
 *
 * These are plain data contracts. Keeping them beside the adapter prevents
 * the rendering pipeline from depending on the retired editor schema.
 */
export interface LegacyParagraphAttrs extends ParagraphFormatting {
  paraId?: string;
  textId?: string;
  renderedPageBreakBefore?: boolean;
  defaultTextFormatting?: TextFormatting;
  sectionBreakType?: 'nextPage' | 'continuous' | 'oddPage' | 'evenPage';
  bookmarks?: Array<{ id: number; name: string }>;
  _originalFormatting?: ParagraphFormatting;
  _sectionProperties?: SectionProperties;
  listNumFmt?: NumberFormat;
  listIsBullet?: boolean;
  listMarker?: string;
  listMarkerHidden?: boolean;
  listMarkerFontFamily?: string;
  listMarkerFontSize?: number;
  listMarkerSuffix?: 'tab' | 'space' | 'nothing';
  listLevelNumFmts?: NumberFormat[];
  listAbstractNumId?: number;
  listStartOverride?: number;
  pPrIns?: RevisionInfo | null;
  pPrDel?: RevisionInfo | null;
  pPrChange?: ParagraphPropertyChange[] | null;
}

export interface LegacyChartAttrs {
  chartJson?: string;
  chartType?: string;
  title?: string;
  width?: number;
  height?: number;
  rId?: string;
  path?: string;
}

export interface LegacyShapeAttrs {
  shapeType?: string;
  geometryPath?: PresetGeometryPathCommand[] | null;
  children?: string;
  fillPaint?: ShapeFillPaint | null;
  width?: number;
  height?: number;
  fillColor?: string;
  fillType?: string;
  gradientType?: string;
  gradientAngle?: number;
  gradientStops?: string;
  outlineWidth?: number;
  outlineColor?: string;
  outlineStyle?: string;
  transform?: string;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export interface LegacyTextColorAttrs {
  rgb?: string;
  themeColor?: ThemeColorSlot;
  themeTint?: string;
  themeShade?: string;
}

export interface LegacyUnderlineAttrs {
  style?: UnderlineStyle;
  color?: LegacyTextColorAttrs;
}

export interface LegacyFontSizeAttrs {
  size?: number | null;
  sizeCs?: number | null;
}

export interface LegacyFontFamilyAttrs {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  asciiTheme?: string;
  hAnsiTheme?: string;
  eastAsiaTheme?: string;
  csTheme?: string;
}
