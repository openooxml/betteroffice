/**
 * DrawingML chart model for basic rendered charts.
 *
 * This is intentionally normalized for display rather than lossless OOXML
 * round-trip. Unsupported chart families stay in the original package parts.
 */

import type { ImagePosition, ImageSize, ImageWrap } from './image';

export type ChartType = 'bar' | 'column' | 'line' | 'pie' | 'doughnut';

export type ChartGrouping = 'standard' | 'clustered' | 'stacked' | 'percentStacked';

export interface ChartPoint {
  index?: number;
  value?: number;
  category?: string;
  color?: string;
  explosion?: number;
  marker?: { symbol?: string; size?: number; color?: string };
  label?: string;
}

export interface ChartSeries {
  name?: string;
  categories: string[];
  values: number[];
  color?: string;
  index?: number;
  order?: number;
  categoryFormula?: string;
  valueFormula?: string;
  axisIds?: string[];
  points?: ChartPoint[];
  grouping?: ChartGrouping;
  marker?: { symbol?: string; size?: number; color?: string };
  smooth?: boolean;
}

export interface ChartAxis {
  id?: string;
  title?: string;
  min?: number;
  max?: number;
  labels?: string[];
  axisType?: 'category' | 'value' | 'date' | 'series';
  position?: 'left' | 'right' | 'top' | 'bottom';
  crossAxisId?: string;
  crosses?: 'autoZero' | 'min' | 'max' | 'value';
  crossesAt?: number;
  majorUnit?: number;
  minorUnit?: number;
  logarithmicBase?: number;
  reversed?: boolean;
  numberFormat?: string;
  majorTickMark?: string;
  minorTickMark?: string;
  tickLabelPosition?: string;
  hidden?: boolean;
}

/** One chart-family element inside `c:plotArea`. */
export interface ChartPlotGroup {
  chartType?: ChartType | 'area' | 'scatter' | 'radar' | 'stock' | 'bubble' | 'ofPie' | 'surface';
  grouping?: ChartGrouping;
  overlap?: number;
  gapWidth?: number;
  axisIds?: string[];
  series?: ChartSeries[];
  varyColors?: boolean;
  firstSliceAngle?: number;
  holeSize?: number;
  showDataLabels?: boolean;
}

export interface ChartLegend {
  position?: 'left' | 'right' | 'top' | 'bottom';
  visible?: boolean;
}

export interface Chart {
  type: 'chart';
  chartType: ChartType;
  /** Relationship id used by the owning drawing part. */
  rId?: string;
  /** Normalized package path, e.g. word/charts/chart1.xml. */
  path?: string;
  title?: string;
  legend?: ChartLegend;
  series: ChartSeries[];
  axes?: {
    category?: ChartAxis;
    value?: ChartAxis;
  };
  /** Drawing extent from wp:extent in EMUs. */
  size?: ImageSize;
  /** Wrap metadata from the containing drawing; v1 layout treats charts as inline blocks. */
  wrap?: ImageWrap;
  /** Floating anchor. Undefined = inline legacy behavior. */
  position?: ImagePosition;
  /** Multiple/combo plot groups. Undefined = synthesize one from legacy fields. */
  plotGroups?: ChartPlotGroup[];
  /** Complete axis collection keyed through plot-group `axisIds`. */
  axisList?: ChartAxis[];
  /** Chart accessibility description. */
  description?: string;
  /** Decorative flag. Undefined = false. */
  decorative?: boolean;
  /** Stable z-order for anchored charts. Undefined = source order. */
  relativeHeight?: number;
}
