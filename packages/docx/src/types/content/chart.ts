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

/** Run properties from a `c:txPr`; unset fields inherit from the chart. */
export interface ChartTextProperties {
  font?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

/** One `c:dLbl`: an index plus the switches it overrides. */
export interface ChartPointLabel {
  index?: number;
  text?: string;
  labels: ChartDataLabels;
}

/** A `c:dLbls`; every switch is optional so an unset one inherits. */
export interface ChartDataLabels {
  delete?: boolean;
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showPercent?: boolean;
  showLegendKey?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
  position?: string;
  numberFormat?: string;
  text?: ChartTextProperties;
  points?: ChartPointLabel[];
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
  /** `c:xVal` of a scatter or bubble series. */
  xValues?: number[];
  /** `c:bubbleSize` of a bubble series. */
  bubbleSizes?: number[];
  dataLabels?: ChartDataLabels;
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
  majorGridlines?: boolean;
  minorGridlines?: boolean;
  text?: ChartTextProperties;
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
  /** `c:scatterStyle`. */
  scatterStyle?: 'none' | 'line' | 'lineMarker' | 'marker' | 'smooth' | 'smoothMarker';
  /** `c:radarStyle`. */
  radarStyle?: 'standard' | 'marker' | 'filled';
  /** `c:bubbleScale`, a percentage of the default bubble size. */
  bubbleScale?: number;
  /** `c:sizeRepresents`. */
  sizeRepresents?: 'area' | 'w';
  /** `c:wireframe` of a surface chart. */
  wireframe?: boolean;
  hiLowLines?: boolean;
  upDownBars?: boolean;
  /** `c:marker` of a line chart, which switches every series marker off. */
  marker?: boolean;
  dataLabels?: ChartDataLabels;
}

export interface ChartLegend {
  position?: 'left' | 'right' | 'top' | 'bottom';
  visible?: boolean;
  /** `c:txPr` on the legend; unset fields inherit from the chart. */
  text?: ChartTextProperties;
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
  /** `c:txPr` on `c:chartSpace`: the root of chart text inheritance. */
  text?: ChartTextProperties;
  /** `c:txPr` on `c:title`. */
  titleText?: ChartTextProperties;
}
