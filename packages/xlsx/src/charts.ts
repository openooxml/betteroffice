export interface ChartSeriesInfo {
  index: number;
  name: string | null;
  nameRef: string | null;
  categoryRef: string | null;
  valueRef: string | null;
}

export interface ChartInfo {
  id: string;
  index: number;
  part: string;
  type: string;
  anchorRange: string | null;
  /** false when the chart part cannot be read; error says why and the rest is empty. */
  editable: boolean;
  error: string | null;
  title: string | null;
  /** a title Excel fills in, with no text of its own. */
  autoTitle: boolean;
  richTitle: boolean;
  series: ChartSeriesInfo[];
}

/**
 * How far each edge of a chart frame moves, in content pixels: a positive
 * value moves the left or right edge right and the top or bottom edge down.
 * The engine computes the new anchor and clamps the result to the grid.
 */
export interface ChartResizeEdges {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export type ChartSeriesName = { text: string } | { ref: string };
export type ChartEdit = { action: 'set_chart_title'; title: string | null } | { action: 'set_series'; series: number; name?: ChartSeriesName; categories?: string; values?: string };
export type ChartMutation = { sheet: number; chart: number } & ChartEdit;
