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
  title: string | null;
  richTitle: boolean;
  series: ChartSeriesInfo[];
}

export type ChartSeriesName = { text: string } | { ref: string };
export type ChartEdit = { action: 'set_chart_title'; title: string | null } | { action: 'set_series'; series: number; name?: ChartSeriesName; categories?: string; values?: string };
export type ChartMutation = { sheet: number; chart: number } & ChartEdit;
