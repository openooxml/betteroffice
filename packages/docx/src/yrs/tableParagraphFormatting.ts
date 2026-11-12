import type {
  ParagraphFormatting,
  Table,
  TableBorders,
  TableCellFormatting,
  TableLook,
  TextFormatting,
} from '../types/document';
import type { Style } from '../types/styles';
import { mergeParagraphFormatting } from '../utils/paragraphFormattingMerge';
import { mergeTextFormatting } from '../utils/textFormattingMerge';

type RegionEdges = [top: boolean, bottom: boolean, left: boolean, right: boolean];

export function tableCellParagraphFormatting(
  table: Table,
  style: Style | undefined,
  rowIndex: number,
  startColumn: number,
  endColumn: number,
  columns: number
): ParagraphFormatting | undefined {
  return tableCellStyleFormatting(table, style, rowIndex, startColumn, endColumn, columns, 'pPr');
}

export function tableCellRunFormatting(
  table: Table,
  style: Style | undefined,
  rowIndex: number,
  startColumn: number,
  endColumn: number,
  columns: number
): TextFormatting | undefined {
  return tableCellStyleFormatting(table, style, rowIndex, startColumn, endColumn, columns, 'rPr');
}

/** The shading and edge borders a cell takes from its table style's `tcPr` and active conditional parts. */
export function tableCellStyleProperties(
  table: Table,
  style: Style | undefined,
  rowIndex: number,
  startColumn: number,
  endColumn: number,
  columns: number
): Pick<TableCellFormatting, 'shading' | 'borders'> | undefined {
  if (!style?.tcPr && !style?.tblStylePr?.some((part) => part.tcPr)) return undefined;
  const layers: Array<[TableCellFormatting | undefined, RegionEdges]> = [
    [style.tcPr, [true, true, true, true]],
    ...tableStyleRegions(table, style, rowIndex, startColumn, endColumn, columns).map(
      ([region, edges]): [TableCellFormatting | undefined, RegionEdges] => [
        style.tblStylePr?.find((part) => part.type === region)?.tcPr,
        edges,
      ]
    ),
  ];
  let shading: TableCellFormatting['shading'];
  const borders: TableBorders = {};
  for (const [properties, [top, bottom, left, right]] of layers) {
    if (properties?.shading) shading = properties.shading;
    const sides = properties?.borders;
    if (!sides) continue;
    for (const [edge, side] of [
      ['top', top ? 'top' : 'insideH'],
      ['bottom', bottom ? 'bottom' : 'insideH'],
      ['left', left ? 'left' : 'insideV'],
      ['right', right ? 'right' : 'insideV'],
    ] as const) {
      if (sides[side]) borders[edge] = sides[side];
    }
  }
  const hasBorders = Object.keys(borders).length > 0;
  if (!shading && !hasBorders) return undefined;
  return { shading, borders: hasBorders ? borders : undefined };
}

function tableCellStyleFormatting<K extends 'pPr' | 'rPr'>(
  table: Table,
  style: Style | undefined,
  rowIndex: number,
  startColumn: number,
  endColumn: number,
  columns: number,
  property: K
): Style[K] {
  const merge = (target: Style[K], source: Style[K]): Style[K] =>
    (property === 'rPr'
      ? mergeTextFormatting(target as TextFormatting | undefined, source as TextFormatting | undefined)
      : mergeParagraphFormatting(target as ParagraphFormatting | undefined, source as ParagraphFormatting | undefined)) as Style[K];
  let result = style?.[property];
  if (!style?.tblStylePr?.some((part) => part[property])) return result;
  for (const [region] of tableStyleRegions(table, style, rowIndex, startColumn, endColumn, columns)) {
    if (region === 'wholeTable' && property !== 'rPr') continue;
    result = merge(
      result,
      style.tblStylePr.find((part) => part.type === region)?.[property]
    );
  }
  return result;
}

/** The table style regions active for a cell in precedence order, each with whether the cell sits on the region's top, bottom, left and right edge. */
function tableStyleRegions(
  table: Table,
  style: Style,
  rowIndex: number,
  startColumn: number,
  endColumn: number,
  columns: number
): Array<[string, RegionEdges]> {
  const look = table.formatting?.look ?? style.tblPr?.look;
  const flag = (key: Exclude<keyof TableLook, 'value'>, mask: number) =>
    look?.[key] ?? (Number.parseInt(look?.value ?? (look ? '0' : '04A0'), 16) & mask) !== 0;
  const firstRow = flag('firstRow', 0x20);
  const lastRow = flag('lastRow', 0x40);
  const firstColumn = flag('firstColumn', 0x80);
  const lastColumn = flag('lastColumn', 0x100);
  const gridBefore = table.rows[rowIndex].formatting?.gridBefore ?? 0;
  startColumn += gridBefore;
  endColumn += gridBefore;
  const atFirstRow = !!firstRow && rowIndex === 0;
  const atLastRow = !!lastRow && rowIndex === table.rows.length - 1;
  const atFirstColumn = !!firstColumn && startColumn === 0;
  const atLastColumn = !!lastColumn && endColumn === columns;
  const rowBandSize = table.formatting?.styleRowBandSize ?? style.tblPr?.styleRowBandSize ?? 1;
  const columnBandSize = table.formatting?.styleColBandSize ?? style.tblPr?.styleColBandSize ?? 1;
  const top = rowIndex === 0;
  const bottom = rowIndex === table.rows.length - 1;
  const left = startColumn === 0;
  const right = endColumn === columns;
  const bandOf = (position: number, size: number) => Math.floor(position / size);
  const regions: Array<[string, RegionEdges]> = [['wholeTable', [top, bottom, left, right]]];
  if (!flag('noHBand', 0x200) && !atFirstRow && !atLastRow && rowBandSize > 0) {
    const position = Math.max(0, rowIndex - Number(!!firstRow));
    const band = bandOf(position, rowBandSize);
    const bodyEnd = table.rows.length - Number(!!lastRow);
    regions.push([
      band % 2 === 0 ? 'band1Horz' : 'band2Horz',
      [
        position === 0 || bandOf(position - 1, rowBandSize) !== band,
        rowIndex + 1 >= bodyEnd || bandOf(position + 1, rowBandSize) !== band,
        left,
        right,
      ],
    ]);
  }
  if (!flag('noVBand', 0x400) && !atFirstColumn && !atLastColumn && columnBandSize > 0) {
    const position = Math.max(0, startColumn - Number(!!firstColumn));
    const band = bandOf(position, columnBandSize);
    const last = Math.max(0, endColumn - Number(!!firstColumn) - 1);
    const bodyEnd = columns - Number(!!lastColumn);
    regions.push([
      band % 2 === 0 ? 'band1Vert' : 'band2Vert',
      [
        top,
        bottom,
        position === 0 || bandOf(position - 1, columnBandSize) !== band,
        endColumn >= bodyEnd || bandOf(last + 1, columnBandSize) !== bandOf(last, columnBandSize),
      ],
    ]);
  }
  for (const [region, active, edges] of [
    ['firstCol', atFirstColumn, [top, bottom, true, true]],
    ['lastCol', atLastColumn, [top, bottom, true, true]],
    ['firstRow', atFirstRow, [true, true, left, right]],
    ['lastRow', atLastRow, [true, true, left, right]],
    ['nwCell', atFirstRow && atFirstColumn, [true, true, true, true]],
    ['neCell', atFirstRow && atLastColumn, [true, true, true, true]],
    ['swCell', atLastRow && atFirstColumn, [true, true, true, true]],
    ['seCell', atLastRow && atLastColumn, [true, true, true, true]],
  ] as const) {
    if (active) regions.push([region, [...edges]]);
  }
  return regions;
}

export function tableColumnCount(table: Table): number {
  return Math.max(
    table.columnWidths?.length ?? 0,
    ...table.rows.map((row) =>
      row.cells.reduce(
        (sum, cell) => sum + (cell.formatting?.gridSpan ?? 1),
        (row.formatting?.gridBefore ?? 0) + (row.formatting?.gridAfter ?? 0)
      )
    )
  );
}
