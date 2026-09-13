import type { ParagraphFormatting, Table, TableLook } from '../types/document';
import type { Style } from '../types/styles';
import { mergeParagraphFormatting } from '../utils/paragraphFormattingMerge';

export function tableCellParagraphFormatting(
  table: Table,
  style: Style | undefined,
  rowIndex: number,
  startColumn: number,
  endColumn: number,
  columns: number
): ParagraphFormatting | undefined {
  let result = style?.pPr;
  if (!style?.tblStylePr?.some((part) => part.pPr)) return result;
  const look = table.formatting?.look ?? style.tblPr?.look;
  const flag = (key: Exclude<keyof TableLook, 'value'>, mask: number) =>
    look?.[key] ?? (Number.parseInt(look?.value ?? '0', 16) & mask) !== 0;
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
  const regions: string[] = [];
  if (!flag('noHBand', 0x200) && !atFirstRow && !atLastRow && rowBandSize > 0) {
    regions.push(
      Math.floor((rowIndex - Number(!!firstRow)) / rowBandSize) % 2 === 0
        ? 'band1Horz'
        : 'band2Horz'
    );
  }
  if (!flag('noVBand', 0x400) && !atFirstColumn && !atLastColumn && columnBandSize > 0) {
    regions.push(
      Math.floor((startColumn - Number(!!firstColumn)) / columnBandSize) % 2 === 0
        ? 'band1Vert'
        : 'band2Vert'
    );
  }
  for (const [region, active] of [
    ['firstCol', atFirstColumn],
    ['lastCol', atLastColumn],
    ['firstRow', atFirstRow],
    ['lastRow', atLastRow],
    ['nwCell', atFirstRow && atFirstColumn],
    ['neCell', atFirstRow && atLastColumn],
    ['swCell', atLastRow && atFirstColumn],
    ['seCell', atLastRow && atLastColumn],
  ] as const) {
    if (active) regions.push(region);
  }
  for (const region of regions) {
    result = mergeParagraphFormatting(
      result,
      style.tblStylePr.find((part) => part.type === region)?.pPr
    );
  }
  return result;
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
