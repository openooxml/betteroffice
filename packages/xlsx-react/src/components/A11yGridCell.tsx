import type { CSSProperties } from 'react';
import type { A11yCell, A11yGrid } from '@betteroffice/xlsx';

/**
 * One cell of the offscreen screen-reader grid. A cell with a hyperlink holds
 * a real link a screen reader can follow; following it takes the same path as
 * clicking the link on the canvas, so only targets that path accepts open.
 */
export function A11yGridCell({ cell, onFollowLink }: {
  cell: A11yCell;
  onFollowLink?: (row: number, col: number) => void;
}) {
  return <span role="gridcell" aria-colindex={cell.col + 2} aria-selected={cell.selected}>
    {cell.href && onFollowLink ? <a href={cell.href} tabIndex={-1} onClick={(event) => {
      event.preventDefault();
      onFollowLink(cell.row, cell.col);
    }}>{cell.label}</a> : cell.label}
  </span>;
}

/**
 * The offscreen grid holds only the painted viewport, so each row and cell says
 * where it sits in the sheet: the header row and column come first, then sheet
 * row and column numbers, and the sheet's own size is left unknown as -1.
 */
export function A11yGridTable({ grid, style, onFollowLink }: {
  grid: A11yGrid;
  style: CSSProperties;
  onFollowLink?: (row: number, col: number) => void;
}) {
  return <div style={style} role="grid" aria-label={grid.label} aria-rowcount={-1} aria-colcount={-1}>
    <div role="row" aria-rowindex={1}><span role="gridcell" aria-colindex={1} />{grid.columnHeaders.map((header) =>
      <span key={header.col} role="columnheader" aria-colindex={header.col + 2}>{header.label}</span>)}</div>
    {grid.rows.map((row) => <div key={row.row} role="row" aria-rowindex={row.row + 2}>
      <span role="rowheader" aria-colindex={1}>{row.header}</span>
      {row.cells.map((cell) => <A11yGridCell key={cell.col} cell={cell} onFollowLink={onFollowLink} />)}
    </div>)}
  </div>;
}
