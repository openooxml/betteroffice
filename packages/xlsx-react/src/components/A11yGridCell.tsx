import type { A11yCell } from '@betteroffice/xlsx';

/**
 * One cell of the offscreen screen-reader grid. A cell with a hyperlink holds
 * a real link a screen reader can follow; following it takes the same path as
 * clicking the link on the canvas, so only targets that path accepts open.
 */
export function A11yGridCell({ cell, onFollowLink }: {
  cell: A11yCell;
  onFollowLink?: (row: number, col: number) => void;
}) {
  return <span role="gridcell" aria-selected={cell.selected}>
    {cell.href && onFollowLink ? <a href={cell.href} tabIndex={-1} onClick={(event) => {
      event.preventDefault();
      onFollowLink(cell.row, cell.col);
    }}>{cell.label}</a> : cell.label}
  </span>;
}
