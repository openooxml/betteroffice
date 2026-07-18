/** Public table property/query compatibility helpers; Rust owns table stories. */

import type { TableFormatting } from '../types/document';
import {
  findChild,
  getAttribute,
  parseBooleanElement,
  parseNumericAttribute,
  type XmlElement,
} from './xmlParser';
import {
  parseCellMargins,
  parseFloatingTableProperties,
  parseShading,
  parseTableBorders,
  parseTableLook,
  parseWidth,
} from './tableParser/properties';

export {
  parseTableMeasurement,
  parseBorderSpec,
  parseTableBorders,
  parseCellMargins,
  parseShading,
  parseTableLook,
  parseFloatingTableProperties,
} from './tableParser/properties';
export {
  getTableColumnCount,
  getTableRowCount,
  isCellMergeContinuation,
  isCellMergeStart,
  isCellHorizontallyMerged,
  getTableText,
  hasHeaderRow,
  getHeaderRows,
  isFloatingTable,
} from './tableParser/queries';

export function parseTableProperties(element: XmlElement | null): TableFormatting | undefined {
  if (!element) return undefined;
  const formatting: TableFormatting = {};
  const width = parseWidth(findChild(element, 'w', 'tblW'));
  if (width) formatting.width = width;
  const justification = getAttribute(findChild(element, 'w', 'jc'), 'w', 'val');
  if (
    justification === 'left' ||
    justification === 'center' ||
    justification === 'right' ||
    justification === 'start'
  ) {
    formatting.justification = justification === 'start' ? 'left' : justification;
  }
  const cellSpacing = parseWidth(findChild(element, 'w', 'tblCellSpacing'));
  if (cellSpacing) formatting.cellSpacing = cellSpacing;
  const indent = parseWidth(findChild(element, 'w', 'tblInd'));
  if (indent) formatting.indent = indent;
  const borders = parseTableBorders(findChild(element, 'w', 'tblBorders'));
  if (borders) formatting.borders = borders;
  const cellMargins = parseCellMargins(findChild(element, 'w', 'tblCellMar'));
  if (cellMargins) formatting.cellMargins = cellMargins;
  const layout = getAttribute(findChild(element, 'w', 'tblLayout'), 'w', 'type');
  if (layout === 'fixed' || layout === 'autofit') formatting.layout = layout;
  const styleId = getAttribute(findChild(element, 'w', 'tblStyle'), 'w', 'val');
  if (styleId) formatting.styleId = styleId;
  const rowBandSize = parseNumericAttribute(
    findChild(element, 'w', 'tblStyleRowBandSize'),
    'w',
    'val'
  );
  if (rowBandSize !== undefined && rowBandSize > 0 && rowBandSize <= 1024) {
    formatting.styleRowBandSize = rowBandSize;
  }
  const colBandSize = parseNumericAttribute(
    findChild(element, 'w', 'tblStyleColBandSize'),
    'w',
    'val'
  );
  if (colBandSize !== undefined && colBandSize > 0 && colBandSize <= 1024) {
    formatting.styleColBandSize = colBandSize;
  }
  const look = parseTableLook(findChild(element, 'w', 'tblLook'));
  if (look) formatting.look = look;
  const shading = parseShading(findChild(element, 'w', 'shd'));
  if (shading) formatting.shading = shading;
  const overlap = getAttribute(findChild(element, 'w', 'tblOverlap'), 'w', 'val');
  if (overlap === 'never' || overlap === 'overlap') formatting.overlap = overlap;
  const floating = parseFloatingTableProperties(findChild(element, 'w', 'tblpPr'));
  if (floating) formatting.floating = floating;
  if (parseBooleanElement(findChild(element, 'w', 'bidiVisual'))) formatting.bidi = true;
  return Object.keys(formatting).length ? formatting : undefined;
}
