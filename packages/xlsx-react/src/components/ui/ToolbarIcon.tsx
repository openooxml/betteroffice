import type { CSSProperties } from 'react';
import { ToolbarSymbol } from '../../../../../shared/react-toolbar/icons';
import {
  add,
  arrowDropDown,
  borderAll,
  borderColor,
  cellMerge,
  check,
  decimalDecrease,
  decimalIncrease,
  download,
  formatAlignLeft,
  formatBold,
  formatColorFill,
  formatColorText,
  formatItalic,
  formatPaint,
  image,
  moreVert,
  print,
  rateReview,
  redo,
  remove,
  search,
  strikethroughS,
  undo,
  verticalAlignCenter,
  wrapText,
} from '../../../../../shared/react-toolbar/symbols';

const SYMBOLS = {
  search,
  undo,
  redo,
  print,
  formatPaint,
  decimalDecrease,
  decimalIncrease,
  bold: formatBold,
  italic: formatItalic,
  strikethrough: strikethroughS,
  textColor: formatColorText,
  fillColor: formatColorFill,
  borders: borderAll,
  borderColor,
  merge: cellMerge,
  alignLeft: formatAlignLeft,
  verticalAlignCenter,
  wrap: wrapText,
  more: moreVert,
  chevronDown: arrowDropDown,
  remove,
  add,
  check,
  save: download,
  image,
  proposals: rateReview,
};

export type ToolbarIconName = keyof typeof SYMBOLS;

export interface ToolbarIconProps {
  name: ToolbarIconName;
  size?: number;
  style?: CSSProperties;
}

export function ToolbarIcon({ name, size, style }: ToolbarIconProps) {
  return <ToolbarSymbol path={SYMBOLS[name]} size={size} style={style} />;
}
