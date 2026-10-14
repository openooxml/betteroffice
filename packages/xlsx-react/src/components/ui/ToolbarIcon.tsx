import type { CSSProperties } from 'react';
import { ToolbarSymbol, type ToolbarSymbolName } from '../../../../../shared/react-toolbar/icons';

const SYMBOLS = {
  search: 'search',
  undo: 'undo',
  redo: 'redo',
  print: 'print',
  formatPaint: 'format_paint',
  decimalDecrease: 'decimal_decrease',
  decimalIncrease: 'decimal_increase',
  bold: 'format_bold',
  italic: 'format_italic',
  strikethrough: 'strikethrough_s',
  textColor: 'format_color_text',
  fillColor: 'format_color_fill',
  borders: 'border_all',
  borderColor: 'border_color',
  merge: 'cell_merge',
  alignLeft: 'format_align_left',
  verticalAlignCenter: 'vertical_align_center',
  wrap: 'wrap_text',
  more: 'more_vert',
  chevronDown: 'arrow_drop_down',
  remove: 'remove',
  add: 'add',
  check: 'check',
  save: 'download',
  image: 'image',
  proposals: 'rate_review',
} as const satisfies Record<string, ToolbarSymbolName>;

export type ToolbarIconName = keyof typeof SYMBOLS;

export interface ToolbarIconProps {
  name: ToolbarIconName;
  size?: number;
  style?: CSSProperties;
}

export function ToolbarIcon({ name, size, style }: ToolbarIconProps) {
  return <ToolbarSymbol name={SYMBOLS[name]} size={size} style={style} />;
}
