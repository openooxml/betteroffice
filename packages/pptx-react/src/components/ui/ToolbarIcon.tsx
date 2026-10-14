import type { CSSProperties } from 'react';
import { ToolbarSymbol, type ToolbarSymbolName } from '../../../../../shared/react-toolbar/icons';

const SYMBOLS = {
  save: 'download',
  image: 'image',
  insertImage: 'add_photo_alternate',
  undo: 'undo',
  redo: 'redo',
  newSlide: 'add_box',
  select: 'arrow_selector_tool',
  textBox: 'insert_text',
  shape: 'shapes',
  bringToFront: 'flip_to_front',
  sendToBack: 'flip_to_back',
  bringForward: 'move_selection_up',
  sendBackward: 'move_selection_down',
  fillColor: 'format_color_fill',
  borderColor: 'border_color',
  borderWidth: 'line_weight',
  bold: 'format_bold',
  italic: 'format_italic',
  underline: 'format_underlined',
  textColor: 'format_color_text',
  alignLeft: 'format_align_left',
  alignCenter: 'format_align_center',
  alignRight: 'format_align_right',
  alignJustify: 'format_align_justify',
  more: 'more_vert',
  chevronDown: 'arrow_drop_down',
  remove: 'remove',
  add: 'add',
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
