import type { CSSProperties } from 'react';
import { ToolbarSymbol } from '../../../../../shared/react-toolbar/icons';
import {
  add,
  addBox,
  addPhotoAlternate,
  arrowDropDown,
  arrowSelectorTool,
  borderColor,
  check,
  download,
  flipToBack,
  flipToFront,
  formatAlignCenter,
  formatAlignJustify,
  formatAlignLeft,
  formatAlignRight,
  formatBold,
  formatColorFill,
  formatColorText,
  formatItalic,
  formatUnderlined,
  image,
  insertText,
  lineWeight,
  moreVert,
  moveSelectionDown,
  moveSelectionUp,
  redo,
  remove,
  shapes,
  undo,
} from '../../../../../shared/react-toolbar/symbols';

const SYMBOLS = {
  save: download,
  image,
  insertImage: addPhotoAlternate,
  undo,
  redo,
  newSlide: addBox,
  select: arrowSelectorTool,
  textBox: insertText,
  shape: shapes,
  bringToFront: flipToFront,
  sendToBack: flipToBack,
  bringForward: moveSelectionUp,
  sendBackward: moveSelectionDown,
  fillColor: formatColorFill,
  borderColor,
  borderWidth: lineWeight,
  bold: formatBold,
  italic: formatItalic,
  underline: formatUnderlined,
  textColor: formatColorText,
  alignLeft: formatAlignLeft,
  alignCenter: formatAlignCenter,
  alignRight: formatAlignRight,
  alignJustify: formatAlignJustify,
  more: moreVert,
  chevronDown: arrowDropDown,
  remove,
  add,
  check,
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
