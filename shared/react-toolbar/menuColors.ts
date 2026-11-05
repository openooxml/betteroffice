import { toolbarTheme } from './theme';
import type { OverflowMenuColors } from './MenuList';

export const DEFAULT_MENU_COLORS: OverflowMenuColors = {
  surface: toolbarTheme.surface,
  text: toolbarTheme.ink,
  mutedText: toolbarTheme.inkMuted,
  border: toolbarTheme.line,
  hover: toolbarTheme.hover,
  shadow: 'rgba(0, 0, 0, 0.12)',
  focus: toolbarTheme.accent,
};
