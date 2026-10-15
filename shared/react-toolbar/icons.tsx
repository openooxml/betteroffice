import type { CSSProperties } from 'react';
import { TOOLBAR_ICON_SIZE } from './theme';

export interface ToolbarSymbolProps {
  /** A path from `./symbols`, drawn in Material Symbols' 960-unit box. */
  path: string;
  size?: number;
  style?: CSSProperties;
}

export function ToolbarSymbol({ path, size = TOOLBAR_ICON_SIZE, style }: ToolbarSymbolProps) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 -960 960 960"
      fill="currentColor"
      aria-hidden="true"
      focusable="false"
      style={{ display: 'inline-flex', flexShrink: 0, ...style }}
    >
      <path d={path} />
    </svg>
  );
}
