import type { CSSProperties } from 'react';

export type RibbonIconName = 'undo' | 'redo' | 'delete' | 'add' | 'download' | 'fill' | 'line' | 'weight' | 'pattern' | 'front' | 'forward' | 'backward' | 'back' | 'rotateLeft' | 'rotateRight' | 'flipHorizontal' | 'flipVertical';

export function RibbonIcon({ name, size = 20, style }: { name: RibbonIconName; size?: number; style?: CSSProperties }) {
  const path = {
    undo: 'm9 7-5 5 5 5M5 12h9a6 6 0 0 1 6 6', redo: 'm15 7 5 5-5 5m4-5h-9a6 6 0 0 0-6 6', delete: 'M5 7h14m-9 4v6m4-6v6M9 7l1-3h4l1 3m-8 0 1 13h8l1-13', add: 'M12 5v14M5 12h14', download: 'M12 3v12m-4-4 4 4 4-4M5 20h14', fill: 'm7 4 10 10-5 5-7-7Zm10 13h3', line: 'M5 18h14M8 15l8-8 2 2-8 8H8Z', weight: 'M5 7h14M5 12h14M5 18h14', pattern: 'M4 12h3m2 0h3m2 0h3m2 0h1', front: 'M7 17h10V7M5 13h10V3', forward: 'M6 17h10V7m-5-4 5 4-5 4', backward: 'M18 17H8V7m5-4-5 4 5 4', back: 'M17 17H7V7m12 6H9V3', rotateLeft: 'M7 8V4l-4 4 4 4V8a7 7 0 1 1-1 9', rotateRight: 'M17 8V4l4 4-4 4V8a7 7 0 1 0 1 9', flipHorizontal: 'M12 4v16M7 6l-3 6 3 6m10-12 3 6-3 6', flipVertical: 'M4 12h16M6 7l6-3 6 3m-12 10 6 3 6-3',
  }[name];
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false" style={{ display: 'block', ...style }}><path d={path} /></svg>;
}
