import type { CSSProperties } from 'react';

export type ToolbarIconName =
  | 'undo'
  | 'redo'
  | 'newSlide'
  | 'select'
  | 'textBox'
  | 'shape'
  | 'fillColor'
  | 'borderColor'
  | 'borderWidth'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'textColor'
  | 'more'
  | 'chevronDown'
  | 'remove'
  | 'add';

export interface ToolbarIconProps {
  name: ToolbarIconName;
  size?: number;
  style?: CSSProperties;
}

export function ToolbarIcon({ name, size = 20, style }: ToolbarIconProps) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
      style={{ display: 'inline-flex', flexShrink: 0, ...style }}
    >
      {name === 'undo' && <path d="m9 7-5 5 5 5M5 12h9a6 6 0 0 1 6 6" />}
      {name === 'redo' && <path d="m15 7 5 5-5 5m4-5h-9a6 6 0 0 0-6 6" />}
      {name === 'newSlide' && (
        <>
          <rect x="3" y="5" width="15" height="14" rx="1.5" />
          <path d="M7 9h7M7 13h5M20 9v6m-3-3h6" />
        </>
      )}
      {name === 'select' && <path d="m6 3 12 9-6 1.5L9 19Z" />}
      {name === 'textBox' && (
        <>
          <rect x="3" y="4" width="18" height="16" rx="1.5" />
          <path d="M8 8h8m-4 0v8m-3 0h6" />
        </>
      )}
      {name === 'shape' && <rect x="4" y="6" width="16" height="12" rx="3" />}
      {name === 'fillColor' && (
        <>
          <path d="m7 4 10 10-5 5-7-7Z" />
          <path d="m5 12 7-7m6 12h2" />
        </>
      )}
      {name === 'borderColor' && (
        <>
          <path d="M5 18h14M8 15l8-8 2 2-8 8H8Z" />
        </>
      )}
      {name === 'borderWidth' && (
        <>
          <path d="M5 7h14" strokeWidth="1" />
          <path d="M5 12h14" strokeWidth="2" />
          <path d="M5 18h14" strokeWidth="3" />
        </>
      )}
      {name === 'bold' && <path d="M7 4h6a4 4 0 0 1 0 8H7zm0 8h7a4 4 0 0 1 0 8H7z" />}
      {name === 'italic' && <path d="M10 4h8M6 20h8M14 4 10 20" />}
      {name === 'underline' && (
        <>
          <path d="M7 4v7a5 5 0 0 0 10 0V4" />
          <path d="M5 20h14" />
        </>
      )}
      {name === 'textColor' && (
        <>
          <path d="m7 16 5-12 5 12M9 11h6" />
          <path d="M5 20h14" strokeWidth="3" />
        </>
      )}
      {name === 'more' && (
        <>
          <circle cx="5" cy="12" r="1" fill="currentColor" stroke="none" />
          <circle cx="12" cy="12" r="1" fill="currentColor" stroke="none" />
          <circle cx="19" cy="12" r="1" fill="currentColor" stroke="none" />
        </>
      )}
      {name === 'chevronDown' && <path d="m7 10 5 5 5-5" />}
      {name === 'remove' && <path d="M5 12h14" />}
      {name === 'add' && <path d="M12 5v14M5 12h14" />}
    </svg>
  );
}
