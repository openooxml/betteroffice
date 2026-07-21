import type { CSSProperties } from 'react';

export type ToolbarIconName =
  | 'search'
  | 'undo'
  | 'redo'
  | 'print'
  | 'formatPaint'
  | 'decimalDecrease'
  | 'decimalIncrease'
  | 'bold'
  | 'italic'
  | 'strikethrough'
  | 'textColor'
  | 'fillColor'
  | 'borders'
  | 'merge'
  | 'alignLeft'
  | 'verticalAlignCenter'
  | 'wrap'
  | 'more'
  | 'chevronDown'
  | 'remove'
  | 'add'
  | 'check'
  | 'save'
  | 'image'
  | 'proposals';

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
      {name === 'search' && (
        <>
          <circle cx="10.5" cy="10.5" r="6" />
          <path d="m15 15 4.5 4.5" />
        </>
      )}
      {name === 'undo' && <path d="m9 7-5 5 5 5M5 12h9a6 6 0 0 1 6 6" />}
      {name === 'redo' && <path d="m15 7 5 5-5 5m4-5h-9a6 6 0 0 0-6 6" />}
      {name === 'print' && (
        <>
          <path d="M7 9V4h10v5M7 18H5a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" />
          <path d="M7 14h10v6H7z" />
          <path d="M17.5 12h.01" />
        </>
      )}
      {name === 'formatPaint' && (
        <>
          <path d="M5 4h11v5H5zM8 9v3h5V9" />
          <path d="M10.5 12v8" />
        </>
      )}
      {name === 'decimalDecrease' && (
        <>
          <path d="M5 7h1M9 7h1M14 7h1M18 7h1" />
          <path d="M5 16h1M9 16h1" />
          <path d="m15 12-4 4 4 4M11 16h8" />
        </>
      )}
      {name === 'decimalIncrease' && (
        <>
          <path d="M5 7h1M9 7h1" />
          <path d="M5 16h1M9 16h1M14 16h1M18 16h1" />
          <path d="m15 3 4 4-4 4M19 7h-8" />
        </>
      )}
      {name === 'bold' && <path d="M7 4h6a4 4 0 0 1 0 8H7zm0 8h7a4 4 0 0 1 0 8H7z" />}
      {name === 'italic' && <path d="M10 4h8M6 20h8M14 4 10 20" />}
      {name === 'strikethrough' && (
        <>
          <path d="M5 12h14M8 7.5A4.5 4.5 0 0 1 12.5 4c2.2 0 3.7 1 4.5 2.5M8 16.5c.8 2.1 2.4 3.5 5 3.5 2.8 0 4.5-1.4 4.5-3.5 0-2.3-2-3.3-5.5-4.5" />
        </>
      )}
      {name === 'textColor' && (
        <>
          <path d="m7 16 5-12 5 12M9 11h6" />
          <path d="M5 20h14" strokeWidth="3" />
        </>
      )}
      {name === 'fillColor' && (
        <>
          <path d="m7 4 10 10-6 6-7-7zM5 12h12" />
          <path d="M18 17c0-1 1-2 1-2s1 1 1 2a1 1 0 0 1-2 0Z" />
        </>
      )}
      {name === 'borders' && (
        <>
          <path d="M4 4h16v16H4zM12 4v16M4 12h16" />
        </>
      )}
      {name === 'merge' && (
        <>
          <path d="M4 5h16v14H4z" />
          <path d="m8 9 3 3-3 3M16 9l-3 3 3 3" />
        </>
      )}
      {name === 'alignLeft' && <path d="M4 6h16M4 10h11M4 14h16M4 18h11" />}
      {name === 'verticalAlignCenter' && (
        <>
          <path d="M4 4h16M4 20h16M8 12h8" />
          <path d="m10 9 2 3 2-3M10 15l2-3 2 3" />
        </>
      )}
      {name === 'wrap' && (
        <>
          <path d="M4 6h16M4 11h12a3 3 0 0 1 0 6h-3" />
          <path d="m15 14-3 3 3 3M4 17h5" />
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
      {name === 'check' && <path d="m5 12 4 4L19 6" />}
      {name === 'save' && <path d="M12 3v11m0 0 4-4m-4 4-4-4M5 19h14" />}
      {name === 'image' && (
        <>
          <rect x="3" y="5" width="18" height="14" rx="2" />
          <circle cx="8.5" cy="9.5" r="1.25" />
          <path d="m6 17 4.5-4.5 3.25 3.25L16 13.5l2 2" />
        </>
      )}
      {name === 'proposals' && (
        <>
          <path d="m12 3 1.25 3.75L17 8l-3.75 1.25L12 13l-1.25-3.75L7 8l3.75-1.25L12 3Z" />
          <path d="m18 14 .75 2.25L21 17l-2.25.75L18 20l-.75-2.25L15 17l2.25-.75L18 14Z" />
        </>
      )}
    </svg>
  );
}
