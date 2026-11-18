import type { ReactNode } from 'react';

/** Covers the editor root while a file opens, without taking space from the workspace. */
export function LoadingLayer({ children }: { children: ReactNode }) {
  return (
    <div
      data-testid="xlsx-loading"
      role="status"
      style={{
        position: 'absolute',
        inset: 0,
        display: 'flex',
        flexDirection: 'column',
        pointerEvents: 'none',
      }}
    >
      {children}
    </div>
  );
}
