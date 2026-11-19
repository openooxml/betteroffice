import type { ReactNode } from 'react';

/** Covers the workspace that contains it while a file opens, without taking space from it; absent for an indicator React renders as nothing. */
export function LoadingLayer({ children }: { children: ReactNode }) {
  if (children == null || typeof children === 'boolean' || children === '') return null;
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
