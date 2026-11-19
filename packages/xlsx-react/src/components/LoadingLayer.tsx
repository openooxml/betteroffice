import type { CSSProperties, ReactNode } from 'react';

/** The editor-owned box the layer covers: a host `style` on the editor root cannot unposition it. */
export const loadingAnchorStyle: CSSProperties = {
  position: 'relative',
  display: 'flex',
  flexDirection: 'column',
  flex: 1,
  minWidth: 0,
  minHeight: 0,
};

/** Covers the editor while a file opens, without taking space from the workspace; absent for an indicator React renders as nothing. */
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
