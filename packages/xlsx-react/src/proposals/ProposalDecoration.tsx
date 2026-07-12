/**
 * A single ghost decoration over a proposed cell: a dashed, agent-colored border
 * and the semi-transparent `newText` preview, positioned from the display-list
 * grid geometry (same source as the selection overlay). Purely visual — it is
 * `aria-hidden`, and the offscreen a11y grid announces the real committed values
 * only, never these previews. A small corner tab is the one hover target.
 */

import type { Rect } from '@betteroffice/xlsx';

/**
 * Props for {@link ProposalDecoration}.
 */
export interface ProposalDecorationProps {
  rect: Rect;
  color: string;
  newText: string;
  /** shown on the corner tab's tooltip: which agent proposed this cell. */
  agentId: string;
}

/**
 * The ghost preview painted over one proposed cell.
 */
export function ProposalDecoration({ rect, color, newText, agentId }: ProposalDecorationProps) {
  return (
    <div
      data-testid="xlsx-proposal-decoration"
      aria-hidden="true"
      style={{
        position: 'absolute',
        left: rect.x,
        top: rect.y,
        width: rect.w,
        height: rect.h,
        boxSizing: 'border-box',
        border: `1.5px dashed ${color}`,
        pointerEvents: 'none',
        overflow: 'hidden',
      }}
    >
      <span
        style={{
          position: 'absolute',
          left: 0,
          right: 0,
          bottom: 1,
          textAlign: 'center',
          opacity: 0.6,
          color,
          font: '11px system-ui, sans-serif',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          padding: '0 2px',
        }}
      >
        {newText}
      </span>
      <span
        title={`${agentId}: ${newText}`}
        style={{
          position: 'absolute',
          top: 0,
          right: 0,
          width: 6,
          height: 6,
          background: color,
          pointerEvents: 'auto',
        }}
      />
    </div>
  );
}
