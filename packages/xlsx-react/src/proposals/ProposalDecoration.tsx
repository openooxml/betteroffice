/**
 * Interaction chrome over a proposed cell: a dashed, agent-colored border with
 * a small corner tab as the hover target. The ghost `old → new` pair itself is
 * painted by the engine into the canvas display list, so this layer carries no
 * text. Purely visual — it is `aria-hidden`, and the offscreen a11y grid
 * announces the real committed values only.
 */

import type { Rect } from '@betteroffice/xlsx';

/**
 * Props for {@link ProposalDecoration}.
 */
export interface ProposalDecorationProps {
  rect: Rect;
  color: string;
  /** shown on the corner tab's tooltip: the proposed display text. */
  newText: string;
  /** shown on the corner tab's tooltip: which agent proposed this cell. */
  agentId: string;
}

/**
 * The border + corner-tab chrome over one proposed cell.
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
