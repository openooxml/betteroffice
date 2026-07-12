/**
 * Deterministic per-agent color for proposal decorations. A stable hash over the
 * agentId picks from a small, visually distinct palette so every cell an agent
 * proposes shares one hue and two agents rarely collide. Chrome-only styling.
 */

// distinct, mid-saturation hues that read on a white grid; the ghost text and
// dashed border both use the agent's entry.
const PALETTE = [
  '#7c3aed', // violet
  '#0891b2', // cyan
  '#c2410c', // orange
  '#4d7c0f', // olive
  '#be185d', // magenta
  '#1d4ed8', // blue
];

// fnv-1a over the agentId, folded to a palette index — same agent, same color.
function hashAgent(agentId: string): number {
  let hash = 0x811c9dc5;
  for (let i = 0; i < agentId.length; i++) {
    hash ^= agentId.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0) % PALETTE.length;
}

/**
 * The palette color for an agentId's proposal decorations.
 */
export function proposalColor(agentId: string): string {
  return PALETTE[hashAgent(agentId)];
}
