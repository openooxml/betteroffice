import type { Layout } from './types';

/**
 * Page index (0-based) whose layout fragments cover `pmPos`, or null if none.
 * Used when the painted DOM may not yet have `[data-doc-start]` for this position (virtualization).
 *
 * Range semantics: `[pmStart, pmEnd)` — half-open, matching ProseMirror's
 * `pos + nodeSize` convention. Boundary positions belong to the next fragment,
 * so when a fragment ends at the same position the next one starts, the next
 * fragment wins (avoids returning the previous page for the start of the
 * next paragraph).
 */
export function findPageIndexContainingPmPos(layout: Layout, pmPos: number): number | null {
  for (let pi = 0; pi < layout.pages.length; pi++) {
    for (const frag of layout.pages[pi].fragments) {
      if (frag.pmStart == null) continue;
      const start = frag.pmStart;
      // Default span of 1 only when pmEnd is missing — matches a caret-only
      // position (cursor between two atoms). Fragments with explicit pmEnd
      // use it as the exclusive upper bound.
      const end = frag.pmEnd ?? start + 1;
      if (pmPos >= start && pmPos < end) {
        return pi;
      }
    }
  }
  return null;
}
