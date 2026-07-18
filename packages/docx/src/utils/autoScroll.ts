/**
 * Drag auto-scroll geometry shared by the React and Vue adapters.
 *
 * When the user drag-selects and the pointer nears the top/bottom edge of the
 * scroll container, the container should auto-scroll at a speed proportional to
 * how deep into the edge zone the pointer is. The per-frame loop and the
 * framework wiring stay in each adapter's hook; this is just the pure delta.
 */

/** Pixel distance from the container edge where auto-scroll activates. */
export const AUTO_SCROLL_EDGE_ZONE = 40;
/** Maximum scroll speed in pixels per frame (~60fps). */
export const AUTO_SCROLL_MAX_SPEED = 12;

/**
 * Vertical scroll delta (px/frame) for a pointer at `mouseY` over a container
 * whose viewport spans `rect.top`..`rect.bottom`. Negative scrolls up, positive
 * down, 0 when the pointer is outside both edge zones. Speed ramps linearly
 * from 0 at the edge-zone boundary to `AUTO_SCROLL_MAX_SPEED` at the edge.
 */
export function computeAutoScrollDelta(
  rect: { top: number; bottom: number },
  mouseY: number
): number {
  if (mouseY < rect.top + AUTO_SCROLL_EDGE_ZONE) {
    const proximity = Math.max(0, rect.top + AUTO_SCROLL_EDGE_ZONE - mouseY);
    return -Math.min(
      AUTO_SCROLL_MAX_SPEED,
      (proximity / AUTO_SCROLL_EDGE_ZONE) * AUTO_SCROLL_MAX_SPEED
    );
  }
  if (mouseY > rect.bottom - AUTO_SCROLL_EDGE_ZONE) {
    const proximity = Math.max(0, mouseY - (rect.bottom - AUTO_SCROLL_EDGE_ZONE));
    return Math.min(
      AUTO_SCROLL_MAX_SPEED,
      (proximity / AUTO_SCROLL_EDGE_ZONE) * AUTO_SCROLL_MAX_SPEED
    );
  }
  return 0;
}
