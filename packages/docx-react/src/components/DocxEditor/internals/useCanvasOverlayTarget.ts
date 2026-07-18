import { useEffect, useState } from 'react';

/**
 * Resolve the DOM element the interactive comment overlays (sidebar, comment
 * margin markers, floating "Add comment" button) portal into while the
 * display-list interactions are available. The overlays are interactive UI
 * chrome, so they render over the visible canvas pages. `editorContentRef` is the right
 * target: it is the positioned (`position:relative`) ancestor of the
 * `.canvas-pages` host and shares its top-left origin, so the display-list-
 * derived anchor Ys (`computeAnchorPositionsFromDisplayList`, relative to the
 * canvas-pages host) and the `50%`-centered X math line up unchanged. It is
 * NOT a descendant of the canvas host, so overlay clicks never bubble into the
 * host's native pointer listeners (which would move the caret).
 *
 * The caller keeps this active independently of renderer readiness.
 */
export function useCanvasOverlayTarget(
  active: boolean,
  targetRef: React.RefObject<HTMLElement | null> | undefined
): HTMLElement | null {
  const [target, setTarget] = useState<HTMLElement | null>(null);

  useEffect(() => {
    if (!active || !targetRef) {
      setTarget(null);
      return;
    }
    let raf = 0;
    const resolve = () => {
      const el = targetRef.current ?? null;
      setTarget(el);
      // The target host mounts in the same commit, so `resolve` normally lands
      // it immediately; retry a frame in case the ref populates just after.
      if (!el) raf = requestAnimationFrame(resolve);
    };
    resolve();
    return () => cancelAnimationFrame(raf);
  }, [active, targetRef]);

  return target;
}
