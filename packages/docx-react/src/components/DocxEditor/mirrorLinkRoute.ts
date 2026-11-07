/**
 * Following a hyperlink from the accessibility mirror. A screen reader or the
 * keyboard activates the mirror's anchor, never the canvas, so the click
 * reaches the anchor itself; the route takes it the way the canvas click path
 * takes a link: only a safe href, an internal link moves the caret to its
 * bookmark, an external one opens in a new tab without an opener.
 */

import { createContext, useContext, useEffect, useMemo } from 'react';
import { sanitizeHref } from '@betteroffice/docx/utils';

export interface MirrorLinkRoute {
  followBookmark: ((name: string) => void) | null;
}

export const MirrorLinkRouteContext = createContext<MirrorLinkRoute | null>(null);

/** A route for the editor's mirrors; the editing surface registers its bookmark follower on it. */
export function useMirrorLinkRoute(): MirrorLinkRoute {
  return useMemo(() => ({ followBookmark: null }), []);
}

export function useMirrorBookmarkFollower(follow: (name: string) => void): void {
  const route = useContext(MirrorLinkRouteContext);
  useEffect(() => {
    if (!route) return;
    route.followBookmark = follow;
    return () => {
      if (route.followBookmark === follow) route.followBookmark = null;
    };
  }, [route, follow]);
}

/** Follows the hyperlink anchor a click inside `mirror` landed on; false when it landed on none. */
export function followMirrorLink(mirror: Element, event: MouseEvent, route: MirrorLinkRoute | null): boolean {
  const anchor = event.target instanceof Element ? event.target.closest('a[data-href]') : null;
  if (!anchor || !mirror.contains(anchor)) return false;
  event.preventDefault();
  event.stopPropagation();
  const href = sanitizeHref(anchor.getAttribute('data-href'));
  if (!href) return true;
  if (href.startsWith('#')) route?.followBookmark?.(href.slice(1));
  else window.open(href, '_blank', 'noopener,noreferrer');
  return true;
}
