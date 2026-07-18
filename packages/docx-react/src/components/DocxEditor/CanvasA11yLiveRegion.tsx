/**
 * Polite live region for the experimental canvas renderer.
 *
 * While pages paint on canvas, caret/selection state has no visible DOM for
 * assistive tech to observe; this region announces the meaningful transitions
 * (selection made/cleared, caret entering commented / tracked-change / list
 * content) derived from the authoritative Yrs selection. Rendered inside the LocaleProvider tree (all strings via
 * `t('a11y.*')`); the host `DocxEditor` — which sits outside that tree —
 * drives it through `notifyRef`, assigned on mount and called from its
 * selection-change handlers.
 */

import { useCallback, useEffect, useRef, useState, type RefObject } from 'react';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { useTranslation } from '../../i18n';
import {
  computeA11yAnnouncements,
  snapshotFromSelectionContext,
  type A11ySelectionSnapshot,
} from '@betteroffice/docx/layout/render';
import { currentYrsSelectionRange } from './yrsCommands';

export interface CanvasA11yLiveRegionProps {
  /** announcements only run while the canvas actually paints */
  active: boolean;
  /** body or HF Yrs session, whichever owns the selection */
  getYrsSession: () => YrsSession | null | undefined;
  /** the host stores the notifier here and calls it on every selection change */
  notifyRef: RefObject<(() => void) | null>;
}

export function CanvasA11yLiveRegion({
  active,
  getYrsSession,
  notifyRef,
}: CanvasA11yLiveRegionProps) {
  const { t } = useTranslation();
  const [announcement, setAnnouncement] = useState('');
  const prevRef = useRef<A11ySelectionSnapshot | null>(null);
  // toggled trailing NBSP so back-to-back identical announcements still
  // register as a DOM change for the live region
  const tickRef = useRef(false);

  const notify = useCallback(() => {
    const session = getYrsSession();
    const yrsContext = (() => {
      try {
        const range = session ? currentYrsSelectionRange(session) : null;
        return session && range ? session.selectionContext(range) : null;
      } catch {
        return null;
      }
    })();
    const numPr = yrsContext?.paragraphProperties.numPr;
    const ctx = yrsContext
      ? {
          hasSelection: yrsContext.hasSelection,
          inInsertion: yrsContext.inInsertion,
          inDeletion: yrsContext.inDeletion,
          inList: numPr != null && typeof numPr === 'object',
          listLevel:
            numPr != null && typeof numPr === 'object'
              ? Number((numPr as { ilvl?: unknown }).ilvl) || 0
              : 0,
          activeCommentIds: [],
        }
      : null;
    if (!ctx) return;
    const snap = snapshotFromSelectionContext(ctx);
    const prev = prevRef.current;
    prevRef.current = snap;
    const messages = computeA11yAnnouncements(prev, snap);
    if (messages.length === 0) return;
    const text = messages.map((m) => t(`a11y.${m.key}`, m.vars)).join(', ');
    tickRef.current = !tickRef.current;
    setAnnouncement(tickRef.current ? `${text} ` : text);
  }, [getYrsSession, t]);

  useEffect(() => {
    if (!active) {
      prevRef.current = null;
      return;
    }
    notifyRef.current = notify;
    return () => {
      notifyRef.current = null;
    };
  }, [active, notify, notifyRef]);

  if (!active) return null;
  return (
    <div className="oox-canvas-live-region" role="status" aria-live="polite">
      {announcement}
    </div>
  );
}
