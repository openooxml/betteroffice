import { memo, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, RefObject } from 'react';
import type { SlideDisplayList } from '@betteroffice/pptx';
import { useTranslation } from '../i18n';
import { navigatePeerlessSlide } from './peerlessNavigation';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { slideMirrorItems, type SlideMirrorItem } from './slideMirror';

const SETTLE_MS = 150;

const visuallyHidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, overflow: 'hidden', clipPath: 'inset(50%)', whiteSpace: 'nowrap',
};

function MirrorItem({ item }: { item: SlideMirrorItem }) {
  const style: CSSProperties = { position: 'absolute', left: item.x, top: item.y, width: item.w, height: item.h, margin: 0 };
  if (item.paragraphs.length === 0) return <div role="img" aria-label={item.label!} style={style} />;
  return <div role={item.label ? 'group' : undefined} aria-label={item.label ?? undefined} style={style}>
    {item.paragraphs.map((paragraph, index) => paragraph.heading
      ? <div key={index} role="heading" aria-level={2}>{paragraph.text}</div>
      : <p key={index} style={{ margin: 0 }}>{paragraph.text}</p>)}
  </div>;
}

/**
 * The current slide's content for assistive technology, laid over the canvas in
 * slide pixels and never painted. It is a document inside the stage's
 * application so a screen reader can browse it once focus lands here.
 */
export const PeerlessSlideMirror = memo(function PeerlessSlideMirror({ run, list, label, scale }: {
  run: PeerlessPresentation;
  list: SlideDisplayList;
  label: string;
  scale: number;
}) {
  const { t } = useTranslation();
  const [focused, setFocused] = useState(false);
  const picture = t('objects.defaultPictureName');
  const items = useMemo(() => slideMirrorItems(list, picture), [list, picture]);
  return <div role="document" tabIndex={0} aria-label={label} data-testid="pptx-slide-mirror"
    onKeyDown={(event) => navigatePeerlessSlide(event, run)}
    onFocus={() => setFocused(true)} onBlur={() => setFocused(false)}
    style={{ position: 'absolute', inset: 0, pointerEvents: 'none', outline: focused ? '2px solid #2563eb' : 'none', outlineOffset: -2 }}>
    <div style={{ position: 'absolute', left: 0, top: 0, width: list.width, height: list.height, opacity: 0,
      overflow: 'hidden', userSelect: 'none', transform: `scale(${scale})`, transformOrigin: '0 0' }}>
      {items.map((item) => <MirrorItem key={item.key} item={item} />)}
    </div>
  </div>;
});

/**
 * Speaks slide and shape changes while focus is in the editor. Quick changes
 * collapse into the one the user stops on, and a slide change outranks the
 * selection reset that follows it.
 */
export const PeerlessSlideAnnouncer = memo(function PeerlessSlideAnnouncer({ slideKey, slideText, selectionKey, selectionText, scope }: {
  slideKey: string;
  slideText: string;
  selectionKey: string;
  selectionText: string | null;
  scope: RefObject<HTMLElement | null>;
}) {
  const [announcement, setAnnouncement] = useState('');
  const latest = useRef({ slideText, selectionText });
  latest.current = { slideText, selectionText };
  const tick = useRef(false);
  const seen = useRef<{ slideKey: string; selectionKey: string } | null>(null);
  const pending = useRef<'slide' | 'selection' | null>(null);
  useEffect(() => {
    const before = seen.current;
    seen.current = { slideKey, selectionKey };
    if (!before) return;
    if (before.slideKey !== slideKey) pending.current = 'slide';
    else if (before.selectionKey !== selectionKey) pending.current ??= 'selection';
    if (!pending.current) return;
    const timer = setTimeout(() => {
      const text = pending.current === 'slide' ? latest.current.slideText : latest.current.selectionText;
      pending.current = null;
      const root = scope.current;
      if (!text || !root?.contains(root.ownerDocument.activeElement)) return;
      tick.current = !tick.current;
      setAnnouncement(tick.current ? `${text}\u00a0` : text);
    }, SETTLE_MS);
    return () => clearTimeout(timer);
  }, [slideKey, selectionKey, scope]);
  return <span role="status" aria-live="polite" data-testid="pptx-live-region" style={visuallyHidden}>{announcement}</span>;
});
