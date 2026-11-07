import { memo, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, MouseEvent, RefObject } from 'react';
import type { SlideDisplayList, TextLinkTarget } from '@betteroffice/pptx';
import { sanitizeHref } from '../../../../shared/safeHref';
import { useTranslation } from '../i18n';
import { activatePeerlessSlide, jumpPeerlessSlide, navigatePeerlessSlide } from './peerlessNavigation';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import {
  sameSlideLink, slideLinkOptions, slideMirrorItems,
  type SlideMirrorItem, type SlideMirrorParagraph, type SlideMirrorSegment,
} from './slideMirror';

const SETTLE_MS = 150;

const visuallyHidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, overflow: 'hidden', clipPath: 'inset(50%)', whiteSpace: 'nowrap',
};

function segmentText(segments: SlideMirrorSegment[], lang: string | null) {
  const parts: SlideMirrorSegment[] = [];
  for (const segment of segments) {
    const previous = parts[parts.length - 1];
    if (previous?.lang === segment.lang) previous.text += segment.text;
    else parts.push({ ...segment });
  }
  if (parts.length === 1 && (!parts[0]!.lang || parts[0]!.lang === lang)) return parts[0]!.text;
  return parts.map((segment, index) => segment.lang && segment.lang !== lang
    ? <span key={index} lang={segment.lang}>{segment.text}</span> : segment.text);
}

function paragraphText(paragraph: SlideMirrorParagraph, run: PeerlessPresentation) {
  const groups: { link?: TextLinkTarget; segments: SlideMirrorSegment[] }[] = [];
  for (const segment of paragraph.segments) {
    let link = segment.link;
    if (link?.kind === 'url' && !sanitizeHref(link.href, slideLinkOptions)) link = undefined;
    if (link?.kind === 'slide' && (!Number.isInteger(link.index) || link.index < 0 || !run.client.state.slides[link.index])) link = undefined;
    const previous = groups[groups.length - 1];
    if (previous && sameSlideLink(previous.link, link)) previous.segments.push(segment);
    else groups.push({ link, segments: [segment] });
  }
  if (groups.length === 1 && !groups[0]!.link) return segmentText(groups[0]!.segments, paragraph.lang);
  return groups.map(({ link, segments }, index) => {
    const text = segmentText(segments, paragraph.lang);
    if (!link) return text;
    return link.kind === 'url'
      ? <a key={index} href={sanitizeHref(link.href, slideLinkOptions)} target="_blank" rel="noopener noreferrer">{text}</a>
      : <a key={index} href="#" data-slide-link={link.kind === 'slide' ? `slide:${link.index}` : `jump:${link.jump}`}>{text}</a>;
  });
}

function followLink(event: MouseEvent<HTMLDivElement>, run: PeerlessPresentation): void {
  const anchor = (event.target as Element).closest('a');
  if (!anchor || !event.currentTarget.contains(anchor)) return;
  event.preventDefault();
  const href = anchor.getAttribute('href');
  const target = anchor.getAttribute('data-slide-link');
  if (target !== null) {
    if (href !== '#') return;
    if (target.startsWith('slide:')) {
      const index = Number(target.slice(6));
      const slide = Number.isInteger(index) && index >= 0 ? run.client.state.slides[index] : undefined;
      if (slide) void activatePeerlessSlide(run, slide.id).catch(() => {});
    } else {
      const jump = target.slice(5);
      if (target.startsWith('jump:') && (jump === 'next' || jump === 'previous' || jump === 'first' || jump === 'last')) {
        jumpPeerlessSlide(run, jump);
      }
    }
    return;
  }
  const safe = sanitizeHref(href, slideLinkOptions);
  if (safe) window.open(safe, '_blank', 'noopener,noreferrer');
}

function MirrorItem({ item, run }: { item: SlideMirrorItem; run: PeerlessPresentation }) {
  const style: CSSProperties = { position: 'absolute', left: item.x, top: item.y, width: item.w, height: item.h, margin: 0 };
  if (item.paragraphs.length === 0) return <div role="img" aria-label={item.label!} style={style} />;
  return <div role={item.label ? 'group' : undefined} aria-label={item.label ?? undefined} style={style}>
    {item.paragraphs.map((paragraph, index) => paragraph.heading
      ? <div key={index} role="heading" aria-level={2} lang={paragraph.lang ?? undefined}>{paragraphText(paragraph, run)}</div>
      : <p key={index} lang={paragraph.lang ?? undefined} style={{ margin: 0 }}>{paragraphText(paragraph, run)}</p>)}
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
    onClick={(event) => followLink(event, run)}
    onFocus={() => setFocused(true)} onBlur={() => setFocused(false)}
    style={{ position: 'absolute', inset: 0, pointerEvents: 'none', outline: focused ? '2px solid #2563eb' : 'none', outlineOffset: -2 }}>
    <div style={{ position: 'absolute', left: 0, top: 0, width: list.width, height: list.height, opacity: 0,
      overflow: 'hidden', userSelect: 'none', transform: `scale(${scale})`, transformOrigin: '0 0' }}>
      {items.map((item) => <MirrorItem key={item.key} item={item} run={run} />)}
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
