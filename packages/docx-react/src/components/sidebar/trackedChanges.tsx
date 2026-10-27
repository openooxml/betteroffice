import { createContext, useEffect, useState, type CSSProperties, type ReactNode } from 'react';
import { resolvePresenceColor, type CollaborationPeer, type CollaborationPresence, type CollaborationUser } from '@betteroffice/docx/collaboration';
import type { ReactSidebarItem, SidebarItemRenderProps } from '../../plugin-api/types';
import { getInitials, type TrackedChangeEntry } from './cardUtils';

/** Review author category. */
export type TrackedChangeAuthorKind = 'human' | 'agent';

/** Optional author appearance overrides. */
export interface TrackedChangeAuthorAppearance {
  icon?: ReactNode;
  color?: string;
  kind?: TrackedChangeAuthorKind;
}

/** Resolved review author appearance. */
export interface TrackedChangeAuthor {
  name: string;
  kind: TrackedChangeAuthorKind;
  color: string;
  icon: ReactNode;
}

/** Custom card props; attach `measureRef` to the root element. */
export interface TrackedChangeCardRenderProps extends SidebarItemRenderProps {
  change: TrackedChangeEntry;
  author: TrackedChangeAuthor;
  isActive: boolean;
  index: number;
  total: number;
  /** True when the card cannot be matched to its change by position alone. */
  ambiguous: boolean;
  accept(): void;
  reject(): void;
  next(): void;
  previous(): void;
}

/** Review display and author customization. */
export interface TrackedChangesOptions {
  display?: 'cards' | 'dots';
  getAuthorAppearance?: (author: { name: string; kind: TrackedChangeAuthorKind }) => TrackedChangeAuthorAppearance | undefined;
  renderCard?: (props: TrackedChangeCardRenderProps) => ReactNode;
}

export interface ReviewSidebarItem extends ReactSidebarItem {
  review?: {
    change: TrackedChangeEntry;
    kind: TrackedChangeAuthorKind;
    proposalId?: string;
  };
}

export const TrackedChangesContext = createContext<{
  options?: TrackedChangesOptions;
  peers?: readonly CollaborationPeer[];
  user?: CollaborationUser;
  clientId?: number;
  items?: ReviewSidebarItem[];
  itemIndices?: ReadonlyMap<string, number>;
  navigate?: (id: string, direction: number) => void;
  measureCard?: (id: string, element: HTMLDivElement | null) => void;
  ambiguous?: ReadonlySet<string>;
}>({});

export interface ReviewLine {
  pageIndex: number;
  top: number;
  height: number;
}

/** A change's first line, which its card aligns to, and every rect it covers. */
export interface ReviewAnchor extends ReviewLine {
  id: string;
  lines: readonly ReviewLine[];
}

const sameLine = (first: ReviewLine, second: ReviewLine) => first.pageIndex === second.pageIndex
  && Math.abs(first.top + first.height / 2 - second.top - second.height / 2) < Math.min(first.height, second.height) / 2;

/** Ids whose change shares a rendered line with another change or whose card sits more than a line away from its first line. */
export function ambiguousReviewIds(anchors: readonly ReviewAnchor[], cardTops: ReadonlyMap<string, number> = new Map()): string[] {
  return anchors.filter(anchor => {
    const top = cardTops.get(anchor.id);
    return (top !== undefined && Math.abs(top - anchor.top) > anchor.height)
      || anchors.some(other => other !== anchor && anchor.lines.some(line => other.lines.some(otherLine => sameLine(line, otherLine))));
  }).map(anchor => anchor.id);
}

export function ReviewSpark() {
  return <svg className="docx-review-spark" viewBox="0 0 16 16" aria-hidden="true"><path className="docx-review-spark__path" fill="currentColor" d="M8 .8c.5 4.1 2.6 6.4 7.2 7.2-4.6.8-6.7 3.1-7.2 7.2C7.5 11.1 5.4 8.8.8 8 5.4 7.2 7.5 4.9 8 .8Z" /></svg>;
}

export function resolveReviewAuthor(name: string, kind: TrackedChangeAuthorKind = 'human', options?: TrackedChangesOptions,
  peers: readonly CollaborationPeer[] = [], user?: CollaborationUser, clientId = 0): TrackedChangeAuthor {
  const appearance = options?.getAuthorAppearance?.({ name, kind });
  const resolvedKind = appearance?.kind ?? kind;
  const peer = peers.find(peer => peer.user.name === name);
  let hash = 0;
  for (const char of name) hash = (Math.imul(hash, 31) + char.charCodeAt(0)) >>> 0;
  const color = appearance?.color ?? (peer ? resolvePresenceColor(peer.clientId, peer.user.color)
    : user?.name === name ? resolvePresenceColor(clientId, user.color) : resolvePresenceColor(hash, undefined));
  return { name, kind: resolvedKind, color, icon: appearance?.icon ?? (resolvedKind === 'agent' ? <ReviewSpark /> : getInitials(name)) };
}

export function useReviewPeers(presence?: CollaborationPresence) {
  const [remote, setRemote] = useState({ presence, peers: presence?.peers });
  useEffect(() => {
    setRemote({ presence, peers: presence?.peers });
    return presence?.onPeers(peers => setRemote({ presence, peers }));
  }, [presence]);
  return remote.presence === presence ? remote.peers : presence?.peers;
}

export function reviewAuthorStyle(author: TrackedChangeAuthor): CSSProperties {
  return { '--doc-review-author': author.color } as CSSProperties;
}

export function ReviewAvatar({ author }: { author: TrackedChangeAuthor }) {
  return <span className="docx-review-avatar" data-kind={author.kind} style={reviewAuthorStyle(author)} aria-hidden="true">{author.icon}</span>;
}

export function reviewRevisionIds(change: TrackedChangeEntry): number[] {
  return [...new Set([change.revisionId, ...(change.type === 'replacement' && change.insertionRevisionId != null ? [change.insertionRevisionId] : []),
    ...(change.coalescedRevisionIds ?? [])])];
}

export function decideReviewChange(change: TrackedChangeEntry, byId?: (id: number) => void, byRange?: (from: number, to: number) => void) {
  if (byId) for (const id of reviewRevisionIds(change)) byId(id);
  else byRange?.(change.from, change.to);
}
