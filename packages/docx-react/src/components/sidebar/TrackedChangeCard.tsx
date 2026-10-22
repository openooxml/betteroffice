import { useCallback, useContext } from 'react';
import type { Comment } from '@betteroffice/docx/types/content';
import type { SidebarItemRenderProps } from '../../plugin-api/types';
import type { TrackedChangeEntry } from './cardUtils';
import { ReplyThread } from './ReplyThread';
import { ReplyInput } from './ReplyInput';
import { useTranslation } from '../../i18n';
import { ReviewCheckIcon, ReviewXIcon } from './reviewIcons';
import { decideReviewChange, resolveReviewAuthor, ReviewAvatar, reviewAuthorStyle, TrackedChangesContext,
  type TrackedChangeAuthor, type TrackedChangeAuthorKind, type TrackedChangesOptions } from './trackedChanges';

export interface TrackedChangeCardProps extends SidebarItemRenderProps {
  change: TrackedChangeEntry;
  replies: Comment[];
  /**
   * @deprecated Prefer `onAcceptById`. Range-based accept only clears
   * marks within `(from, to)` and silently leaves paragraph-mark and
   * coalesced sibling sites behind. Kept as fallback for hosts that
   * haven't migrated to the by-id channel.
   */
  onAccept?: (from: number, to: number) => void;
  /**
   * @deprecated Prefer `onRejectById`. Same caveat as `onAccept`.
   */
  onReject?: (from: number, to: number) => void;
  /**
   * Accept every site of the revision. Walks the doc for all sites
   * sharing the `revisionId` (inline marks + paragraph attrs + table
   * row/cell attrs) and clears them in one transaction. This is the
   * right channel for any coalesced revision.
   */
  onAcceptById?: (revisionId: number) => void;
  /** Reject every site of the revision. Counterpart to `onAcceptById`. */
  onRejectById?: (revisionId: number) => void;
  author?: TrackedChangeAuthor;
  defaultKind?: TrackedChangeAuthorKind;
  reviewItemId?: string;
  isActive?: boolean;
  index?: number;
  total?: number;
  onNext?: () => void;
  onPrevious?: () => void;
  renderCard?: TrackedChangesOptions['renderCard'];
  onReply?: (revisionId: number, text: string) => void;
}

export function TrackedChangeCard({
  change, replies, isExpanded, onToggleExpand, measureRef, onAccept, onReject, onAcceptById, onRejectById, onReply,
  author: suppliedAuthor, defaultKind = 'human', reviewItemId, isActive = isExpanded, index: suppliedIndex, total: suppliedTotal,
  onNext, onPrevious, renderCard,
}: TrackedChangeCardProps) {
  const { t } = useTranslation();
  const context = useContext(TrackedChangesContext);
  const author = suppliedAuthor ?? resolveReviewAuthor(change.author || t('trackedChanges.unknown'), defaultKind,
    context.options, context.peers, context.user, context.clientId);
  const index = suppliedIndex ?? Math.max(0, context.items?.findIndex(item => item.id === reviewItemId) ?? 0);
  const total = suppliedTotal ?? context.items?.length ?? 1;
  const accept = () => decideReviewChange(change, onAcceptById, onAccept);
  const reject = () => decideReviewChange(change, onRejectById, onReject);
  const next = onNext ?? (() => { if (reviewItemId) context.navigate?.(reviewItemId, 1); });
  const previous = onPrevious ?? (() => { if (reviewItemId) context.navigate?.(reviewItemId, -1); });
  const register = context.measureCard;
  const measured = useCallback((element: HTMLDivElement | null) => {
    measureRef(element);
    if (reviewItemId) register?.(reviewItemId, element);
  }, [measureRef, register, reviewItemId]);
  const custom = renderCard ?? context.options?.renderCard;
  if (custom) return custom({ change, author, isActive, index, total, accept, reject, next, previous,
    isExpanded, onToggleExpand, measureRef: measured });

  const date = change.date && new Date(change.date);
  const time = date && !Number.isNaN(date.getTime()) ? date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '';
  return (
    <div ref={measured} className="docx-tracked-change-card docx-review-card" data-active={isActive} data-kind={author.kind}
      data-revision-id={change.revisionId} onClick={onToggleExpand} onMouseDown={event => event.stopPropagation()} style={reviewAuthorStyle(author)}>
      <div className="docx-review-card__header">
        <ReviewAvatar author={author} />
        <span className="docx-review-card__name">{author.name}</span>
        <span className="docx-review-card__meta">{author.kind === 'agent' ? t('trackedChanges.agent') : ''}{author.kind === 'agent' && time ? ' · ' : ''}{time}</span>
        <span className="docx-review-card__actions">
          <button type="button" className="docx-review-accept" aria-label={t('commands.reviewAccept')} title={t('commands.reviewAccept')}
            onClick={event => { event.stopPropagation(); accept(); }}><ReviewCheckIcon /></button>
          <button type="button" className="docx-review-reject" aria-label={t('commands.reviewReject')} title={t('commands.reviewReject')}
            onClick={event => { event.stopPropagation(); reject(); }}><ReviewXIcon /></button>
        </span>
      </div>
      {isActive && <div className="docx-review-nav">
        <button type="button" className="docx-review-nav-prev" aria-label={t('trackedChanges.previousChange')} onClick={event => { event.stopPropagation(); previous(); }}>‹</button>
        <span className="docx-review-nav__count">{index + 1}/{total}</span>
        <button type="button" className="docx-review-nav-next" aria-label={t('trackedChanges.nextChange')} onClick={event => { event.stopPropagation(); next(); }}>›</button>
      </div>}
      <div className="docx-review-card__replies"><ReplyThread replies={replies} isExpanded={isExpanded} />
        {isExpanded && <ReplyInput onSubmit={text => onReply?.(change.revisionId, text)} />}</div>
    </div>
  );
}
