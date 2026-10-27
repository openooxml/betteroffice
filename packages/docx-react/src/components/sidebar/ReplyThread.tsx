import { useContext } from 'react';
import type { Comment } from '@betteroffice/docx/types/content';
import { getCommentText, formatDate } from './cardUtils';
import { resolveReviewAuthor, ReviewAvatar, TrackedChangesContext } from './trackedChanges';
import { useTranslation } from '../../i18n';

export interface ReplyThreadProps {
  replies: Comment[];
  isExpanded: boolean;
}

export function ReplyThread({ replies, isExpanded }: ReplyThreadProps) {
  const { t } = useTranslation();
  const context = useContext(TrackedChangesContext);
  if (replies.length === 0) return null;
  const visibleReplies = isExpanded ? replies : replies.slice(-1);
  const hiddenCount = isExpanded ? 0 : replies.length - 1;

  return (
    <div className="docx-review-thread">
      {hiddenCount > 0 && (
        <div className="docx-review-thread__more">
          {t('comments.replyCount', { count: hiddenCount })}
        </div>
      )}
      {visibleReplies.map((reply) => {
        const author = resolveReviewAuthor(reply.author || t('comments.unknown'), 'human', context.options,
          context.peers, context.user, context.clientId);
        return (
          <div key={reply.id} className="docx-review-reply">
            <div className="docx-review-reply__header">
              <ReviewAvatar author={author} />
              <span className="docx-review-card__name">{author.name}</span>
              <span className="docx-review-card__meta">{formatDate(reply.date)}</span>
            </div>
            <div className="docx-review-reply__body" data-clamped={!isExpanded}>
              {getCommentText(reply.content)}
            </div>
          </div>
        );
      })}
    </div>
  );
}
