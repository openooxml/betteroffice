import { useContext, useEffect, useState } from 'react';
import type { Comment } from '@betteroffice/docx/types/content';
import type { SidebarItemRenderProps } from '../../plugin-api/types';
import { getCommentText, formatDate } from './cardUtils';
import { ReplyThread } from './ReplyThread';
import { ReplyInput } from './ReplyInput';
import { ReviewCheckIcon, ReviewMoreIcon, ReviewUndoIcon } from './reviewIcons';
import { resolveReviewAuthor, ReviewAvatar, reviewAuthorStyle, TrackedChangesContext } from './trackedChanges';
import { useTranslation } from '../../i18n';

export interface CommentCardProps extends SidebarItemRenderProps {
  comment: Comment;
  replies: Comment[];
  onReply?: (commentId: number, text: string) => void;
  onResolve?: (commentId: number) => void;
  onUnresolve?: (commentId: number) => void;
  onDelete?: (commentId: number) => void;
}

export function CommentCard({
  comment,
  replies,
  isExpanded,
  onToggleExpand,
  measureRef,
  onReply,
  onResolve,
  onUnresolve,
  onDelete,
}: CommentCardProps) {
  const [menuOpen, setMenuOpen] = useState(false);
  const { t } = useTranslation();
  const context = useContext(TrackedChangesContext);
  const author = resolveReviewAuthor(comment.author || t('comments.unknown'), 'human', context.options,
    context.peers, context.user, context.clientId);
  const date = formatDate(comment.date);
  const resolveLabel = comment.done ? t('comments.reopen') : t('comments.resolve');

  useEffect(() => {
    if (!menuOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setMenuOpen(false);
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [menuOpen]);

  return (
    <div
      ref={measureRef}
      data-comment-id={comment.id}
      data-active={isExpanded}
      data-resolved={comment.done ? 'true' : undefined}
      data-kind={author.kind}
      className="docx-comment-card docx-review-card"
      onClick={onToggleExpand}
      onMouseDown={(e) => e.stopPropagation()}
      style={reviewAuthorStyle(author)}
    >
      <div className="docx-review-card__header">
        <ReviewAvatar author={author} />
        <span className="docx-review-card__name">{author.name}</span>
        <span className="docx-review-card__meta">
          {comment.done && <span className="docx-review-card__status">{t('comments.resolved')}</span>}
          {comment.done && date ? ' · ' : ''}
          {date}
        </span>
        {isExpanded && (
          <span className="docx-review-card__actions">
            <button
              type="button"
              className="docx-review-resolve"
              onClick={(e) => {
                e.stopPropagation();
                if (comment.done) {
                  onUnresolve?.(comment.id);
                } else {
                  onResolve?.(comment.id);
                }
              }}
              aria-label={resolveLabel}
              title={resolveLabel}
            >
              {comment.done ? <ReviewUndoIcon /> : <ReviewCheckIcon />}
            </button>
            <span className="docx-review-menu-anchor">
              <button
                type="button"
                className="docx-review-more"
                onClick={(e) => {
                  e.stopPropagation();
                  setMenuOpen(!menuOpen);
                }}
                aria-expanded={menuOpen}
                aria-haspopup="menu"
                aria-label={t('comments.moreOptions')}
                title={t('comments.moreOptions')}
              >
                <ReviewMoreIcon />
              </button>
              {menuOpen && (
                <div
                  role="menu"
                  className="docx-review-menu"
                  onClick={(e) => e.stopPropagation()}
                  onMouseDown={(e) => e.stopPropagation()}
                >
                  <button
                    type="button"
                    role="menuitem"
                    className="docx-review-menu__item docx-review-menu__item--danger"
                    onClick={() => {
                      setMenuOpen(false);
                      onDelete?.(comment.id);
                    }}
                  >
                    {t('common.delete')}
                  </button>
                </div>
              )}
            </span>
          </span>
        )}
      </div>

      <div className="docx-review-card__body">{getCommentText(comment.content)}</div>

      <div className="docx-review-card__replies">
        <ReplyThread replies={replies} isExpanded={isExpanded} />
        {isExpanded && !comment.done && (
          <ReplyInput onSubmit={(text) => onReply?.(comment.id, text)} />
        )}
      </div>
    </div>
  );
}
