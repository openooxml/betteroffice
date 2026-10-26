import type { Comment } from '@betteroffice/docx/types/content';
import type { SidebarItemRenderProps } from '../../plugin-api/types';
import { ReviewResolvedIcon } from './reviewIcons';
import { useTranslation } from '../../i18n';

export interface ResolvedCommentMarkerProps extends SidebarItemRenderProps {
  comment: Comment;
}

export function ResolvedCommentMarker({
  comment,
  measureRef,
  onToggleExpand,
}: ResolvedCommentMarkerProps) {
  const { t } = useTranslation();
  return (
    <div
      ref={measureRef}
      data-comment-id={comment.id}
      data-comment-resolved="true"
      className="docx-review-resolved-marker"
      title={t('comments.resolved')}
      onClick={onToggleExpand}
      onMouseDown={(e) => e.stopPropagation()}
    >
      <ReviewResolvedIcon />
    </div>
  );
}
