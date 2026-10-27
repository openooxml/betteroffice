import { useState } from 'react';
import type { SidebarItemRenderProps } from '../../plugin-api/types';
import { useTranslation } from '../../i18n';

export interface AddCommentCardProps extends SidebarItemRenderProps {
  onSubmit?: (text: string) => void | Promise<boolean>;
  onCancel?: () => void;
}

export function AddCommentCard({ measureRef, onSubmit, onCancel }: AddCommentCardProps) {
  const [text, setText] = useState('');
  const [pending, setPending] = useState(false);
  const { t } = useTranslation();

  const handleSubmit = async () => {
    if (!text.trim() || pending) return;
    setPending(true);
    const accepted = await onSubmit?.(text.trim());
    setPending(false);
    if (accepted !== false) setText('');
  };

  return (
    <div
      ref={measureRef}
      className="docx-comment-composer docx-review-card"
      data-active="true"
      onMouseDown={(e) => e.stopPropagation()}
      style={{ zIndex: 50 }}
    >
      <div className="docx-review-composer" data-active="true">
        <textarea
          ref={(el) => el?.focus({ preventScroll: true })}
          className="docx-review-input"
          value={text}
          readOnly={pending}
          onChange={(e) => setText(e.target.value)}
          onMouseDown={(e) => e.stopPropagation()}
          onKeyDown={(e) => {
            e.stopPropagation();
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              void handleSubmit();
            }
            if (e.key === 'Escape') {
              onCancel?.();
              setText('');
            }
          }}
          placeholder={t('comments.addComment')}
        />
        <div className="docx-review-composer__actions">
          <button
            type="button"
            className="docx-review-button"
            onClick={() => {
              onCancel?.();
              setText('');
            }}
          >
            {t('common.cancel')}
          </button>
          <button
            type="button"
            className="docx-review-button docx-review-button--primary"
            onClick={() => { void handleSubmit(); }}
            disabled={!text.trim() || pending}
          >
            {t('common.comment')}
          </button>
        </div>
      </div>
    </div>
  );
}
