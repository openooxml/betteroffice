import { useState } from 'react';
import { useTranslation } from '../../i18n';

export interface ReplyInputProps {
  onSubmit: (text: string) => void;
}

export function ReplyInput({ onSubmit }: ReplyInputProps) {
  const [active, setActive] = useState(false);
  const [text, setText] = useState('');
  const { t } = useTranslation();

  if (!active) {
    return (
      <div className="docx-review-composer" onClick={(e) => e.stopPropagation()}>
        <input
          key="idle"
          readOnly
          className="docx-review-input"
          onMouseDown={(e) => e.stopPropagation()}
          onClick={(e) => {
            e.stopPropagation();
            setActive(true);
          }}
          placeholder={t('comments.replyPlaceholder')}
        />
      </div>
    );
  }

  const trimmed = text.trim();

  return (
    <div className="docx-review-composer" data-active="true" onClick={(e) => e.stopPropagation()}>
      <input
        key="edit"
        ref={(el) => el?.focus({ preventScroll: true })}
        type="text"
        className="docx-review-input"
        value={text}
        onChange={(e) => setText(e.target.value)}
        onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          e.stopPropagation();
          if (e.key === 'Enter') {
            e.preventDefault();
            if (trimmed) onSubmit(trimmed);
            setText('');
            setActive(false);
          }
          if (e.key === 'Escape') {
            setActive(false);
            setText('');
          }
        }}
        placeholder={t('comments.replyPlaceholder')}
      />
      <div className="docx-review-composer__actions">
        <button
          type="button"
          className="docx-review-button"
          onClick={(e) => {
            e.stopPropagation();
            setActive(false);
            setText('');
          }}
        >
          {t('common.cancel')}
        </button>
        <button
          type="button"
          className="docx-review-button docx-review-button--primary"
          onClick={(e) => {
            e.stopPropagation();
            if (trimmed) onSubmit(trimmed);
            setText('');
            setActive(false);
          }}
          disabled={!trimmed}
        >
          {t('common.reply')}
        </button>
      </div>
    </div>
  );
}
