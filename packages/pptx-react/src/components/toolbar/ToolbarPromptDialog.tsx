import { useId, useLayoutEffect, useRef, useState } from 'react';
import type { KeyboardEvent as ReactKeyboardEvent } from 'react';
import { useTranslation } from '../../i18n';
import { interactiveButtonStyle, toolbarColors } from '../ui/ToolbarPrimitives';
import type { OverflowPrompt } from './overflowRegistry';

const FOCUSABLE = 'input, button:not([disabled])';

/** A modal dialog asking for one value an overflow action needs. */
export function ToolbarPromptDialog({
  prompt,
  onClose,
}: {
  prompt: OverflowPrompt;
  onClose(): void;
}) {
  const { t } = useTranslation();
  const [value, setValue] = useState(prompt.initialValue ?? '');
  const dialogRef = useRef<HTMLFormElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const titleId = useId();
  const inputId = useId();
  const trimmed = value.trim();
  const valid = prompt.valid(trimmed);

  useLayoutEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.select();
  }, []);

  const onKeyDown = (event: ReactKeyboardEvent) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onClose();
      return;
    }
    if (event.key !== 'Tab') return;
    const focusable = Array.from(dialogRef.current?.querySelectorAll<HTMLElement>(FOCUSABLE) ?? []);
    if (focusable.length === 0) return;
    const index = focusable.indexOf(document.activeElement as HTMLElement);
    const next = event.shiftKey
      ? index <= 0
        ? focusable.length - 1
        : index - 1
      : index === focusable.length - 1
      ? 0
      : index + 1;
    event.preventDefault();
    focusable[next].focus();
  };

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 10001,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: 16,
        background: 'rgba(32, 33, 36, 0.32)',
      }}
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <form
        ref={dialogRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        onKeyDown={onKeyDown}
        onSubmit={(event) => {
          event.preventDefault();
          if (!valid) return;
          onClose();
          prompt.submit(trimmed);
        }}
        style={{
          width: 'min(320px, 100%)',
          boxSizing: 'border-box',
          padding: 16,
          borderRadius: 8,
          background: toolbarColors.surface,
          color: toolbarColors.text,
          boxShadow: '0 8px 24px rgba(60, 64, 67, 0.3)',
          font: '400 13px ui-sans-serif, system-ui, sans-serif',
        }}
      >
        <h2 id={titleId} style={{ margin: '0 0 12px', fontSize: 15, fontWeight: 600 }}>
          {prompt.title}
        </h2>
        <label htmlFor={inputId} style={{ display: 'block', marginBottom: 4, fontSize: 12 }}>
          {prompt.label}
        </label>
        <input
          id={inputId}
          ref={inputRef}
          value={value}
          placeholder={prompt.placeholder}
          aria-invalid={trimmed !== '' && !valid}
          onChange={(event) => setValue(event.target.value)}
          style={{
            width: '100%',
            boxSizing: 'border-box',
            padding: '6px 8px',
            border: `1px solid ${toolbarColors.border}`,
            borderRadius: 4,
            background: toolbarColors.surface,
            color: toolbarColors.text,
          }}
        />
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
          <button
            type="button"
            onClick={onClose}
            style={interactiveButtonStyle(false, false, false)}
          >
            {t('commands.cancel')}
          </button>
          <button
            type="submit"
            disabled={!valid}
            style={interactiveButtonStyle(!valid, true, false)}
          >
            {t('commands.apply')}
          </button>
        </div>
      </form>
    </div>
  );
}
