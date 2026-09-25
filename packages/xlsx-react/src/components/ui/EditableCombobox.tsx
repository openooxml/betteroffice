import { useCallback, useEffect, useId, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { ToolbarIcon } from './ToolbarIcon';
import { toolbarColors } from './ToolbarPrimitives';

export interface ComboboxOption {
  value: string;
  label: string;
}

export interface EditableComboboxProps {
  value: string;
  options: readonly ComboboxOption[];
  label: string;
  /** `pointer` is true when an option was clicked. */
  onCommit?: (value: string, pointer: boolean) => void;
  disabled?: boolean;
  /** Why the field is disabled; announced and shown on hover, and it stays focusable. */
  description?: string;
  width?: number;
  inputStyle?: CSSProperties;
  testId?: string;
}

export function EditableCombobox({
  value,
  options,
  label,
  onCommit,
  disabled = false,
  description,
  width = 72,
  inputStyle,
  testId,
}: EditableComboboxProps) {
  const descriptionId = useId();
  const reason = disabled && description ? description : undefined;
  const [draft, setDraft] = useState(value);
  const [open, setOpen] = useState(false);
  const [position, setPosition] = useState({ top: 0, left: 0, width });
  const rootRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => setDraft(value), [value]);

  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open || !rootRef.current) return;
    const rect = rootRef.current.getBoundingClientRect();
    setPosition({
      top: rect.bottom + 4,
      left: rect.left,
      width: Math.max(rect.width, width),
    });
  }, [open, width]);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (!rootRef.current?.contains(target) && !menuRef.current?.contains(target)) close();
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setDraft(value);
        close();
      }
    };
    document.addEventListener('mousedown', onPointerDown);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('mousedown', onPointerDown);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [open, close, value]);

  const commit = useCallback(() => {
    const next = draft.trim();
    if (!disabled && onCommit && next && next !== value) onCommit(next, false);
    setDraft(value);
  }, [disabled, onCommit, draft, value]);

  return (
    <div
      ref={rootRef}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        width,
        height: 28,
        border: `1px solid ${toolbarColors.border}`,
        borderRadius: 4,
        background: disabled ? 'rgba(255,255,255,0.5)' : toolbarColors.surface,
        opacity: disabled ? 0.48 : 1,
        boxSizing: 'border-box',
        flex: '0 0 auto',
      }}
    >
      <input
        ref={inputRef}
        data-testid={testId}
        role="combobox"
        aria-label={label}
        aria-autocomplete="list"
        aria-expanded={open}
        disabled={disabled && !reason}
        readOnly={Boolean(reason)}
        aria-disabled={reason ? true : undefined}
        aria-describedby={reason ? descriptionId : undefined}
        title={reason ? `${label}: ${reason}` : undefined}
        value={draft}
        onFocus={(event) => {
          if (disabled) return;
          setOpen(true);
          event.currentTarget.select();
        }}
        onChange={(event) => setDraft(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            commit();
            close();
          } else if (event.key === 'ArrowDown') {
            setOpen(true);
            event.preventDefault();
          }
        }}
        onBlur={() => {
          requestAnimationFrame(() => {
            if (!menuRef.current?.contains(document.activeElement)) commit();
          });
        }}
        style={{
          appearance: 'none',
          width: `calc(100% - 22px)`,
          height: 26,
          padding: '0 2px 0 7px',
          border: 0,
          outline: 0,
          background: 'transparent',
          color: toolbarColors.text,
          font: '500 13px ui-sans-serif, system-ui, sans-serif',
          boxSizing: 'border-box',
          ...inputStyle,
        }}
      />
      <button
        type="button"
        disabled={disabled}
        tabIndex={-1}
        aria-label={label}
        title={label}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => {
          if (disabled) return;
          setOpen((current) => !current);
          inputRef.current?.focus();
        }}
        style={{
          appearance: 'none',
          display: 'grid',
          placeItems: 'center',
          width: 21,
          height: 26,
          padding: 0,
          border: 0,
          background: 'transparent',
          color: toolbarColors.muted,
          cursor: disabled ? 'default' : 'pointer',
        }}
      >
        <ToolbarIcon name="chevronDown" size={14} />
      </button>
      {reason && (
        <span id={descriptionId} hidden>
          {reason}
        </span>
      )}
      {open && !disabled && (
        <div
          ref={menuRef}
          role="listbox"
          aria-label={label}
          style={{
            position: 'fixed',
            top: position.top,
            left: position.left,
            zIndex: 10000,
            width: position.width,
            maxHeight: 260,
            overflowY: 'auto',
            padding: 4,
            border: `1px solid ${toolbarColors.border}`,
            borderRadius: 6,
            background: toolbarColors.surface,
            boxShadow: '0 4px 16px rgba(60, 64, 67, 0.24)',
            boxSizing: 'border-box',
          }}
        >
          {options.map((option) => (
            <button
              key={option.value}
              type="button"
              role="option"
              aria-selected={option.value === value}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                setDraft(option.label);
                onCommit?.(option.value, true);
                close();
              }}
              style={{
                appearance: 'none',
                display: 'block',
                width: '100%',
                minHeight: 28,
                padding: '4px 7px',
                border: 0,
                borderRadius: 3,
                background: option.value === value ? toolbarColors.active : 'transparent',
                color: toolbarColors.text,
                cursor: 'pointer',
                font: '400 13px ui-sans-serif, system-ui, sans-serif',
                textAlign: 'left',
              }}
            >
              {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
