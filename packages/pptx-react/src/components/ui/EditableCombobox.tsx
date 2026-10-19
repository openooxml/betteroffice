import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { ToolbarIcon } from './ToolbarIcon';
import { tooltipText, useDisabledDescription } from './ToolbarPrimitives';
import {
  TOOLBAR_DISABLED_OPACITY,
  toolbarFocusStyle,
  toolbarMenuItemStyle,
  toolbarMenuStyle,
  toolbarTheme,
  useFocusRing,
} from '../../../../../shared/react-toolbar/theme';

export interface ComboboxOption {
  value: string;
  label: string;
}

export interface EditableComboboxProps {
  value: string;
  options: readonly ComboboxOption[];
  label: string;
  onCommit?: (value: string) => void;
  disabled?: boolean;
  /** Why the combobox is disabled; announced and shown on hover. */
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
  const described = useDisabledDescription(disabled, description);
  const [draft, setDraft] = useState(value);
  const [open, setOpen] = useState(false);
  const [hoveredOption, setHoveredOption] = useState<string | null>(null);
  const [focusedOption, setFocusedOption] = useState<string | null>(null);
  const focus = useFocusRing<HTMLInputElement>();
  const [position, setPosition] = useState({ top: 0, left: 0, width });
  const rootRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const editedRef = useRef(false);

  useEffect(() => {
    editedRef.current = false;
    setDraft(value);
  }, [value]);

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
    const onKeyDown = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape') {
        editedRef.current = false;
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
    if (!editedRef.current) return;
    editedRef.current = false;
    if (!disabled && onCommit && draft.trim()) onCommit(draft.trim());
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
        border: `1px solid ${toolbarTheme.line}`,
        borderRadius: 6,
        background: toolbarTheme.surface,
        opacity: disabled ? TOOLBAR_DISABLED_OPACITY : 1,
        boxSizing: 'border-box',
        flex: '0 0 auto',
        ...toolbarFocusStyle(focus.focusRing),
      }}
    >
      <input
        ref={inputRef}
        data-testid={testId}
        role="combobox"
        aria-label={label}
        aria-autocomplete="list"
        aria-expanded={open}
        {...described.props}
        readOnly={disabled}
        title={described.reason ? tooltipText(label, null, described.reason) : undefined}
        value={draft}
        onFocus={(event) => {
          focus.focusProps.onFocus(event);
          if (disabled) return;
          setOpen(true);
          event.currentTarget.select();
        }}
        onChange={(event) => {
          editedRef.current = true;
          setDraft(event.target.value);
        }}
        onKeyDown={(event) => {
          if (disabled) return;
          if (event.key === 'Enter') {
            commit();
            close();
            event.currentTarget.blur();
          } else if (event.key === 'Escape') {
            editedRef.current = false;
            setDraft(value);
          } else if (event.key === 'ArrowDown') {
            setOpen(true);
            event.preventDefault();
          }
        }}
        onBlur={(event) => {
          focus.focusProps.onBlur(event);
          if (disabled) {
            editedRef.current = false;
            setDraft(value);
            return;
          }
          requestAnimationFrame(() => {
            if (!menuRef.current?.contains(document.activeElement)) commit();
          });
        }}
        style={{
          appearance: 'none',
          width: 'calc(100% - 22px)',
          height: 26,
          padding: '0 2px 0 7px',
          border: 0,
          outline: 0,
          background: 'transparent',
          color: toolbarTheme.ink,
          font: '500 13px ui-sans-serif, system-ui, sans-serif',
          boxSizing: 'border-box',
          ...inputStyle,
        }}
      />
      {described.node}
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
          color: toolbarTheme.inkMuted,
          cursor: disabled ? 'default' : 'pointer',
        }}
      >
        <ToolbarIcon name="chevronDown" size={18} />
      </button>
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
            ...toolbarMenuStyle,
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
                editedRef.current = false;
                setDraft(option.label);
                onCommit?.(option.value);
                close();
              }}
              onMouseEnter={() => setHoveredOption(option.value)}
              onMouseLeave={() => setHoveredOption(null)}
              onFocus={() => setFocusedOption(option.value)}
              onBlur={() => setFocusedOption(null)}
              style={{
                ...toolbarMenuItemStyle({
                  hovered: hoveredOption === option.value,
                  focusRing: focusedOption === option.value,
                }),
                gap: 6,
                minHeight: 28,
                padding: '4px 8px 4px 4px',
              }}
            >
              <span style={{ display: 'inline-flex', width: 18, flex: '0 0 auto' }}>
                {option.value === value && <ToolbarIcon name="check" size={18} />}
              </span>
              {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
