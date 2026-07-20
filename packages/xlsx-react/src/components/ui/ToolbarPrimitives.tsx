import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';

export const toolbarColors = {
  text: '#3c4043',
  muted: '#5f6368',
  disabled: '#9aa0a6',
  hover: '#e2e7ef',
  active: '#d3e3fd',
  border: '#c7cacf',
  surface: '#ffffff',
  rail: '#edf2fa',
};

const baseButtonStyle: CSSProperties = {
  appearance: 'none',
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  gap: 2,
  minWidth: 28,
  height: 28,
  padding: '0 5px',
  border: 0,
  borderRadius: 4,
  color: toolbarColors.text,
  font: '500 13px ui-sans-serif, system-ui, sans-serif',
  lineHeight: 1,
  whiteSpace: 'nowrap',
  boxSizing: 'border-box',
};

function interactiveButtonStyle(
  disabled: boolean,
  active: boolean,
  hovered: boolean,
  style?: CSSProperties
): CSSProperties {
  return {
    ...baseButtonStyle,
    background: active
      ? toolbarColors.active
      : hovered && !disabled
      ? toolbarColors.hover
      : 'transparent',
    color: disabled ? toolbarColors.disabled : toolbarColors.text,
    cursor: disabled ? 'default' : 'pointer',
    opacity: disabled ? 0.48 : 1,
    ...style,
  };
}

export interface ToolbarButtonProps {
  active?: boolean;
  disabled?: boolean;
  title: string;
  onClick?: () => void;
  children: ReactNode;
  style?: CSSProperties;
  testId?: string;
  ariaExpanded?: boolean;
}

export function ToolbarButton({
  active = false,
  disabled = false,
  title,
  onClick,
  children,
  style,
  testId,
  ariaExpanded,
}: ToolbarButtonProps) {
  const [hovered, setHovered] = useState(false);
  return (
    <button
      type="button"
      data-testid={testId}
      disabled={disabled}
      aria-label={title}
      aria-pressed={active || undefined}
      aria-expanded={ariaExpanded}
      title={title}
      onMouseDown={(event) => event.preventDefault()}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      onClick={disabled ? undefined : onClick}
      style={interactiveButtonStyle(disabled, active, hovered, style)}
    >
      {children}
    </button>
  );
}

export function ToolbarGroup({
  label,
  children,
  style,
}: {
  label: string;
  children: ReactNode;
  style?: CSSProperties;
}) {
  return (
    <div
      role="group"
      aria-label={label}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        gap: 1,
        flex: '0 0 auto',
        ...style,
      }}
    >
      {children}
    </div>
  );
}

export function ToolbarSeparator() {
  return (
    <div
      role="separator"
      style={{
        width: 1,
        height: 24,
        margin: '0 5px',
        background: toolbarColors.border,
        flex: '0 0 auto',
      }}
    />
  );
}

export interface ToolbarDropdownProps {
  title: string;
  trigger: ReactNode;
  children: (close: () => void) => ReactNode;
  disabled?: boolean;
  active?: boolean;
  menuWidth?: number;
  style?: CSSProperties;
  testId?: string;
}

export function ToolbarDropdown({
  title,
  trigger,
  children,
  disabled = false,
  active = false,
  menuWidth = 220,
  style,
  testId,
}: ToolbarDropdownProps) {
  const [open, setOpen] = useState(false);
  const [hovered, setHovered] = useState(false);
  const [position, setPosition] = useState({ top: 0, left: 0 });
  const triggerRef = useRef<HTMLButtonElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);

  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const left = Math.min(rect.left, Math.max(8, window.innerWidth - menuWidth - 8));
    setPosition({ top: rect.bottom + 4, left });
  }, [open, menuWidth]);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (!triggerRef.current?.contains(target) && !menuRef.current?.contains(target)) close();
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') close();
    };
    const onScroll = () => close();
    document.addEventListener('mousedown', onPointerDown);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('scroll', onScroll, true);
    return () => {
      document.removeEventListener('mousedown', onPointerDown);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('scroll', onScroll, true);
    };
  }, [open, close]);

  return (
    <div style={{ position: 'relative', display: 'inline-flex', flex: '0 0 auto' }}>
      <button
        ref={triggerRef}
        type="button"
        data-testid={testId}
        disabled={disabled}
        aria-label={title}
        aria-haspopup="menu"
        aria-expanded={open}
        title={title}
        onMouseDown={(event) => event.preventDefault()}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        onClick={() => !disabled && setOpen((value) => !value)}
        style={interactiveButtonStyle(disabled, active || open, hovered, style)}
      >
        {trigger}
      </button>
      {open && (
        <div
          ref={menuRef}
          role="menu"
          aria-label={title}
          onMouseDown={(event) => event.preventDefault()}
          style={{
            position: 'fixed',
            top: position.top,
            left: position.left,
            zIndex: 10000,
            width: menuWidth,
            maxHeight: 'min(440px, calc(100vh - 16px))',
            overflowY: 'auto',
            padding: 6,
            border: `1px solid ${toolbarColors.border}`,
            borderRadius: 8,
            background: toolbarColors.surface,
            boxShadow: '0 4px 16px rgba(60, 64, 67, 0.24)',
            boxSizing: 'border-box',
          }}
        >
          {children(close)}
        </div>
      )}
    </div>
  );
}

export interface ToolbarMenuItemProps {
  label: string;
  icon?: ReactNode;
  selected?: boolean;
  disabled?: boolean;
  onClick?: () => void;
  close?: () => void;
}

export function ToolbarMenuItem({
  label,
  icon,
  selected = false,
  disabled = false,
  onClick,
  close,
}: ToolbarMenuItemProps) {
  const [hovered, setHovered] = useState(false);
  return (
    <button
      type="button"
      role="menuitem"
      disabled={disabled}
      aria-label={label}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      onClick={() => {
        if (disabled) return;
        onClick?.();
        close?.();
      }}
      style={{
        appearance: 'none',
        display: 'flex',
        alignItems: 'center',
        gap: 10,
        width: '100%',
        minHeight: 32,
        padding: '5px 9px',
        border: 0,
        borderRadius: 4,
        background: hovered && !disabled ? toolbarColors.hover : 'transparent',
        color: disabled ? toolbarColors.disabled : toolbarColors.text,
        cursor: disabled ? 'default' : 'pointer',
        opacity: disabled ? 0.48 : 1,
        font: '400 13px ui-sans-serif, system-ui, sans-serif',
        textAlign: 'left',
        boxSizing: 'border-box',
      }}
    >
      <span
        style={{
          display: 'inline-grid',
          placeItems: 'center',
          width: 20,
          flex: '0 0 auto',
        }}
      >
        {icon}
      </span>
      <span style={{ flex: 1 }}>{label}</span>
      {selected && <span aria-hidden="true">✓</span>}
    </button>
  );
}

export function ToolbarMenuSeparator() {
  return (
    <div
      role="separator"
      style={{ height: 1, margin: '5px 2px', background: toolbarColors.border }}
    />
  );
}
