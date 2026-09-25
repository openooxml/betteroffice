import { useCallback, useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react';
import type { OverflowMenuEntry } from '../../../../../shared/react-toolbar/OverflowMenu';
import { useRovingFocus } from '../../../../../shared/react-toolbar/useRovingFocus';
import { useOverflowSource } from '../toolbar/overflowRegistry';

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

function hint(title: string, shortcut: string | null | undefined, reason: string | undefined) {
  const named = shortcut ? `${title} (${shortcut})` : title;
  return reason ? `${named}: ${reason}` : named;
}

export interface ToolbarButtonProps {
  active?: boolean;
  disabled?: boolean;
  /** Why the button is disabled; announced and shown on hover, and it stays focusable. */
  description?: string;
  title: string;
  onClick?: () => void;
  children: ReactNode;
  style?: CSSProperties;
  testId?: string;
  ariaExpanded?: boolean;
}

export interface ToolbarButtonBaseProps extends Omit<ToolbarButtonProps, 'active'> {
  active?: boolean | 'mixed';
  /** Always exposes the pressed state, also when not pressed. */
  toggle?: boolean;
  shortcut?: string | null;
  className?: string;
  /** Overflow-menu entries; defaults to one item running `onClick`, `null` for none. */
  overflow?: (() => readonly OverflowMenuEntry[]) | null;
}

/** The button behind `ToolbarButton` and command buttons. */
export function ToolbarButtonBase({
  active,
  toggle = false,
  disabled = false,
  description,
  title,
  shortcut,
  onClick,
  children,
  style,
  testId,
  ariaExpanded,
  className,
  overflow,
}: ToolbarButtonBaseProps) {
  const [hovered, setHovered] = useState(false);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const id = useId();
  const reason = disabled && description ? description : undefined;
  const pressed = toggle ? (active ?? false) : active || undefined;
  useOverflowSource(
    buttonRef,
    overflow !== undefined
      ? overflow
      : () => [
          {
            kind: 'item',
            id,
            label: title,
            shortcut: shortcut ?? undefined,
            checked: toggle ? (active ?? false) : active,
            disabled,
            description: reason,
            onSelect: () => onClick?.(),
          },
        ]
  );
  return (
    <>
      <button
        ref={buttonRef}
        type="button"
        className={className}
        data-testid={testId}
        disabled={disabled && !reason}
        aria-disabled={reason ? true : undefined}
        aria-describedby={reason ? id : undefined}
        aria-label={title}
        aria-pressed={pressed}
        aria-expanded={ariaExpanded}
        title={hint(title, shortcut, reason)}
        onMouseDown={(event) => event.preventDefault()}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        onClick={disabled ? undefined : onClick}
        style={interactiveButtonStyle(disabled, active === true, hovered, style)}
      >
        {children}
      </button>
      {reason && (
        <span id={id} hidden>
          {reason}
        </span>
      )}
    </>
  );
}

export function ToolbarButton(props: ToolbarButtonProps) {
  return <ToolbarButtonBase {...props} />;
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
      aria-orientation="vertical"
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
  /** Why the dropdown is disabled; announced and shown on hover, and it stays focusable. */
  description?: string;
  active?: boolean;
  menuWidth?: number;
  style?: CSSProperties;
  testId?: string;
}

const MENU_ITEMS = '[role="menuitem"], [role="menuitemradio"], [role="menuitemcheckbox"]';
const MENU_ROLES: ReadonlySet<string> = new Set([
  'menuitem',
  'menuitemradio',
  'menuitemcheckbox',
  'separator',
]);
const FOCUSABLE =
  'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [href], [tabindex]:not([tabindex="-1"])';
const VIEWPORT_MARGIN = 8;
const MAX_POPUP_HEIGHT = 440;

type OpenMode = 'pointer' | 'first' | 'last';

/** Below the trigger when the popup fits or has more room there, else above; always on screen. */
function placeBelowOrAbove(
  trigger: DOMRect,
  height: number,
  width: number
): { top: number; left: number; maxHeight: number } {
  const viewportHeight = window.innerHeight;
  const below = Math.max(0, viewportHeight - VIEWPORT_MARGIN - (trigger.bottom + 4));
  const above = Math.max(0, trigger.top - 4 - VIEWPORT_MARGIN);
  const left = Math.max(
    VIEWPORT_MARGIN,
    Math.min(trigger.left, window.innerWidth - width - VIEWPORT_MARGIN)
  );
  if (Math.min(height, MAX_POPUP_HEIGHT) <= below || below >= above) {
    const maxHeight = Math.min(MAX_POPUP_HEIGHT, below);
    const top = Math.min(trigger.bottom + 4, viewportHeight - VIEWPORT_MARGIN - maxHeight);
    return { top: Math.max(VIEWPORT_MARGIN, top), left, maxHeight };
  }
  const maxHeight = Math.min(MAX_POPUP_HEIGHT, above);
  const top = trigger.top - 4 - Math.min(height, maxHeight);
  return { top: Math.max(VIEWPORT_MARGIN, top), left, maxHeight };
}

/**
 * A button with a popup. Content made only of menu items is a menu with arrow
 * keys, Home/End and typeahead; other content is a non-modal dialog. Escape
 * returns focus to the button.
 */
export function ToolbarDropdown({
  title,
  trigger,
  children,
  disabled = false,
  description,
  active = false,
  menuWidth = 220,
  style,
  testId,
}: ToolbarDropdownProps) {
  const [open, setOpen] = useState<OpenMode | null>(null);
  const [kind, setKind] = useState<'menu' | 'dialog'>('menu');
  const [hovered, setHovered] = useState(false);
  const [position, setPosition] = useState({ top: 0, left: 0, maxHeight: 440 });
  const triggerRef = useRef<HTMLButtonElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const popupId = useId();
  const descriptionId = useId();
  const reason = disabled && description ? description : undefined;
  const roving = useRovingFocus({
    container: menuRef,
    itemSelector: MENU_ITEMS,
    labelOf: (item) => item.getAttribute('aria-label') ?? item.textContent ?? '',
  });
  const rovingRef = useRef(roving);
  rovingRef.current = roving;

  const close = useCallback((restoreFocus = false) => {
    setOpen(null);
    if (restoreFocus) triggerRef.current?.focus({ preventScroll: true });
  }, []);

  const closeFromContent = useCallback(
    () => close(menuRef.current?.contains(document.activeElement) ?? false),
    [close]
  );

  useLayoutEffect(() => {
    const trigger = triggerRef.current;
    const menu = menuRef.current;
    if (!open || !trigger || !menu) return;
    setPosition(placeBelowOrAbove(trigger.getBoundingClientRect(), menu.scrollHeight, menuWidth));
    const isMenu = Array.from(menu.children).every(
      (child) => child.hasAttribute('hidden') || MENU_ROLES.has(child.getAttribute('role') ?? '')
    );
    setKind(isMenu ? 'menu' : 'dialog');
    if (open === 'pointer') return;
    if (isMenu) {
      if (open === 'first') rovingRef.current.focusFirst();
      else rovingRef.current.focusLast();
    } else {
      menu.querySelector<HTMLElement>(FOCUSABLE)?.focus({ preventScroll: true });
    }
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
    const onScroll = (event: Event) => {
      if (!(event.target instanceof Node && menuRef.current?.contains(event.target))) close();
    };
    document.addEventListener('mousedown', onPointerDown);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('scroll', onScroll, true);
    return () => {
      document.removeEventListener('mousedown', onPointerDown);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('scroll', onScroll, true);
    };
  }, [open, close]);

  const onMenuKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      close(true);
      return;
    }
    if (kind !== 'menu') return;
    if (event.key === 'Tab') {
      close();
      return;
    }
    if (roving.onKeyDown(event)) event.stopPropagation();
  };

  return (
    <div style={{ position: 'relative', display: 'inline-flex', flex: '0 0 auto' }}>
      <button
        ref={triggerRef}
        type="button"
        data-testid={testId}
        disabled={disabled && !reason}
        aria-disabled={reason ? true : undefined}
        aria-describedby={reason ? descriptionId : undefined}
        aria-label={title}
        aria-haspopup={kind}
        aria-expanded={open !== null}
        aria-controls={open ? popupId : undefined}
        title={hint(title, null, reason)}
        onMouseDown={(event) => event.preventDefault()}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        onClick={() => !disabled && setOpen((value) => (value ? null : 'pointer'))}
        onKeyDown={(event) => {
          if (disabled) return;
          if (event.key === 'Enter' || event.key === ' ' || event.key === 'ArrowDown') {
            event.preventDefault();
            setOpen('first');
          } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            setOpen('last');
          } else if (event.key === 'Escape' && open) {
            event.preventDefault();
            close();
          }
        }}
        style={interactiveButtonStyle(disabled, active || open !== null, hovered, style)}
      >
        {trigger}
      </button>
      {reason && (
        <span id={descriptionId} hidden>
          {reason}
        </span>
      )}
      {open && (
        <div
          ref={menuRef}
          id={popupId}
          role={kind}
          aria-label={title}
          aria-orientation={kind === 'menu' ? 'vertical' : undefined}
          tabIndex={-1}
          onKeyDown={onMenuKeyDown}
          onBlur={(event) => {
            if (kind !== 'dialog') return;
            const next = event.relatedTarget as Node | null;
            if (next && !menuRef.current?.contains(next) && !triggerRef.current?.contains(next)) {
              close();
            }
          }}
          onMouseDown={(event) => event.preventDefault()}
          style={{
            position: 'fixed',
            top: position.top,
            left: position.left,
            zIndex: 10000,
            width: menuWidth,
            maxWidth: `calc(100vw - ${2 * VIEWPORT_MARGIN}px)`,
            maxHeight: position.maxHeight,
            overflowY: 'auto',
            padding: 6,
            border: `1px solid ${toolbarColors.border}`,
            borderRadius: 8,
            background: toolbarColors.surface,
            boxShadow: '0 4px 16px rgba(60, 64, 67, 0.24)',
            boxSizing: 'border-box',
            outline: 'none',
          }}
        >
          {children(closeFromContent)}
        </div>
      )}
    </div>
  );
}

export interface ToolbarMenuItemProps {
  label: string;
  icon?: ReactNode;
  /** Makes the item a radio choice, checked when true. */
  selected?: boolean;
  disabled?: boolean;
  /** Why the item is disabled; announced and shown on hover. */
  description?: string;
  onClick?: () => void;
  close?: () => void;
}

export function ToolbarMenuItem({
  label,
  icon,
  selected,
  disabled = false,
  description,
  onClick,
  close,
}: ToolbarMenuItemProps) {
  const [hovered, setHovered] = useState(false);
  const id = useId();
  const reason = disabled && description ? description : undefined;
  return (
    <>
      <button
        type="button"
        role={selected === undefined ? 'menuitem' : 'menuitemradio'}
        tabIndex={-1}
        aria-checked={selected}
        aria-disabled={disabled || undefined}
        aria-describedby={reason ? id : undefined}
        aria-label={label}
        title={reason}
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
      {reason && (
        <span id={id} hidden>
          {reason}
        </span>
      )}
    </>
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
