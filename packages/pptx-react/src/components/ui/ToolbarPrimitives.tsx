import {
  forwardRef,
  useCallback,
  useEffect,
  useId,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { ToolbarIcon } from './ToolbarIcon';
import {
  inheritedToolbarTokens,
  toolbarButtonStyle,
  toolbarMenuItemStyle,
  toolbarMenuSeparatorStyle,
  toolbarMenuStyle,
  toolbarSeparatorStyle,
  useFocusRing,
} from '../../../../../shared/react-toolbar/theme';
import { useRovingFocus } from '../../../../../shared/react-toolbar/useRovingFocus';
import { useCommandChromeRef } from '../../commands/hooks';
import { useOverflowRegistry, useOverflowSource } from '../toolbar/overflowRegistry';

/** A tooltip naming the control, its shortcut and, when disabled, why. */
export function tooltipText(title: string, shortcut?: string | null, reason?: string): string {
  const named = shortcut ? `${title} (${shortcut})` : title;
  return reason ? `${named}: ${reason}` : named;
}

/** Keeps a disabled control with a stated reason focusable, so the reason can be read. */
export function useDisabledDescription(disabled: boolean, description: string | undefined) {
  const id = useId();
  const reason = disabled && description ? description : undefined;
  return {
    reason,
    props: reason
      ? { disabled: false, 'aria-disabled': true as const, 'aria-describedby': id }
      : { disabled },
    node: reason ? (
      <span id={id} hidden>
        {reason}
      </span>
    ) : null,
  };
}

export interface ToolbarButtonProps {
  /** Pressed state; omit for buttons that are not toggles. */
  active?: boolean | 'mixed';
  disabled?: boolean;
  /** Why the button is disabled; announced and shown on hover. */
  description?: string;
  /** Accessible name and tooltip. */
  title: string;
  /** Keyboard shortcut shown in the tooltip and overflow menu, such as `Ctrl+B`. */
  shortcut?: string;
  onClick?: () => void;
  children: ReactNode;
  className?: string;
  style?: CSSProperties;
  testId?: string;
  ariaExpanded?: boolean;
  /** Label of the overflow-menu entry; defaults to `title`. */
  overflowLabel?: string;
}

/** A toolbar button without an overflow entry, for composite controls. */
export const ToolbarButtonBase = forwardRef<HTMLButtonElement, ToolbarButtonProps>(
  function ToolbarButtonBase(
    {
      active,
      disabled = false,
      description,
      title,
      shortcut,
      onClick,
      children,
      className,
      style,
      testId,
      ariaExpanded,
    },
    ref
  ) {
    const [hovered, setHovered] = useState(false);
    const focus = useFocusRing<HTMLButtonElement>();
    const described = useDisabledDescription(disabled, description);
    return (
      <>
        <button
          ref={ref}
          type="button"
          className={className}
          data-testid={testId}
          {...described.props}
          aria-label={title}
          aria-pressed={active === undefined ? undefined : active}
          aria-expanded={ariaExpanded}
          title={tooltipText(title, shortcut, described.reason)}
          onMouseDown={(event) => event.preventDefault()}
          onMouseEnter={() => setHovered(true)}
          onMouseLeave={() => setHovered(false)}
          onClick={disabled ? undefined : onClick}
          {...focus.focusProps}
          style={{
            ...toolbarButtonStyle({
              disabled,
              pressed: active === true,
              hovered,
              focusRing: focus.focusRing,
            }),
            ...style,
          }}
        >
          {children}
        </button>
        {described.node}
      </>
    );
  }
);

/** A toolbar button; at narrow widths it moves into the overflow menu. */
export function ToolbarButton(props: ToolbarButtonProps) {
  const buttonRef = useRef<HTMLButtonElement>(null);
  const id = useId();
  const { active, disabled = false, description, title, shortcut, onClick, overflowLabel } = props;
  useOverflowSource(buttonRef, () => [
    {
      kind: 'item',
      id,
      label: overflowLabel ?? title,
      shortcut,
      checked: active,
      disabled,
      description: disabled ? description : undefined,
      onSelect: () => onClick?.(),
    },
  ]);
  return <ToolbarButtonBase ref={buttonRef} {...props} />;
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
      style={toolbarSeparatorStyle}
    />
  );
}

const MENU_ITEM_SELECTOR = '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]';
const FOCUSABLE_SELECTOR =
  'button, input, select, textarea, a[href], [tabindex]:not([tabindex="-1"])';
const VIEWPORT_MARGIN = 8;

export interface ToolbarDropdownProps {
  title: string;
  trigger: ReactNode;
  children: (close: () => void) => ReactNode;
  disabled?: boolean;
  /** Why the dropdown is disabled; announced and shown on hover. */
  description?: string;
  active?: boolean;
  menuWidth?: number;
  style?: CSSProperties;
  testId?: string;
  /**
   * `menu` for content made of `ToolbarMenuItem`s, `dialog` for other controls.
   * Omitted, it follows the content.
   */
  popup?: 'menu' | 'dialog';
}

export interface DropdownHandle {
  /** Opens the popup at `anchor`, returning focus there when it closes. */
  openAt(anchor: HTMLElement | null): void;
}

type OpenState = { focus: 'first' | 'last' | 'popup'; anchor: HTMLElement | null } | null;

function placePopup(anchor: DOMRect, popup: HTMLElement, width: number) {
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const left = Math.max(
    VIEWPORT_MARGIN,
    Math.min(
      anchor.left,
      viewportWidth - Math.min(width, viewportWidth - 2 * VIEWPORT_MARGIN) - VIEWPORT_MARGIN
    )
  );
  const height = popup.scrollHeight;
  const below = viewportHeight - VIEWPORT_MARGIN - (anchor.bottom + 4);
  const above = anchor.top - 4 - VIEWPORT_MARGIN;
  if (height <= below || below >= above) {
    return { top: anchor.bottom + 4, left, maxHeight: Math.max(below, 80) };
  }
  const maxHeight = Math.max(above, 80);
  return {
    top: Math.max(VIEWPORT_MARGIN, anchor.top - 4 - Math.min(height, maxHeight)),
    left,
    maxHeight,
  };
}

/** A button with a popup, without an overflow entry, for composite controls. */
export const ToolbarDropdownBase = forwardRef<DropdownHandle, ToolbarDropdownProps>(
  function ToolbarDropdownBase(
    {
      title,
      trigger,
      children,
      disabled = false,
      description,
      active = false,
      menuWidth = 220,
      style,
      testId,
      popup,
    },
    ref
  ) {
    const [open, setOpen] = useState<OpenState>(null);
    const [hovered, setHovered] = useState(false);
    const focus = useFocusRing<HTMLButtonElement>();
    const [detected, setDetected] = useState<'menu' | 'dialog'>('menu');
    const [tokens, setTokens] = useState<CSSProperties>({});
    const [position, setPosition] = useState<{
      top: number;
      left: number;
      maxHeight: number;
    } | null>(null);
    const triggerRef = useRef<HTMLButtonElement>(null);
    const popupRef = useRef<HTMLDivElement>(null);
    const chromeRef = useCommandChromeRef();
    const setPopupRef = useCallback(
      (element: HTMLDivElement | null) => {
        popupRef.current = element;
        chromeRef(element);
      },
      [chromeRef]
    );
    const popupId = useId();
    const kind = popup ?? detected;
    const described = useDisabledDescription(disabled, description);
    const roving = useRovingFocus({
      container: popupRef,
      itemSelector: MENU_ITEM_SELECTOR,
      scope: '[role="menu"]',
      labelOf: (item) => item.getAttribute('aria-label') ?? item.textContent ?? '',
    });

    const restoreTarget = () => open?.anchor ?? triggerRef.current;
    const close = useCallback(() => setOpen(null), []);
    const closeAndRestore = () => {
      const target = restoreTarget();
      setOpen(null);
      target?.focus({ preventScroll: true });
    };

    useImperativeHandle(
      ref,
      () => ({
        openAt: (anchor) => {
          if (!disabled) setOpen({ focus: 'first', anchor });
        },
      }),
      [disabled]
    );

    useLayoutEffect(() => {
      const element = popupRef.current;
      if (!open || !element) {
        setPosition(null);
        return;
      }
      if (!popup) {
        const focusable = Array.from(element.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR));
        const menu = focusable.every((item) => item.matches(MENU_ITEM_SELECTOR));
        setDetected(menu ? 'menu' : 'dialog');
      }
      const anchorElement = open.anchor ?? triggerRef.current;
      const anchor = anchorElement?.getBoundingClientRect();
      if (anchor) setPosition(placePopup(anchor, element, menuWidth));
      setTokens(inheritedToolbarTokens(anchorElement));
    }, [open, popup, menuWidth]);

    useLayoutEffect(() => {
      const element = popupRef.current;
      if (!open || !element || !position) return;
      if (element.contains(document.activeElement)) return;
      if (kind === 'menu') {
        if (open.focus === 'first') roving.focusFirst();
        else if (open.focus === 'last') roving.focusLast();
        else element.focus({ preventScroll: true });
      } else {
        const first = element.querySelector<HTMLElement>(FOCUSABLE_SELECTOR);
        (first ?? element).focus({ preventScroll: true });
      }
    }, [open, position, kind, roving]);

    useEffect(() => {
      if (!open) return;
      const onPointerDown = (event: MouseEvent) => {
        const target = event.target as Node;
        if (!triggerRef.current?.contains(target) && !popupRef.current?.contains(target)) close();
      };
      const onScroll = (event: Event) => {
        if (!popupRef.current?.contains(event.target as Node)) close();
      };
      const onResize = () => close();
      document.addEventListener('mousedown', onPointerDown);
      window.addEventListener('scroll', onScroll, true);
      window.addEventListener('resize', onResize);
      return () => {
        document.removeEventListener('mousedown', onPointerDown);
        window.removeEventListener('scroll', onScroll, true);
        window.removeEventListener('resize', onResize);
      };
    }, [open, close]);

    const onPopupKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        closeAndRestore();
        return;
      }
      if (event.key === 'Tab') {
        if (kind === 'menu') setTimeout(close, 0);
        return;
      }
      if (kind === 'menu' && roving.onKeyDown(event)) event.stopPropagation();
    };

    return (
      <div style={{ position: 'relative', display: 'inline-flex', flex: '0 0 auto' }}>
        <button
          ref={triggerRef}
          type="button"
          data-testid={testId}
          {...described.props}
          aria-label={title}
          aria-haspopup={kind}
          aria-expanded={open !== null}
          aria-controls={open ? popupId : undefined}
          title={tooltipText(title, null, described.reason)}
          onMouseDown={(event) => event.preventDefault()}
          onMouseEnter={() => setHovered(true)}
          onMouseLeave={() => setHovered(false)}
          onClick={() => {
            if (!disabled)
              setOpen((current) => (current ? null : { focus: 'popup', anchor: null }));
          }}
          onKeyDown={(event) => {
            if (disabled) return;
            if (event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') {
              event.preventDefault();
              setOpen({ focus: 'first', anchor: null });
            } else if (event.key === 'ArrowUp') {
              event.preventDefault();
              setOpen({ focus: 'last', anchor: null });
            }
          }}
          {...focus.focusProps}
          style={{
            ...toolbarButtonStyle({
              disabled,
              pressed: active,
              open: open !== null,
              hovered,
              focusRing: focus.focusRing,
            }),
            ...style,
          }}
        >
          {trigger}
        </button>
        {described.node}
        {open &&
          createPortal(
            <div
              ref={setPopupRef}
              id={popupId}
              role={kind}
              aria-label={title}
              aria-orientation={kind === 'menu' ? 'vertical' : undefined}
              tabIndex={-1}
              onKeyDown={onPopupKeyDown}
              onBlur={(event) => {
                if (kind !== 'dialog') return;
                const next = event.relatedTarget as Node | null;
                if (
                  next &&
                  !popupRef.current?.contains(next) &&
                  !triggerRef.current?.contains(next)
                ) {
                  close();
                }
              }}
              onMouseDown={(event) => {
                if (kind === 'menu') event.preventDefault();
              }}
              style={{
                position: 'fixed',
                top: position?.top ?? 0,
                left: position?.left ?? 0,
                visibility: position ? 'visible' : 'hidden',
                zIndex: 10000,
                width: Math.min(menuWidth, window.innerWidth - 2 * VIEWPORT_MARGIN),
                maxHeight: position ? Math.min(440, position.maxHeight) : 440,
                overflowY: 'auto',
                ...toolbarMenuStyle,
                ...tokens,
              }}
            >
              {children(() => {
                const target = restoreTarget();
                setOpen(null);
                if (kind === 'menu') target?.focus({ preventScroll: true });
              })}
            </div>,
            document.body
          )}
      </div>
    );
  }
);

/** A button with a popup; hidden at narrow widths, it opens from the overflow menu. */
export function ToolbarDropdown(props: ToolbarDropdownProps) {
  const handle = useRef<DropdownHandle>(null);
  const element = useRef<HTMLSpanElement>(null);
  const registry = useOverflowRegistry();
  const id = useId();
  const { title, disabled = false, description } = props;
  useOverflowSource(element, () => [
    {
      kind: 'item',
      id,
      label: title,
      disabled,
      description: disabled ? description : undefined,
      onSelect: () => handle.current?.openAt(registry?.anchor() ?? null),
    },
  ]);
  return (
    <span ref={element} style={{ display: 'inline-flex', flex: '0 0 auto' }}>
      <ToolbarDropdownBase ref={handle} {...props} />
    </span>
  );
}

export interface ToolbarMenuItemProps {
  label: string;
  icon?: ReactNode;
  selected?: boolean;
  /** Presents the item as one choice of several, checked when `selected`. */
  radio?: boolean;
  disabled?: boolean;
  /** Why the item is disabled. */
  description?: string;
  onClick?: () => void;
  close?: () => void;
}

export function ToolbarMenuItem({
  label,
  icon,
  selected = false,
  radio = false,
  disabled = false,
  description,
  onClick,
  close,
}: ToolbarMenuItemProps) {
  const [hovered, setHovered] = useState(false);
  const itemFocus = useFocusRing<HTMLButtonElement>();
  const descriptionId = useId();
  const reason = disabled ? description : undefined;
  return (
    <>
      <button
        type="button"
        role={radio ? 'menuitemradio' : 'menuitem'}
        tabIndex={-1}
        aria-disabled={disabled || undefined}
        aria-checked={radio ? selected : undefined}
        aria-label={label}
        aria-describedby={reason ? descriptionId : undefined}
        title={reason}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        onClick={() => {
          if (disabled) return;
          onClick?.();
          close?.();
        }}
        {...itemFocus.focusProps}
        style={toolbarMenuItemStyle({ disabled, hovered, focusRing: itemFocus.focusRing })}
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
        <span style={{ flex: 1, overflowWrap: 'anywhere' }}>{label}</span>
        {selected && <ToolbarIcon name="check" size={18} />}
      </button>
      {reason && (
        <span id={descriptionId} hidden>
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
      style={toolbarMenuSeparatorStyle}
    />
  );
}
