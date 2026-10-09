/**
 * MenuDropdown — a reusable dropdown menu with text label trigger
 *
 * Uses position:fixed so dropdowns escape overflow:auto/hidden ancestors.
 * Supports submenu panels that appear to the right on hover (Google Docs style).
 */

import { useState, useRef, useEffect, useLayoutEffect, useCallback, useId } from 'react';
import type { CSSProperties, KeyboardEvent, ReactNode } from 'react';
import { MaterialSymbol } from './MaterialSymbol';

export interface MenuItem {
  icon?: string;
  label: string;
  shortcut?: string;
  onClick?: () => void;
  disabled?: boolean;
  /** Why the item is disabled. */
  description?: string;
  /** Custom content to render instead of a simple menu item */
  customContent?: ReactNode;
  /** Submenu content that appears to the right on hover */
  submenuContent?: (closeMenu: () => void) => ReactNode;
  submenuRole?: 'menu' | 'dialog';
}

export interface MenuSeparator {
  type: 'separator';
}

export type MenuEntry = MenuItem | MenuSeparator;

function isSeparator(entry: MenuEntry): entry is MenuSeparator {
  return 'type' in entry && entry.type === 'separator';
}

interface MenuDropdownProps {
  label: string;
  items: MenuEntry[];
  disabled?: boolean;
  tabIndex?: number;
  onFocus?: () => void;
  /** When true, the trigger renders a down-arrow caret next to the label.
   *  Default `false` — every in-tree caller is a top-level menubar button. */
  showChevron?: boolean;
}

const triggerStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 2,
  padding: '2px 8px',
  border: 'none',
  background: 'transparent',
  borderRadius: 4,
  cursor: 'pointer',
  fontSize: 13,
  fontWeight: 400,
  color: 'var(--doc-text)',
  whiteSpace: 'nowrap',
  height: 28,
  lineHeight: '28px',
};

const triggerOpenStyle: CSSProperties = {
  ...triggerStyle,
  background: 'var(--doc-bg-hover)',
};

const menuItemStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '6px 12px',
  border: 'none',
  background: 'transparent',
  cursor: 'pointer',
  fontSize: 13,
  color: 'var(--doc-text)',
  width: '100%',
  textAlign: 'left',
  whiteSpace: 'nowrap',
};

const menuItemDisabledStyle: CSSProperties = {
  ...menuItemStyle,
  opacity: 0.4,
  cursor: 'default',
};

const separatorStyle: CSSProperties = {
  height: 1,
  backgroundColor: 'var(--doc-border)',
  margin: '4px 0',
};

const shortcutStyle: CSSProperties = {
  marginLeft: 'auto',
  fontSize: 12,
  color: 'var(--doc-text-muted)',
};

const submenuPanelStyle: CSSProperties = {
  position: 'absolute',
  left: '100%',
  top: -4,
  marginLeft: 2,
  backgroundColor: 'var(--doc-surface)',
  border: '1px solid var(--doc-border)',
  borderRadius: 6,
  boxShadow: '0 4px 12px var(--doc-shadow)',
  padding: 8,
  zIndex: 1001,
};

function focusMenuItem(menu: HTMLElement, key: string) {
  const items = Array.from(menu.querySelectorAll<HTMLElement>('[role="menuitem"]')).filter(
    (item) => item.closest('[role="menu"]') === menu
  );
  if (items.length === 0) return;
  const current = items.indexOf(document.activeElement as HTMLElement);
  const index =
    key === 'Home'
      ? 0
      : key === 'End'
      ? items.length - 1
      : key === 'ArrowDown'
      ? (current + 1) % items.length
      : (current <= 0 ? items.length : current) - 1;
  items[index]?.focus();
}

export function MenuDropdown({
  label,
  items,
  disabled,
  tabIndex = 0,
  onFocus,
  showChevron = false,
}: MenuDropdownProps) {
  const menuId = useId();
  const [isOpen, setIsOpen] = useState(false);
  const [openSubmenuLabel, setOpenSubmenuLabel] = useState<string | null>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const submenuRef = useRef<HTMLDivElement>(null);
  const pendingFocus = useRef<'Home' | 'End' | 'submenu' | null>(null);
  const [dropdownPos, setDropdownPos] = useState<{ top: number; left: number }>({
    top: 0,
    left: 0,
  });

  const closeMenu = useCallback(() => {
    if (dropdownRef.current?.contains(document.activeElement)) triggerRef.current?.focus();
    pendingFocus.current = null;
    setIsOpen(false);
    setOpenSubmenuLabel(null);
  }, []);

  const focusSubmenu = () => {
    submenuRef.current
      ?.querySelector<HTMLElement>('[role="gridcell"][tabindex="0"], [role="menuitem"]')
      ?.focus();
  };

  const openMenu = (key: 'Home' | 'End') => {
    if (isOpen && dropdownRef.current) {
      focusMenuItem(dropdownRef.current, key);
    } else {
      pendingFocus.current = key;
      setIsOpen(true);
    }
  };

  const openSubmenu = (submenuLabel: string) => {
    if (openSubmenuLabel === submenuLabel) {
      focusSubmenu();
    } else {
      pendingFocus.current = 'submenu';
      setOpenSubmenuLabel(submenuLabel);
    }
  };

  useLayoutEffect(() => {
    if (!isOpen || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    setDropdownPos({ top: rect.bottom + 2, left: rect.left });
  }, [isOpen]);

  useLayoutEffect(() => {
    const target = pendingFocus.current;
    pendingFocus.current = null;
    if (target === 'submenu') focusSubmenu();
    else if (target && dropdownRef.current) focusMenuItem(dropdownRef.current, target);
  });

  useEffect(() => {
    if (!isOpen) return;

    function handleClickOutside(e: MouseEvent) {
      const target = e.target as Node;
      if (
        triggerRef.current &&
        !triggerRef.current.contains(target) &&
        dropdownRef.current &&
        !dropdownRef.current.contains(target)
      ) {
        closeMenu();
      }
    }

    function handleEscape(e: globalThis.KeyboardEvent) {
      if (e.key === 'Escape' && !e.defaultPrevented) {
        e.preventDefault();
        closeMenu();
        triggerRef.current?.focus();
      }
    }

    function handleScroll() {
      closeMenu();
    }

    function handleFocusOutside(e: FocusEvent) {
      const target = e.target as Node;
      if (!triggerRef.current?.contains(target) && !dropdownRef.current?.contains(target)) {
        closeMenu();
      }
    }

    document.addEventListener('focusin', handleFocusOutside);
    document.addEventListener('mousedown', handleClickOutside);
    document.addEventListener('keydown', handleEscape);
    window.addEventListener('scroll', handleScroll, true);
    return () => {
      document.removeEventListener('focusin', handleFocusOutside);
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleEscape);
      window.removeEventListener('scroll', handleScroll, true);
    };
  }, [isOpen, closeMenu]);

  const handleItemClick = (item: MenuItem) => {
    if (item.disabled || item.submenuContent) return;
    if (!item.onClick) return;
    item.onClick();
    closeMenu();
  };

  const moveMenubarFocus = (key: string, open = false) => {
    const menubar = triggerRef.current?.closest('[role="menubar"]');
    if (!menubar) return false;
    const triggers = Array.from(
      menubar.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')
    ).filter((item) => item.closest('[role="menu"], [role="menubar"]') === menubar);
    const current = triggers.indexOf(triggerRef.current!);
    const index =
      key === 'Home'
        ? 0
        : key === 'End'
        ? triggers.length - 1
        : (current + (key === 'ArrowRight' ? 1 : -1) + triggers.length) % triggers.length;
    const next = triggers[index];
    if (next === triggerRef.current) {
      next?.focus();
      return true;
    }
    closeMenu();
    next?.focus();
    if (open) next?.click();
    return true;
  };

  const handleMenuKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Tab') {
      closeMenu();
      return;
    }
    if (
      event.target instanceof Element &&
      event.target.closest('[role="menu"]') !== event.currentTarget
    )
      return;
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      closeMenu();
      triggerRef.current?.focus();
    } else if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
      event.preventDefault();
      event.stopPropagation();
      if (!moveMenubarFocus(event.key, true) && event.key === 'ArrowLeft') closeMenu();
    } else if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) {
      event.preventDefault();
      event.stopPropagation();
      focusMenuItem(event.currentTarget, event.key);
    }
  };

  return (
    <div style={{ position: 'relative' }}>
      <button
        ref={triggerRef}
        type="button"
        role="menuitem"
        aria-haspopup="menu"
        aria-expanded={isOpen}
        aria-controls={menuId}
        tabIndex={tabIndex}
        onFocus={onFocus}
        onKeyDown={(event) => {
          if (event.key === 'Tab') {
            closeMenu();
            return;
          }
          if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) {
            if (moveMenubarFocus(event.key)) {
              event.preventDefault();
              return;
            }
          }
          if (disabled) return;
          if (['Enter', ' ', 'ArrowDown', 'ArrowUp', 'ArrowRight'].includes(event.key)) {
            event.preventDefault();
            openMenu(event.key === 'ArrowUp' ? 'End' : 'Home');
          }
        }}
        onClick={() => {
          if (disabled) return;
          if (isOpen) closeMenu();
          else setIsOpen(true);
        }}
        onMouseDown={(e) => e.preventDefault()}
        aria-disabled={disabled || undefined}
        style={isOpen ? triggerOpenStyle : triggerStyle}
      >
        {label}
        {showChevron && <MaterialSymbol name="arrow_drop_down" size={16} />}
      </button>

      {isOpen && (
        <div
          ref={dropdownRef}
          id={menuId}
          role="menu"
          aria-label={label}
          data-docx-escape-layer="true"
          style={{
            position: 'fixed',
            top: dropdownPos.top,
            left: dropdownPos.left,
            backgroundColor: 'var(--doc-surface)',
            border: '1px solid var(--doc-border)',
            borderRadius: 6,
            boxShadow: '0 4px 12px var(--doc-shadow)',
            padding: '4px 0',
            zIndex: 10000,
            minWidth: 200,
          }}
          onMouseDown={(e) => e.preventDefault()}
          onKeyDown={handleMenuKeyDown}
        >
          {items.map((entry, i) => {
            if (isSeparator(entry)) {
              return <div key={`sep-${i}`} role="separator" style={separatorStyle} />;
            }
            const item = entry;
            if (item.customContent) {
              return (
                <div key={item.label} onMouseDown={(e) => e.preventDefault()}>
                  {item.customContent}
                </div>
              );
            }

            const hasSubmenu = !!item.submenuContent;
            const isSubmenuOpen = openSubmenuLabel === item.label;
            const reason = item.disabled ? item.description : undefined;
            const reasonId = reason ? `${menuId}-${i}` : undefined;
            const submenuId = `${menuId}-submenu-${i}`;

            return (
              <div
                key={item.label}
                style={{ position: 'relative' }}
                onMouseEnter={(event) => {
                  if (!hasSubmenu || item.disabled) return;
                  if (submenuRef.current?.contains(document.activeElement) && !isSubmenuOpen) {
                    event.currentTarget
                      .querySelector<HTMLButtonElement>('button[aria-haspopup]')
                      ?.focus();
                  }
                  setOpenSubmenuLabel(item.label);
                }}
                onMouseLeave={(event) => {
                  if (isSubmenuOpen && !event.currentTarget.contains(document.activeElement))
                    setOpenSubmenuLabel(null);
                }}
              >
                <button
                  type="button"
                  role="menuitem"
                  aria-haspopup={hasSubmenu ? item.submenuRole ?? 'menu' : undefined}
                  aria-expanded={hasSubmenu ? isSubmenuOpen : undefined}
                  aria-controls={hasSubmenu ? submenuId : undefined}
                  tabIndex={-1}
                  onFocus={() => {
                    if (!isSubmenuOpen) setOpenSubmenuLabel(null);
                  }}
                  style={item.disabled ? menuItemDisabledStyle : menuItemStyle}
                  onKeyDown={(event) => {
                    if (
                      hasSubmenu &&
                      !item.disabled &&
                      (event.key === 'ArrowRight' || event.key === 'Enter' || event.key === ' ')
                    ) {
                      event.preventDefault();
                      event.stopPropagation();
                      openSubmenu(item.label);
                    }
                  }}
                  onClick={() => handleItemClick(item)}
                  onMouseDown={(e) => e.preventDefault()}
                  onMouseOver={(e) => {
                    if (!item.disabled) {
                      (e.currentTarget as HTMLButtonElement).style.backgroundColor =
                        'var(--doc-bg-hover)';
                    }
                  }}
                  onMouseOut={(e) => {
                    (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
                  }}
                  aria-disabled={item.disabled || undefined}
                  aria-describedby={reasonId}
                  title={reason}
                >
                  {reason && (
                    <span id={reasonId} hidden>
                      {reason}
                    </span>
                  )}
                  {item.icon && <MaterialSymbol name={item.icon} size={18} />}
                  <span>{item.label}</span>
                  {item.shortcut && <span style={shortcutStyle}>{item.shortcut}</span>}
                  {hasSubmenu && (
                    <span style={{ marginLeft: 'auto' }}>
                      <MaterialSymbol name="keyboard_arrow_right" size={16} />
                    </span>
                  )}
                </button>
                {hasSubmenu && isSubmenuOpen && (
                  <div
                    ref={submenuRef}
                    id={submenuId}
                    role={item.submenuRole ?? 'menu'}
                    aria-label={item.label}
                    style={submenuPanelStyle}
                    onMouseDown={(e) => e.preventDefault()}
                    onKeyDown={(event) => {
                      if (event.key !== 'ArrowLeft' && event.key !== 'Escape') {
                        if (item.submenuRole !== 'dialog') handleMenuKeyDown(event);
                        return;
                      }
                      event.preventDefault();
                      event.stopPropagation();
                      setOpenSubmenuLabel(null);
                      event.currentTarget.parentElement
                        ?.querySelector<HTMLButtonElement>('button[aria-haspopup]')
                        ?.focus();
                    }}
                  >
                    {item.submenuContent!(closeMenu)}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
