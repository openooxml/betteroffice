/**
 * MenuDropdown — a reusable dropdown menu with text label trigger
 *
 * Uses position:fixed so dropdowns escape overflow:auto/hidden ancestors.
 * Supports submenu panels that appear to the right on hover (Google Docs style).
 */

import { useState, useRef, useEffect, useCallback, useId } from 'react';
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
  submenuRole?: 'menu' | 'group';
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

export function MenuDropdown({ label, items, disabled, showChevron = false }: MenuDropdownProps) {
  const menuId = useId();
  const [isOpen, setIsOpen] = useState(false);
  const [openSubmenuLabel, setOpenSubmenuLabel] = useState<string | null>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const focusFirstMenuItem = useRef(false);
  const focusSubmenu = useRef(false);
  const [focusRequest, setFocusRequest] = useState(0);
  const [dropdownPos, setDropdownPos] = useState<{ top: number; left: number }>({
    top: 0,
    left: 0,
  });

  const closeMenu = useCallback(() => {
    setIsOpen(false);
    setOpenSubmenuLabel(null);
  }, []);

  const openSubmenu = (submenuLabel: string) => {
    focusSubmenu.current = true;
    setOpenSubmenuLabel(submenuLabel);
    setFocusRequest((request) => request + 1);
  };

  // Calculate position when opening
  useEffect(() => {
    if (!isOpen || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    setDropdownPos({ top: rect.bottom + 2, left: rect.left });
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen || !focusFirstMenuItem.current) return;
    focusFirstMenuItem.current = false;
    dropdownRef.current
      ?.querySelector<HTMLElement>(
        '[role="menuitem"]:not([aria-disabled="true"]):not(:disabled)'
      )
      ?.focus();
  }, [isOpen, focusRequest]);

  useEffect(() => {
    if (!openSubmenuLabel || !focusSubmenu.current) return;
    focusSubmenu.current = false;
    const panel = Array.from(
      dropdownRef.current?.querySelectorAll<HTMLElement>('[data-submenu-label]') ?? []
    ).find((element) => element.dataset.submenuLabel === openSubmenuLabel);
    panel
      ?.querySelector<HTMLElement>(
        '[role="gridcell"][tabindex="0"], [role="menuitem"]:not([aria-disabled="true"]):not(:disabled)'
      )
      ?.focus();
  }, [openSubmenuLabel, focusRequest]);

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
      if (e.key === 'Escape') {
        closeMenu();
        triggerRef.current?.focus();
      }
    }

    // Close on scroll of any ancestor (dropdown position would be stale)
    function handleScroll() {
      closeMenu();
    }

    document.addEventListener('mousedown', handleClickOutside);
    document.addEventListener('keydown', handleEscape);
    window.addEventListener('scroll', handleScroll, true);
    return () => {
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

  const handleMenuKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.target instanceof Element && event.target.closest('[role="menu"]') !== event.currentTarget) {
      return;
    }
    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      closeMenu();
      triggerRef.current?.focus();
      return;
    }
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const menuItems = Array.from(
      event.currentTarget.querySelectorAll<HTMLElement>('[role="menuitem"]')
    ).filter(
      (item) =>
        item.closest('[role="menu"]') === event.currentTarget &&
        item.getAttribute('aria-disabled') !== 'true' &&
        !item.hasAttribute('disabled')
    );
    if (menuItems.length === 0) return;
    const active =
      event.target instanceof Element
        ? event.target.closest<HTMLElement>('[role="menuitem"]')
        : null;
    const currentIndex = menuItems.indexOf(active as HTMLElement);
    const index =
      event.key === 'Home'
        ? 0
        : event.key === 'End'
          ? menuItems.length - 1
          : event.key === 'ArrowDown'
            ? (currentIndex + 1 + menuItems.length) % menuItems.length
            : (currentIndex <= 0 ? menuItems.length : currentIndex) - 1;
    menuItems[index]?.focus();
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
        onKeyDown={(event) => {
          if (!disabled && (event.key === 'ArrowRight' || event.key === 'ArrowDown')) {
            event.preventDefault();
            focusFirstMenuItem.current = true;
            setFocusRequest((request) => request + 1);
            setIsOpen(true);
          } else if (!isOpen && (event.key === 'Enter' || event.key === ' ')) {
            focusFirstMenuItem.current = true;
            setFocusRequest((request) => request + 1);
          }
        }}
        onClick={() => !disabled && setIsOpen(!isOpen)}
        onMouseDown={(e) => e.preventDefault()}
        disabled={disabled}
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
              return <div key={`sep-${i}`} style={separatorStyle} />;
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
                onMouseEnter={() => hasSubmenu && setOpenSubmenuLabel(item.label)}
                onMouseLeave={() => hasSubmenu && setOpenSubmenuLabel(null)}
              >
                <button
                  type="button"
                  role="menuitem"
                  aria-haspopup={
                    hasSubmenu ? (item.submenuRole === 'group' ? 'grid' : 'menu') : undefined
                  }
                  aria-expanded={hasSubmenu ? isSubmenuOpen : undefined}
                  aria-controls={hasSubmenu ? submenuId : undefined}
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
                  disabled={item.disabled && !reason}
                  aria-disabled={reason ? true : undefined}
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
                    id={submenuId}
                    role={item.submenuRole ?? 'menu'}
                    aria-label={item.label}
                    data-submenu-label={item.label}
                    style={submenuPanelStyle}
                    onMouseDown={(e) => e.preventDefault()}
                    onKeyDown={(event) => {
                      if (event.key !== 'ArrowLeft') return;
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
