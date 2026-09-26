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
import type {
  CSSProperties,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
  ReactNode,
} from 'react';
import { useRovingFocus } from './useRovingFocus';

/** One entry of an overflow menu; plain render data without behavior beyond `onSelect`. */
export type OverflowMenuEntry =
  | {
      kind: 'item';
      id: string;
      label: string;
      icon?: ReactNode;
      shortcut?: string;
      /** Renders a checkbox item with this state. */
      checked?: boolean | 'mixed';
      /** Renders a radio item; `checked` marks the chosen one. */
      radio?: boolean;
      disabled?: boolean;
      /** Explains the disabled state to assistive technology and on hover. */
      description?: string;
      onSelect(): void;
    }
  | {
      kind: 'submenu';
      id: string;
      label: string;
      icon?: ReactNode;
      disabled?: boolean;
      description?: string;
      entries: readonly OverflowMenuEntry[];
    }
  | { kind: 'group'; id: string; label?: string; entries: readonly OverflowMenuEntry[] }
  | { kind: 'separator'; id: string };

export interface OverflowMenuColors {
  surface: string;
  text: string;
  mutedText: string;
  border: string;
  hover: string;
  shadow: string;
}

export interface OverflowMenuHandle {
  focus(): void;
  close(): void;
}

export interface OverflowMenuProps {
  /** Accessible name of the trigger and the menu. */
  label: string;
  trigger: ReactNode;
  entries: readonly OverflowMenuEntry[];
  colors?: Partial<OverflowMenuColors>;
  triggerClassName?: string;
  triggerStyle?: CSSProperties;
  testId?: string;
  /** Called after an item ran, with the input that chose it. */
  onItemSelected?(modality: 'pointer' | 'keyboard'): void;
}

const DEFAULT_COLORS: OverflowMenuColors = {
  surface: '#ffffff',
  text: '#202124',
  mutedText: '#5f6368',
  border: '#dadce0',
  hover: '#f1f3f4',
  shadow: 'rgba(60, 64, 67, 0.24)',
};

const ITEM_SELECTOR = '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]';
const VIEWPORT_MARGIN = 8;
const MENU_MAX_WIDTH = 320;

type Placement = { top: number; left: number; maxHeight: number };

function place(anchor: DOMRect, menu: HTMLElement, beside: boolean): Placement {
  const width = menu.offsetWidth;
  const height = menu.scrollHeight;
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  let left = beside ? anchor.right : anchor.right - width;
  if (beside && left + width > viewportWidth - VIEWPORT_MARGIN) left = anchor.left - width;
  left = Math.max(VIEWPORT_MARGIN, Math.min(left, viewportWidth - width - VIEWPORT_MARGIN));
  const below = viewportHeight - VIEWPORT_MARGIN - (beside ? anchor.top : anchor.bottom + 4);
  const above = (beside ? anchor.bottom : anchor.top - 4) - VIEWPORT_MARGIN;
  if (height <= below || below >= above) {
    const top = beside ? anchor.top : anchor.bottom + 4;
    return { top, left, maxHeight: Math.max(below, 80) };
  }
  const maxHeight = Math.max(above, 80);
  const bottom = beside ? anchor.bottom : anchor.top - 4;
  return { top: Math.max(VIEWPORT_MARGIN, bottom - Math.min(height, maxHeight)), left, maxHeight };
}

function CheckGlyph() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
      <path d="M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z" fill="currentColor" />
    </svg>
  );
}

function ChevronGlyph() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
      <path d="M10 6 8.6 7.4 13.2 12l-4.6 4.6L10 18l6-6z" fill="currentColor" />
    </svg>
  );
}

type InitialFocus = 'first' | 'last' | 'menu' | 'none';

interface MenuListProps {
  id?: string;
  label: string;
  entries: readonly OverflowMenuEntry[];
  anchor: () => DOMRect | null;
  beside: boolean;
  colors: OverflowMenuColors;
  initialFocus: InitialFocus;
  onClose(reason: 'escape' | 'tab' | 'parent'): void;
  onSelect(entry: Extract<OverflowMenuEntry, { kind: 'item' }>, modality: 'pointer' | 'keyboard'): void;
  registerMenu(element: HTMLElement): () => void;
}

function MenuList({
  id,
  label,
  entries,
  anchor,
  beside,
  colors,
  initialFocus,
  onClose,
  onSelect,
  registerMenu,
}: MenuListProps) {
  const menuRef = useRef<HTMLDivElement>(null);
  const baseId = useId();
  const [placement, setPlacement] = useState<Placement | null>(null);
  const focused = useRef(false);
  const [submenu, setSubmenu] = useState<{ id: string; focus: InitialFocus } | null>(null);
  const submenuAnchors = useRef(new Map<string, HTMLElement>());
  const roving = useRovingFocus({
    container: menuRef,
    itemSelector: ITEM_SELECTOR,
    scope: '[role="menu"]',
    labelOf: (item) => item.dataset.label ?? item.textContent ?? '',
  });

  useLayoutEffect(() => (menuRef.current ? registerMenu(menuRef.current) : undefined), [registerMenu]);

  useLayoutEffect(() => {
    const rect = anchor();
    const menu = menuRef.current;
    if (!rect || !menu) return;
    const next = place(rect, menu, beside);
    setPlacement((current) =>
      current &&
      current.top === next.top &&
      current.left === next.left &&
      current.maxHeight === next.maxHeight
        ? current
        : next
    );
  }, [anchor, beside, entries]);

  useLayoutEffect(() => {
    if (!placement || focused.current) return;
    focused.current = true;
    if (initialFocus === 'first') roving.focusFirst();
    else if (initialFocus === 'last') roving.focusLast();
    else if (initialFocus === 'menu') menuRef.current?.focus({ preventScroll: true });
  }, [initialFocus, placement, roving]);

  const openSubmenu = (entryId: string, focus: InitialFocus) =>
    setSubmenu({ id: entryId, focus });

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.target instanceof Node && !menuRef.current?.contains(event.target)) return;
    const item =
      event.target instanceof HTMLElement ? event.target.closest<HTMLElement>(ITEM_SELECTOR) : null;
    const ownItem = item && item.closest('[role="menu"]') === menuRef.current ? item : null;
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onClose(beside ? 'parent' : 'escape');
      return;
    }
    if (event.key === 'Tab') {
      setTimeout(() => onClose('tab'), 0);
      return;
    }
    if (beside && event.key === 'ArrowLeft') {
      event.preventDefault();
      event.stopPropagation();
      onClose('parent');
      return;
    }
    if (ownItem?.dataset.submenu && (event.key === 'ArrowRight' || event.key === 'Enter' || event.key === ' ')) {
      event.preventDefault();
      event.stopPropagation();
      if (ownItem.getAttribute('aria-disabled') !== 'true') openSubmenu(ownItem.dataset.submenu, 'first');
      return;
    }
    if (ownItem && (event.key === 'Enter' || event.key === ' ')) {
      event.preventDefault();
      event.stopPropagation();
      const entry = findItem(entries, ownItem.dataset.entry ?? '');
      if (entry && !entry.disabled) onSelect(entry, 'keyboard');
      return;
    }
    if (roving.onKeyDown(event)) event.stopPropagation();
  };

  const renderEntries = (list: readonly OverflowMenuEntry[]): ReactNode =>
    list.map((entry) => {
      if (entry.kind === 'separator') {
        return (
          <div
            key={entry.id}
            role="separator"
            style={{ height: 1, margin: '4px 0', background: colors.border }}
          />
        );
      }
      if (entry.kind === 'group') {
        const labelId = `${baseId}-${entry.id}-label`;
        return (
          <div
            key={entry.id}
            role="group"
            aria-labelledby={entry.label ? labelId : undefined}
          >
            {entry.label && (
              <div
                id={labelId}
                style={{
                  padding: '6px 12px 2px',
                  fontSize: 11,
                  fontWeight: 500,
                  color: colors.mutedText,
                  overflowWrap: 'anywhere',
                }}
              >
                {entry.label}
              </div>
            )}
            {renderEntries(entry.entries)}
          </div>
        );
      }
      const descriptionId = entry.description ? `${baseId}-${entry.id}-description` : undefined;
      const isSubmenu = entry.kind === 'submenu';
      const role = isSubmenu
        ? 'menuitem'
        : entry.radio
          ? 'menuitemradio'
          : entry.checked !== undefined
            ? 'menuitemcheckbox'
            : 'menuitem';
      const checked = !isSubmenu ? entry.checked : undefined;
      const expanded = isSubmenu && submenu?.id === entry.id;
      return (
        <div key={entry.id} style={{ position: 'relative' }}>
          <div
            role={role}
            tabIndex={-1}
            data-entry={entry.id}
            data-label={entry.label}
            data-submenu={isSubmenu ? entry.id : undefined}
            aria-haspopup={isSubmenu ? 'menu' : undefined}
            aria-expanded={isSubmenu ? expanded : undefined}
            aria-checked={
              role === 'menuitemcheckbox' || role === 'menuitemradio'
                ? checked === 'mixed'
                  ? 'mixed'
                  : checked === true
                : undefined
            }
            aria-disabled={entry.disabled ? true : undefined}
            aria-describedby={descriptionId}
            title={entry.description}
            ref={(element) => {
              if (isSubmenu && element) submenuAnchors.current.set(entry.id, element);
            }}
            onMouseDown={(event: ReactMouseEvent) => event.preventDefault()}
            onMouseEnter={(event) => {
              if (isSubmenu && !entry.disabled) openSubmenu(entry.id, 'none');
              else if (submenu) setSubmenu(null);
              (event.currentTarget as HTMLElement).style.background = entry.disabled
                ? 'transparent'
                : colors.hover;
            }}
            onMouseLeave={(event) => {
              (event.currentTarget as HTMLElement).style.background = 'transparent';
            }}
            onClick={() => {
              if (entry.disabled) return;
              if (isSubmenu) openSubmenu(entry.id, 'none');
              else onSelect(entry, 'pointer');
            }}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 8,
              minHeight: 30,
              padding: '4px 12px 4px 8px',
              fontSize: 13,
              cursor: entry.disabled ? 'default' : 'pointer',
              color: entry.disabled ? colors.mutedText : colors.text,
              opacity: entry.disabled ? 0.6 : 1,
              outlineOffset: -2,
            }}
          >
            <span
              style={{ width: 16, flexShrink: 0, display: 'inline-flex', justifyContent: 'center' }}
            >
              {checked === true ? <CheckGlyph /> : checked === 'mixed' ? '–' : null}
            </span>
            {entry.icon && (
              <span aria-hidden="true" style={{ display: 'inline-flex', flexShrink: 0 }}>
                {entry.icon}
              </span>
            )}
            <span style={{ flex: 1, minWidth: 0, overflowWrap: 'anywhere' }}>{entry.label}</span>
            {!isSubmenu && entry.shortcut && (
              <span
                aria-hidden="true"
                style={{
                  marginLeft: 16,
                  flexShrink: 0,
                  whiteSpace: 'nowrap',
                  fontSize: 12,
                  color: colors.mutedText,
                }}
              >
                {entry.shortcut}
              </span>
            )}
            {isSubmenu && (
              <span style={{ display: 'inline-flex', flexShrink: 0 }}>
                <ChevronGlyph />
              </span>
            )}
          </div>
          {descriptionId && (
            <span id={descriptionId} hidden>
              {entry.description}
            </span>
          )}
          {isSubmenu && expanded && (
            <MenuList
              label={entry.label}
              entries={entry.entries}
              anchor={() => submenuAnchors.current.get(entry.id)?.getBoundingClientRect() ?? null}
              beside
              colors={colors}
              initialFocus={submenu?.focus ?? 'none'}
              onClose={(reason) => {
                setSubmenu(null);
                if (reason === 'parent') submenuAnchors.current.get(entry.id)?.focus();
                else onClose(reason);
              }}
              onSelect={onSelect}
              registerMenu={registerMenu}
            />
          )}
        </div>
      );
    });

  return (
    <div
      ref={menuRef}
      id={id}
      role="menu"
      aria-label={label}
      aria-orientation="vertical"
      tabIndex={-1}
      onKeyDown={handleKeyDown}
      style={{
        position: 'fixed',
        top: placement?.top ?? 0,
        left: placement?.left ?? 0,
        maxHeight: placement?.maxHeight,
        visibility: placement ? 'visible' : 'hidden',
        overflowY: 'auto',
        overflowX: 'hidden',
        minWidth: Math.min(200, window.innerWidth - 2 * VIEWPORT_MARGIN),
        maxWidth: Math.min(MENU_MAX_WIDTH, window.innerWidth - 2 * VIEWPORT_MARGIN),
        padding: '4px 0',
        background: colors.surface,
        color: colors.text,
        border: `1px solid ${colors.border}`,
        borderRadius: 8,
        boxShadow: `0 4px 16px ${colors.shadow}`,
        zIndex: 10000,
        boxSizing: 'border-box',
        outline: 'none',
      }}
    >
      {renderEntries(entries)}
    </div>
  );
}

function findItem(
  entries: readonly OverflowMenuEntry[],
  id: string
): Extract<OverflowMenuEntry, { kind: 'item' }> | null {
  for (const entry of entries) {
    if (entry.kind === 'item' && entry.id === id) return entry;
    if (entry.kind === 'group') {
      const found = findItem(entry.entries, id);
      if (found) return found;
    }
  }
  return null;
}

/**
 * A menu button whose popup follows the WAI-ARIA menu pattern: arrow keys,
 * Home/End and typeahead move between items, Right/Left open and close
 * submenus, Escape restores the trigger and Tab leaves the menu.
 */
export const OverflowMenu = forwardRef<OverflowMenuHandle, OverflowMenuProps>(function OverflowMenu(
  { label, trigger, entries, colors, triggerClassName, triggerStyle, testId, onItemSelected },
  ref
) {
  const triggerRef = useRef<HTMLButtonElement>(null);
  const menus = useRef(new Set<HTMLElement>());
  const menuId = useId();
  const [open, setOpen] = useState<InitialFocus | null>(null);
  const palette = { ...DEFAULT_COLORS, ...colors };

  const close = useCallback(() => setOpen(null), []);
  useImperativeHandle(
    ref,
    () => ({
      focus: () => triggerRef.current?.focus({ preventScroll: true }),
      close,
    }),
    [close]
  );

  const registerMenu = useCallback((element: HTMLElement) => {
    menus.current.add(element);
    return () => {
      menus.current.delete(element);
    };
  }, []);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (triggerRef.current?.contains(target)) return;
      for (const menu of menus.current) if (menu.contains(target)) return;
      close();
    };
    const onResize = () => {
      const active = document.activeElement;
      const focused = [...menus.current].some((menu) => menu.contains(active));
      close();
      if (focused) triggerRef.current?.focus({ preventScroll: true });
    };
    document.addEventListener('mousedown', onPointerDown, true);
    window.addEventListener('resize', onResize);
    return () => {
      document.removeEventListener('mousedown', onPointerDown, true);
      window.removeEventListener('resize', onResize);
    };
  }, [open, close]);

  const anchor = useCallback(() => triggerRef.current?.getBoundingClientRect() ?? null, []);

  return (
    <span style={{ display: 'inline-flex', flex: '0 0 auto' }}>
      <button
        ref={triggerRef}
        type="button"
        className={triggerClassName}
        style={triggerStyle}
        data-testid={testId}
        aria-label={label}
        title={label}
        aria-haspopup="menu"
        aria-expanded={open !== null}
        aria-controls={open ? menuId : undefined}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => setOpen((current) => (current ? null : 'menu'))}
        onKeyDown={(event) => {
          if (event.key === 'Enter' || event.key === ' ' || event.key === 'ArrowDown') {
            event.preventDefault();
            setOpen('first');
          } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            setOpen('last');
          }
        }}
      >
        {trigger}
      </button>
      {open && (
        <MenuList
            id={menuId}
            label={label}
            entries={entries}
            anchor={anchor}
            beside={false}
            colors={palette}
            initialFocus={open}
            onClose={(reason) => {
              close();
              if (reason !== 'tab') triggerRef.current?.focus({ preventScroll: true });
            }}
            onSelect={(entry, modality) => {
              close();
              if (modality === 'keyboard') triggerRef.current?.focus({ preventScroll: true });
              entry.onSelect();
              onItemSelected?.(modality);
            }}
            registerMenu={registerMenu}
          />
      )}
    </span>
  );
});
