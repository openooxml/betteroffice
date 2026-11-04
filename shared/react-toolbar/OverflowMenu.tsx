import { forwardRef, useCallback, useEffect, useId, useImperativeHandle, useRef, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { useFocusRing } from './theme';
import type { ToolbarControlState } from './theme';
import { DEFAULT_MENU_COLORS } from './menuColors';
import { loadedMenuList, loadMenuList, preloadMenuList } from './menuListLoader';
import type { InitialFocus, OverflowMenuColors, OverflowMenuEntry } from './MenuList';

export type { OverflowMenuColors, OverflowMenuEntry } from './MenuList';

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
  /** A style, or one built from the trigger's hover, open and keyboard-focus state. */
  triggerStyle?: CSSProperties | ((state: ToolbarControlState) => CSSProperties);
  testId?: string;
  /** Called after an item ran, with the input that chose it. */
  onItemSelected?(modality: 'pointer' | 'keyboard'): void;
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
  const [, listArrived] = useState(0);
  const MenuList = loadedMenuList();
  const [hovered, setHovered] = useState(false);
  const focus = useFocusRing<HTMLButtonElement>();
  const palette = { ...DEFAULT_MENU_COLORS, ...colors };

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
    const onKeyDown = (event: KeyboardEvent) => {
      if (!loadedMenuList() && (event.key === 'Escape' || event.key === 'Tab')) close();
    };
    document.addEventListener('mousedown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown, true);
    window.addEventListener('resize', onResize);
    return () => {
      document.removeEventListener('mousedown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown, true);
      window.removeEventListener('resize', onResize);
    };
  }, [open, close]);

  useEffect(() => preloadMenuList(), []);
  useEffect(() => {
    if (!open || loadedMenuList()) return;
    let alive = true;
    loadMenuList().then(() => { if (alive) listArrived((count) => count + 1); }, () => {
      if (!alive) return;
      close();
      triggerRef.current?.focus({ preventScroll: true });
    });
    return () => { alive = false; };
  }, [open, close]);

  const anchor = useCallback(() => triggerRef.current?.getBoundingClientRect() ?? null, []);

  return (
    <span style={{ display: 'inline-flex', flex: '0 0 auto' }}>
      <button
        ref={triggerRef}
        type="button"
        className={triggerClassName}
        style={
          typeof triggerStyle === 'function'
            ? triggerStyle({ hovered, open: open !== null, focusRing: focus.focusRing })
            : triggerStyle
        }
        {...focus.focusProps}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
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
      {open && MenuList && (
        <MenuList
            id={menuId}
            label={label}
            entries={entries}
            anchor={anchor}
            mode="below"
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
