import { useCallback, useEffect, useLayoutEffect, useReducer, useRef, useState } from 'react';
import type { ReactNode, RefObject } from 'react';
import { DEFAULT_MENU_COLORS } from './menuColors';
import { guardPendingMenu, loadedMenuList, loadMenuList, preloadMenuList } from './menuListLoader';
import type { OverflowMenuColors, OverflowMenuEntry } from './MenuList';

/** How a context menu was asked for. */
export type ContextMenuModality = 'pointer' | 'keyboard';

export type ContextMenuCloseReason = 'escape' | 'tab' | 'outside' | 'select' | 'resize';

export interface ContextMenuProps {
  /** Accessible name of the menu. */
  label: string;
  entries: readonly OverflowMenuEntry[];
  /** Viewport point the menu opens at; it flips to stay on screen. */
  point: { x: number; y: number };
  /** Keyboard-opened menus focus their first item; pointer-opened ones focus the menu. */
  modality?: ContextMenuModality;
  colors?: Partial<OverflowMenuColors>;
  className?: string;
  testId?: string;
  /** Where Escape and choosing an item send focus; defaults to the element focused before the menu opened. */
  returnFocus?(): void;
  onClose(reason: ContextMenuCloseReason): void;
}

/**
 * A context menu at a point, following the WAI-ARIA menu pattern. Escape and
 * choosing an item return focus to the element focused before it opened; a
 * press outside closes it in the same turn without moving focus.
 */
export function ContextMenu({
  label,
  entries,
  point,
  modality = 'pointer',
  colors,
  className,
  testId,
  returnFocus,
  onClose,
}: ContextMenuProps) {
  const [origin] = useState(() =>
    typeof document !== 'undefined' && document.activeElement instanceof HTMLElement
      ? document.activeElement
      : null
  );
  const menus = useRef(new Set<HTMLElement>());
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  const returnRef = useRef(returnFocus);
  returnRef.current = returnFocus;
  const closed = useRef(false);
  const [, listArrived] = useState(0);
  const MenuList = loadedMenuList();

  const finish = useCallback(
    (reason: ContextMenuCloseReason, restore: boolean) => {
      if (closed.current) return;
      closed.current = true;
      if (restore && returnRef.current) returnRef.current();
      else if (restore && origin?.isConnected) origin.focus({ preventScroll: true });
      closeRef.current(reason);
    },
    [origin]
  );

  useEffect(() => {
    if (MenuList) return;
    let alive = true;
    loadMenuList().then(() => { if (alive) listArrived((count) => count + 1); }, () => {
      if (alive) finish('escape', true);
    });
    return () => { alive = false; };
  }, [finish, MenuList]);

  const registerMenu = useCallback((element: HTMLElement) => {
    menus.current.add(element);
    return () => {
      menus.current.delete(element);
    };
  }, []);

  useLayoutEffect(() => {
    const inside = (node: unknown) =>
      node instanceof Node && [...menus.current].some((menu) => menu.contains(node));
    const onPointerDown = (event: PointerEvent) => {
      if (!inside(event.target)) finish('outside', false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !inside(event.target)) finish('escape', false);
    };
    const onResize = () => finish('resize', inside(document.activeElement));
    document.addEventListener('pointerdown', onPointerDown, true);
    const releaseGuard = guardPendingMenu((key) => finish(key === 'Tab' ? 'tab' : 'escape', false));
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', onResize);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      releaseGuard();
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', onResize);
    };
  }, [finish]);

  const anchor = useCallback(
    () =>
      ({
        x: point.x,
        y: point.y,
        left: point.x,
        right: point.x,
        top: point.y,
        bottom: point.y,
        width: 0,
        height: 0,
      }) as DOMRect,
    [point.x, point.y]
  );

  if (!MenuList || closed.current) return null;
  return (
    <MenuList
      label={label}
      entries={entries}
      anchor={anchor}
      mode="point"
      colors={{ ...DEFAULT_MENU_COLORS, ...colors }}
      initialFocus={modality === 'keyboard' ? 'first' : 'menu'}
      onClose={(reason) => finish(reason === 'tab' ? 'tab' : 'escape', reason !== 'tab')}
      onSelect={(entry) => {
        finish('select', true);
        entry.onSelect();
      }}
      registerMenu={registerMenu}
      className={className}
      testId={testId}
    />
  );
}

/** What a context menu request targets, in viewport coordinates. */
export interface ContextMenuRequest {
  x: number;
  y: number;
  modality: ContextMenuModality;
  target: EventTarget | null;
}

/** The menu a host shows for one request; `entries` is read again whenever `subscribe` fires. */
export interface ContextMenuSpec {
  label: string;
  entries(): readonly OverflowMenuEntry[];
  className?: string;
  testId?: string;
  /** Where Escape and choosing an item send focus; defaults to the element focused before the menu opened. */
  returnFocus?(): void;
}

export interface UseContextMenuOptions {
  /** Element whose context menu events open the menu. */
  surface: RefObject<HTMLElement | null>;
  /** The menu for a request, or null to leave the event to the browser. */
  resolve(request: ContextMenuRequest): ContextMenuSpec | null;
  /** Where a keyboard-opened menu appears, such as the selection's corner. */
  keyboardPoint?(): { x: number; y: number } | null;
  /** Re-reads the open menu's entries when the host's state changes. */
  subscribe?(listener: () => void): () => void;
  colors?: Partial<OverflowMenuColors>;
}

export interface ContextMenuController {
  element: ReactNode;
  isOpen: boolean;
  /** Opens the menu for a request; false when the host has no menu for it. */
  open(request: ContextMenuRequest): boolean;
  close(): void;
}

const POINTER_WINDOW_MS = 1000;
const KEYBOARD_DEDUPE_MS = 500;

/** Whether a keydown asks for the context menu: the Menu key or Shift+F10. */
export function isContextMenuKey(
  event: Pick<KeyboardEvent, 'key' | 'shiftKey' | 'ctrlKey' | 'metaKey' | 'altKey'>
): boolean {
  if (event.key === 'ContextMenu') return true;
  return event.key === 'F10' && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey;
}

/**
 * Opens a {@link ContextMenu} for right-clicks, the Menu key and Shift+F10 on
 * `surface`. Events another handler already prevented are left alone, so a
 * nested surface can show its own menu.
 */
export function useContextMenu(options: UseContextMenuOptions): ContextMenuController {
  const optionsRef = useRef(options);
  optionsRef.current = options;
  const [state, setState] = useState<{
    spec: ContextMenuSpec;
    x: number;
    y: number;
    modality: ContextMenuModality;
    key: number;
  } | null>(null);
  const [, refresh] = useReducer((count: number) => count + 1, 0);
  const sequence = useRef(0);
  const lastPointer = useRef(-Infinity);
  const lastKeyboard = useRef(-Infinity);

  useEffect(() => preloadMenuList(), []);

  const open = useCallback((request: ContextMenuRequest) => {
    const spec = optionsRef.current.resolve(request);
    if (!spec) return false;
    setState({ spec, x: request.x, y: request.y, modality: request.modality, key: ++sequence.current });
    return true;
  }, []);
  const close = useCallback(() => setState(null), []);

  const keyboardRequest = useCallback((target: EventTarget | null): ContextMenuRequest => {
    const point = optionsRef.current.keyboardPoint?.();
    const box = optionsRef.current.surface.current?.getBoundingClientRect();
    return {
      x: point?.x ?? (box ? box.left + 8 : 8),
      y: point?.y ?? (box ? box.top + 8 : 8),
      modality: 'keyboard',
      target,
    };
  }, []);

  useEffect(() => {
    const surface = options.surface.current;
    if (!surface) return;
    const onPointerDown = (event: PointerEvent) => {
      if (event.button === 2 || (event.button === 0 && event.ctrlKey)) lastPointer.current = performance.now();
    };
    const onContextMenu = (event: MouseEvent) => {
      if (event.defaultPrevented) return;
      const now = performance.now();
      const pointer = event.button === 2 || now - lastPointer.current < POINTER_WINDOW_MS;
      lastPointer.current = -Infinity;
      if (!pointer && now - lastKeyboard.current < KEYBOARD_DEDUPE_MS) {
        event.preventDefault();
        return;
      }
      const request = pointer
        ? { x: event.clientX, y: event.clientY, modality: 'pointer' as const, target: event.target }
        : keyboardRequest(event.target);
      if (open(request)) event.preventDefault();
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.isComposing || !isContextMenuKey(event)) return;
      if (open(keyboardRequest(event.target))) {
        event.preventDefault();
        lastKeyboard.current = performance.now();
      }
    };
    surface.addEventListener('pointerdown', onPointerDown, true);
    surface.addEventListener('contextmenu', onContextMenu);
    surface.addEventListener('keydown', onKeyDown);
    return () => {
      surface.removeEventListener('pointerdown', onPointerDown, true);
      surface.removeEventListener('contextmenu', onContextMenu);
      surface.removeEventListener('keydown', onKeyDown);
    };
  });

  const subscribe = options.subscribe;
  const isOpen = state !== null;
  useEffect(() => (isOpen && subscribe ? subscribe(refresh) : undefined), [isOpen, subscribe]);

  const element = state ? (
    <ContextMenu
      key={state.key}
      label={state.spec.label}
      entries={state.spec.entries()}
      point={{ x: state.x, y: state.y }}
      modality={state.modality}
      colors={options.colors}
      className={state.spec.className}
      testId={state.spec.testId}
      returnFocus={state.spec.returnFocus}
      onClose={close}
    />
  ) : null;

  return { element, isOpen, open, close };
}

/** Joins non-empty sections with separators, so hidden items never leave a stray divider. */
export function menuSections(
  sections: readonly (readonly OverflowMenuEntry[])[]
): OverflowMenuEntry[] {
  const out: OverflowMenuEntry[] = [];
  sections.forEach((section, index) => {
    if (section.length === 0) return;
    if (out.length > 0) out.push({ kind: 'separator', id: `separator-${index}` });
    out.push(...section);
  });
  return out;
}

/** The availability of a command as a menu shows it. */
export interface MenuCommandState {
  enabled: boolean;
  disabledReason?: { message: string };
}

/** A menu item for one command: disabled with its reason, labelled with its shortcut. */
export function commandMenuItem(options: {
  id: string;
  label: string;
  state: MenuCommandState;
  shortcut?: string | null;
  icon?: ReactNode;
  run(): void;
}): OverflowMenuEntry {
  const { id, label, state, shortcut, icon, run } = options;
  return {
    kind: 'item',
    id,
    label,
    icon,
    shortcut: shortcut ?? undefined,
    disabled: !state.enabled,
    description: state.enabled ? undefined : state.disabledReason?.message,
    onSelect: run,
  };
}
