import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, RefObject } from 'react';
import { RibbonIcon } from './RibbonIcon';
import type { RibbonIconName } from './RibbonIcon';
import { useRibbonCommands } from './commands';
import type { RibbonCommandId } from './commands';

/** Single entry in a keyboard-navigable command menu, optionally opening a submenu. */
export interface CommandMenuEntry {
  id: RibbonCommandId;
  icon: RibbonIconName;
  shortcut?: string;
  children?: ReadonlyArray<CommandMenuEntry>;
}

interface CommandMenuProps {
  menuLabel: string;
  entries: ReadonlyArray<CommandMenuEntry>;
  position: { top: number; left: number };
  dividerAfter?: ReadonlySet<RibbonCommandId>;
  anchorRef?: RefObject<HTMLElement | null>;
  initialFocus?: 'first' | 'last';
  label: (id: RibbonCommandId) => string;
  onClose: () => void;
  onCloseAndFocus: () => void;
}

/** One menu item bound to a ribbon command id. */
export function CommandMenuItem({ id, icon, label, shortcut, itemRef, onSelect }: { id: RibbonCommandId; icon: RibbonIconName; label: string; shortcut?: string; itemRef: (node: HTMLButtonElement | null) => void; onSelect: () => void }) {
  const command = useRibbonCommands()[id];
  const checkable = command.active !== undefined;
  return <button ref={itemRef} type="button" role={checkable ? 'menuitemcheckbox' : 'menuitem'} aria-checked={checkable ? command.active : undefined} aria-keyshortcuts={shortcut} disabled={!command.enabled} aria-label={shortcut ? `${label} ${shortcut}` : label} data-command-id={id} tabIndex={-1} onMouseDown={(event) => event.preventDefault()} onClick={() => { command.run(); onSelect(); }} onMouseOver={(event) => { if (command.enabled) event.currentTarget.style.backgroundColor = '#f5f5f5'; }} onMouseOut={(event) => { event.currentTarget.style.backgroundColor = 'transparent'; }} style={{ ...styles.menuItem, color: command.enabled ? '#242424' : '#b4b4b4' }}><RibbonIcon name={icon} size={18} /><span>{label}</span>{shortcut && <span aria-hidden="true" style={styles.shortcut}>{shortcut}</span>}</button>;
}

/** Shared keyboard-navigable command menu behind the ribbon split buttons and the canvas context menu. */
export function CommandMenu({ menuLabel, entries, position, dividerAfter, anchorRef, initialFocus = 'first', label, onClose, onCloseAndFocus }: CommandMenuProps) {
  const commands = useRibbonCommands();
  const menuRef = useRef<HTMLDivElement>(null);
  const submenuRef = useRef<HTMLDivElement>(null);
  const [submenuPosition, setSubmenuPosition] = useState({ top: 0, left: 0 });
  const itemRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const subItemRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const [pos, setPos] = useState(position);
  const [openSubmenu, setOpenSubmenu] = useState<RibbonCommandId | null>(null);
  const focusItem = useCallback((index: number) => { itemRefs.current[index]?.focus(); }, []);
  const isSubmenu = useCallback((entry: CommandMenuEntry) => (entry.children?.length ?? 0) > 0, []);
  const entryEnabled = useCallback((entry: CommandMenuEntry) => (isSubmenu(entry) ? (entry.children ?? []).some((child) => commands[child.id].enabled) : commands[entry.id].enabled), [commands, isSubmenu]);
  const openEntry = entries.find((entry) => entry.id === openSubmenu && isSubmenu(entry)) ?? null;
  const openChildren = openEntry?.children ?? [];
  const enabledIndices = entries.map((entry, index) => (entryEnabled(entry) ? index : -1)).filter((index) => index >= 0);
  const firstEnabled = enabledIndices[0] ?? -1;
  const lastEnabled = enabledIndices[enabledIndices.length - 1] ?? -1;
  const checkedEnabled = enabledIndices.find((index) => commands[entries[index].id].active === true) ?? -1;
  const step = useCallback((from: number, delta: 1 | -1) => {
    if (enabledIndices.length === 0) return -1;
    const at = enabledIndices.indexOf(from);
    if (at === -1) return delta === 1 ? (enabledIndices[0] ?? -1) : (enabledIndices[enabledIndices.length - 1] ?? -1);
    return enabledIndices[(at + delta + enabledIndices.length) % enabledIndices.length] ?? -1;
  }, [enabledIndices]);
  const openParentIndex = openEntry ? entries.findIndex((entry) => entry.id === openEntry.id) : -1;
  const closeSubmenu = useCallback((focusParent: boolean) => {
    setOpenSubmenu(null);
    if (focusParent && openParentIndex >= 0) itemRefs.current[openParentIndex]?.focus();
  }, [openParentIndex]);
  const openSubmenuAt = useCallback((index: number) => {
    const entry = entries[index];
    if (!entry || !isSubmenu(entry) || !entryEnabled(entry)) return;
    subItemRefs.current = [];
    setOpenSubmenu(entry.id);
  }, [entries, entryEnabled, isSubmenu]);
  useEffect(() => {
    const target = initialFocus === 'last' ? lastEnabled : (checkedEnabled >= 0 ? checkedEnabled : firstEnabled);
    if (target >= 0) focusItem(target);
  }, [initialFocus, firstEnabled, lastEnabled, checkedEnabled, focusItem]);
  useEffect(() => {
    if (!openEntry) return;
    const target = openChildren.findIndex((child) => commands[child.id].enabled);
    if (target >= 0) subItemRefs.current[target]?.focus();
  }, [openEntry, openChildren, commands]);
  useEffect(() => {
    if (!openEntry || openChildren.some((child) => commands[child.id].enabled)) return;
    setOpenSubmenu(null);
  }, [openEntry, openChildren, commands]);
  useEffect(() => {
    const node = menuRef.current;
    if (!node) return;
    const rect = node.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return;
    setPos((previous) => {
      const next = { top: Math.max(MENU_MARGIN, Math.min(position.top, window.innerHeight - rect.height - MENU_MARGIN)), left: Math.max(MENU_MARGIN, Math.min(position.left, window.innerWidth - rect.width - MENU_MARGIN)) };
      return previous.top === next.top && previous.left === next.left ? previous : next;
    });
  }, [position]);
  useLayoutEffect(() => {
    const submenu = submenuRef.current;
    const anchor = itemRefs.current[openParentIndex];
    if (!submenu || !anchor) return;
    const bounds = submenu.getBoundingClientRect();
    const parent = anchor.getBoundingClientRect();
    const left = parent.right + bounds.width <= window.innerWidth - MENU_MARGIN ? parent.right : parent.left - bounds.width;
    setSubmenuPosition({
      left: Math.max(MENU_MARGIN, Math.min(left, window.innerWidth - bounds.width - MENU_MARGIN)),
      top: Math.max(MENU_MARGIN, Math.min(parent.top, window.innerHeight - bounds.height - MENU_MARGIN)),
    });
  }, [openParentIndex, pos]);
  useEffect(() => {
    function onOutside(event: MouseEvent) {
      const target = event.target as Node;
      if (anchorRef?.current?.contains(target)) return;
      if (menuRef.current && !menuRef.current.contains(target)) onClose();
    }
    function onEscape(event: globalThis.KeyboardEvent) { if (!event.defaultPrevented && event.key === 'Escape') onCloseAndFocus(); }
    function onScroll() { onClose(); }
    document.addEventListener('mousedown', onOutside);
    document.addEventListener('keydown', onEscape);
    window.addEventListener('scroll', onScroll, true);
    return () => {
      document.removeEventListener('mousedown', onOutside);
      document.removeEventListener('keydown', onEscape);
      window.removeEventListener('scroll', onScroll, true);
    };
  }, [anchorRef, onClose, onCloseAndFocus]);
  function onSubmenuKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    const active = document.activeElement;
    const currentIndex = subItemRefs.current.findIndex((node) => node === active);
    const subEnabled = openChildren.map((child, index) => (commands[child.id].enabled ? index : -1)).filter((index) => index >= 0);
    const subStep = (from: number, delta: 1 | -1): number => {
      if (subEnabled.length === 0) return -1;
      const at = subEnabled.indexOf(from);
      if (at === -1) return delta === 1 ? (subEnabled[0] ?? -1) : (subEnabled[subEnabled.length - 1] ?? -1);
      return subEnabled[(at + delta + subEnabled.length) % subEnabled.length] ?? -1;
    };
    if (event.key === 'Escape' || event.key === 'ArrowLeft') { event.preventDefault(); event.stopPropagation(); closeSubmenu(true); return; }
    if (event.key === 'Tab') { event.stopPropagation(); onClose(); return; }
    if (event.key === 'ArrowDown') { event.preventDefault(); event.stopPropagation(); const next = subStep(currentIndex, 1); if (next >= 0) subItemRefs.current[next]?.focus(); }
    else if (event.key === 'ArrowUp') { event.preventDefault(); event.stopPropagation(); const next = subStep(currentIndex, -1); if (next >= 0) subItemRefs.current[next]?.focus(); }
    else if (event.key === 'Home') { event.preventDefault(); event.stopPropagation(); if (subEnabled[0] !== undefined && subEnabled[0] >= 0) subItemRefs.current[subEnabled[0]]?.focus(); }
    else if (event.key === 'End') { event.preventDefault(); event.stopPropagation(); const last = subEnabled[subEnabled.length - 1]; if (last !== undefined && last >= 0) subItemRefs.current[last]?.focus(); }
  }
  function onMenuKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    const active = document.activeElement;
    if (subItemRefs.current.some((node) => node !== null && node === active)) return;
    if (event.key === 'Escape') {
      if (openSubmenu !== null) { event.preventDefault(); event.stopPropagation(); closeSubmenu(true); return; }
      event.preventDefault(); onCloseAndFocus(); return;
    }
    if (event.key === 'Tab') { onClose(); return; }
    const currentIndex = itemRefs.current.findIndex((node) => node === active);
    if (event.key === 'ArrowDown') { event.preventDefault(); const next = step(currentIndex, 1); if (next >= 0) focusItem(next); }
    else if (event.key === 'ArrowUp') { event.preventDefault(); const next = step(currentIndex, -1); if (next >= 0) focusItem(next); }
    else if (event.key === 'Home') { event.preventDefault(); if (firstEnabled >= 0) focusItem(firstEnabled); }
    else if (event.key === 'End') { event.preventDefault(); if (lastEnabled >= 0) focusItem(lastEnabled); }
    else if (event.key === 'ArrowRight' || event.key === 'Enter' || event.key === ' ') {
      const entry = entries[currentIndex];
      if (entry && isSubmenu(entry) && entryEnabled(entry)) { event.preventDefault(); openSubmenuAt(currentIndex); }
    }
    else if (event.key === 'ArrowLeft') {
      if (openSubmenu !== null) { event.preventDefault(); closeSubmenu(true); }
    }
  }
  return (
    <div ref={menuRef} role="menu" aria-label={menuLabel} onMouseDown={(event) => event.preventDefault()} onKeyDown={onMenuKeyDown} style={{ ...styles.menu, top: pos.top, left: pos.left }}>
      {entries.map((entry, index) => {
        const submenu = isSubmenu(entry);
        const enabled = entryEnabled(entry);
        const open = submenu && openSubmenu === entry.id;
        return (
          <span key={entry.id} style={styles.entryWrap}>
            {submenu ? (
              <button ref={(node) => { itemRefs.current[index] = node; }} type="button" role="menuitem" aria-haspopup="menu" aria-expanded={open} disabled={!enabled} aria-label={label(entry.id)} data-submenu-id={entry.id} tabIndex={-1} onMouseDown={(event) => event.preventDefault()} onClick={() => { if (!enabled) return; if (open) closeSubmenu(true); else openSubmenuAt(index); }} onMouseEnter={() => { if (enabled && !open) openSubmenuAt(index); }} onMouseOver={(event) => { if (enabled) event.currentTarget.style.backgroundColor = '#f5f5f5'; }} onMouseOut={(event) => { event.currentTarget.style.backgroundColor = 'transparent'; }} style={{ ...styles.menuItem, color: enabled ? '#242424' : '#b4b4b4' }}><RibbonIcon name={entry.icon} size={18} /><span>{label(entry.id)}</span><span aria-hidden="true" style={styles.chevron}>▸</span></button>
            ) : (
              <CommandMenuItem id={entry.id} icon={entry.icon} label={label(entry.id)} shortcut={entry.shortcut} itemRef={(node) => { itemRefs.current[index] = node; }} onSelect={onCloseAndFocus} />
            )}
            {open && openEntry && (
              <div ref={submenuRef} role="menu" aria-label={label(openEntry.id)} data-submenu={openEntry.id} onKeyDown={onSubmenuKeyDown} style={{ ...styles.submenu, ...submenuPosition }}>
                {(openEntry.children ?? []).map((child, childIndex) => (
                  <CommandMenuItem key={child.id} id={child.id} icon={child.icon} label={label(child.id)} shortcut={child.shortcut} itemRef={(node) => { subItemRefs.current[childIndex] = node; }} onSelect={onCloseAndFocus} />
                ))}
              </div>
            )}
            {dividerAfter?.has(entry.id) && index < entries.length - 1 && <span role="separator" style={styles.separator} />}
          </span>
        );
      })}
    </div>
  );
}

const MENU_MARGIN = 4;

const styles: Record<string, CSSProperties> = {
  menu: { position: 'fixed', minWidth: 200, padding: '4px 0', backgroundColor: '#ffffff', border: '1px solid #e0e0e0', borderRadius: 4, boxShadow: '0 4px 12px rgba(0,0,0,0.14)', zIndex: 10000 },
  menuItem: { display: 'flex', alignItems: 'center', gap: 8, width: '100%', padding: '6px 12px', border: 0, background: 'transparent', cursor: 'pointer', fontSize: 13, textAlign: 'left', whiteSpace: 'nowrap' },
  entryWrap: { display: 'block' },
  submenu: { position: 'fixed', minWidth: 200, padding: '4px 0', backgroundColor: '#ffffff', border: '1px solid #e0e0e0', borderRadius: 4, boxShadow: '0 4px 12px rgba(0,0,0,0.14)', zIndex: 10001 },
  shortcut: { marginLeft: 'auto', paddingLeft: 24, color: '#616161', fontSize: 12 },
  chevron: { marginLeft: 'auto', paddingLeft: 24, color: 'inherit', fontSize: 12 },
  separator: { display: 'block', height: 1, margin: '4px 0', background: '#e0e0e0' },
};
