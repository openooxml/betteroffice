/**
 * Hook for toolbar dropdowns that need position:fixed to escape overflow:auto/hidden ancestors.
 *
 * Returns refs and styles for a dropdown that positions itself below its trigger
 * using fixed coordinates (like MenuDropdown), so it isn't clipped by the toolbar's
 * overflow-x-auto container.
 */

import { useState, useRef, useEffect, useCallback } from 'react';
import type { CSSProperties, RefObject } from 'react';

export interface UseFixedDropdownOptions {
  isOpen: boolean;
  onClose: () => void;
  /** 'left' aligns dropdown left edge to trigger, 'right' aligns right edge */
  align?: 'left' | 'right';
}

export interface UseFixedDropdownReturn {
  containerRef: RefObject<HTMLDivElement | null>;
  dropdownRef: RefObject<HTMLDivElement | null>;
  dropdownStyle: CSSProperties;
  handleMouseDown: (e: React.MouseEvent) => void;
}

export function useFixedDropdown({
  isOpen,
  onClose,
  align = 'left',
}: UseFixedDropdownOptions): UseFixedDropdownReturn {
  const containerRef = useRef<HTMLDivElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });

  const updatePosition = useCallback(() => {
    if (!containerRef.current) return;
    const rect = containerRef.current.getBoundingClientRect();
    const width = dropdownRef.current?.getBoundingClientRect().width ?? 0;
    setPos({
      top: rect.bottom + 4,
      left: align === 'right' && width > 0 ? rect.right - width : rect.left,
    });
  }, [align]);

  // Calculate position when opening
  useEffect(() => {
    if (!isOpen) return;
    // The menu must render before right alignment can measure its width.
    const frame = requestAnimationFrame(updatePosition);
    return () => cancelAnimationFrame(frame);
  }, [isOpen, updatePosition]);

  // Dismiss on outside click / Escape; stay anchored across scroll.
  useEffect(() => {
    if (!isOpen) return;

    const handleClickOutside = (e: MouseEvent) => {
      const target = e.target as Node;
      if (
        containerRef.current &&
        !containerRef.current.contains(target) &&
        dropdownRef.current &&
        !dropdownRef.current.contains(target)
      ) {
        onClose();
      }
    };

    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };

    const handleScroll = (e: Event) => {
      // Ignore scrolls inside the dropdown's own scrollable list (e.g. the font
      // size presets). Canvas layout, caret reveal, and the toolbar's own
      // horizontal overflow can all scroll while a menu is open. Re-anchor
      // the fixed dropdown instead of closing it; outside-click / Escape still
      // own dismissal.
      const target = e.target as Node | null;
      if (target && dropdownRef.current && dropdownRef.current.contains(target)) {
        return;
      }
      requestAnimationFrame(updatePosition);
    };

    document.addEventListener('mousedown', handleClickOutside);
    document.addEventListener('keydown', handleEscape);
    window.addEventListener('scroll', handleScroll, true);

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleEscape);
      window.removeEventListener('scroll', handleScroll, true);
    };
  }, [isOpen, onClose, updatePosition]);

  const handleMouseDown = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const dropdownStyle: CSSProperties = {
    position: 'fixed',
    top: pos.top,
    left: pos.left,
    zIndex: 10000,
  };

  return { containerRef, dropdownRef, dropdownStyle, handleMouseDown };
}
