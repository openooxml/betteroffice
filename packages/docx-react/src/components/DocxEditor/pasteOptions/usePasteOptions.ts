import { useCallback, useEffect, useRef, useState } from 'react';
import type { PasteFormattingOption } from '@betteroffice/docx/yrs';

/** @internal Client coordinates of the end of the pasted range. */
export interface PasteOptionsAnchor {
  left: number;
  top: number;
}

/** @internal State and event handling for the post-paste formatting control. */
export function usePasteOptions(options: {
  enabled: boolean;
  onChoose(option: PasteFormattingOption): void;
  focusEditor(): void;
}) {
  const { enabled } = options;
  const callbacks = useRef(options);
  callbacks.current = options;
  const [visible, setVisible] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [selected, setSelected] = useState<PasteFormattingOption>('keep-source');
  const [anchor, setAnchor] = useState<PasteOptionsAnchor | null>(null);
  const ctrlAlone = useRef(false);

  const dismiss = useCallback(() => {
    ctrlAlone.current = false;
    setMenuOpen(false);
    setVisible(false);
    setAnchor(null);
  }, []);

  const show = useCallback((nextAnchor: PasteOptionsAnchor) => {
    if (!callbacks.current.enabled) return;
    ctrlAlone.current = false;
    setAnchor(nextAnchor);
    setSelected('keep-source');
    setMenuOpen(false);
    setVisible(true);
  }, []);

  const choose = useCallback((option: PasteFormattingOption) => {
    setSelected(option);
    setMenuOpen(false);
    callbacks.current.onChoose(option);
    callbacks.current.focusEditor();
  }, []);

  const closeMenu = useCallback(() => {
    setMenuOpen(false);
    callbacks.current.focusEditor();
  }, []);

  const handleKeyDown = useCallback((event: { key: string }): boolean => {
    if (event.key === 'Control') {
      ctrlAlone.current = true;
      return false;
    }
    if (ctrlAlone.current) ctrlAlone.current = false;
    return false;
  }, []);

  const handleKeyUp = useCallback((event: { key: string }): boolean => {
    if (event.key !== 'Control') return false;
    const alone = ctrlAlone.current;
    ctrlAlone.current = false;
    if (alone && callbacks.current.enabled && visible) {
      setMenuOpen(true);
      return true;
    }
    return false;
  }, [visible]);

  useEffect(() => {
    if (!visible) return;
    const hide = () => dismiss();
    const pointer = (event: PointerEvent) => {
      const target = event.target as Element | null;
      if (target?.closest?.('[data-paste-options]')) return;
      dismiss();
    };
    window.addEventListener('scroll', hide, true);
    window.addEventListener('resize', hide);
    document.addEventListener('pointerdown', pointer, true);
    return () => {
      window.removeEventListener('scroll', hide, true);
      window.removeEventListener('resize', hide);
      document.removeEventListener('pointerdown', pointer, true);
    };
  }, [visible, dismiss]);

  useEffect(() => {
    if (!enabled) dismiss();
  }, [enabled, dismiss]);

  return {
    visible: enabled && visible,
    menuOpen,
    selected,
    anchor,
    show,
    dismiss,
    setAnchor,
    setMenuOpen,
    choose,
    closeMenu,
    handleKeyDown,
    handleKeyUp,
  };
}

/** @internal The control's controller shape, shared by both editor inputs. */
export type PasteOptionsController = ReturnType<typeof usePasteOptions>;
