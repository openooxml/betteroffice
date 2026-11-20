import { useCallback, useEffect, useRef, useState, type RefObject } from 'react';
import type { PasteFormattingOption } from './clipboardFragmentTransforms';

/** @internal Client coordinates of the end of the pasted range. */
export interface PasteOptionsAnchor {
  left: number;
  top: number;
}

/** @internal State and event handling for the post-paste formatting control. */
export function usePasteOptions(options: {
  enabled: boolean;
  /** Returns whether the choice was accepted; only an accepted choice becomes the checked option. */
  onChoose(option: PasteFormattingOption): boolean;
  focusEditor(): void;
  /** The editor input; a Ctrl tap opens the menu only from inside it or the control. */
  surface: RefObject<HTMLElement | null>;
  /** Called when an outside pointer or a window resize hides the control, and each time the control is turned off. */
  onHide?(): void;
}) {
  const { enabled } = options;
  const callbacks = useRef(options);
  callbacks.current = options;
  const controlRef = useRef<HTMLDivElement | null>(null);
  const [visible, setVisible] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [selected, setSelected] = useState<PasteFormattingOption>('keep-source');
  const [offered, setOffered] = useState<readonly PasteFormattingOption[]>([]);
  const [anchor, setAnchor] = useState<PasteOptionsAnchor | null>(null);
  const ctrlAlone = useRef(false);

  const dismiss = useCallback(() => {
    ctrlAlone.current = false;
    setMenuOpen(false);
    setVisible(false);
    setAnchor(null);
  }, []);

  const show = useCallback((nextAnchor: PasteOptionsAnchor, nextOffered: readonly PasteFormattingOption[]) => {
    if (!callbacks.current.enabled) return;
    ctrlAlone.current = false;
    setAnchor(nextAnchor);
    setOffered(nextOffered);
    setSelected('keep-source');
    setMenuOpen(false);
    setVisible(true);
  }, []);

  const choose = useCallback((option: PasteFormattingOption) => {
    if (callbacks.current.onChoose(option)) setSelected(option);
    setMenuOpen(false);
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
    const owned = (target: EventTarget | null) => target instanceof Node &&
      (callbacks.current.surface.current?.contains(target) === true || controlRef.current?.contains(target) === true);
    const keydown = (event: KeyboardEvent) => {
      ctrlAlone.current = event.key === 'Control' && owned(event.target);
    };
    const keyup = (event: KeyboardEvent) => {
      if (event.key !== 'Control') return;
      const alone = ctrlAlone.current;
      ctrlAlone.current = false;
      if (alone && callbacks.current.enabled && owned(event.target)) setMenuOpen(true);
    };
    document.addEventListener('keydown', keydown, true);
    document.addEventListener('keyup', keyup, true);
    return () => {
      document.removeEventListener('keydown', keydown, true);
      document.removeEventListener('keyup', keyup, true);
    };
  }, [visible]);

  useEffect(() => {
    if (!visible) return;
    const hide = () => {
      dismiss();
      callbacks.current.onHide?.();
    };
    const pointer = (event: PointerEvent) => {
      const target = event.target;
      if (target instanceof Node && controlRef.current?.contains(target)) return;
      hide();
    };
    window.addEventListener('resize', hide);
    document.addEventListener('pointerdown', pointer, true);
    return () => {
      window.removeEventListener('resize', hide);
      document.removeEventListener('pointerdown', pointer, true);
    };
  }, [visible, dismiss]);

  useEffect(() => {
    if (enabled) return;
    dismiss();
    callbacks.current.onHide?.();
  }, [enabled, dismiss]);

  return {
    controlRef,
    visible: enabled && visible,
    menuOpen,
    selected,
    offered,
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
