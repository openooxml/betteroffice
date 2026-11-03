import { forwardRef, useLayoutEffect, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent } from 'react';
import { cn } from '../lib/utils';
import { MaterialSymbol } from './ui/MaterialSymbol';
import { ToolbarButton } from './toolbar/ToolbarPrimitives';

/** A client-space rectangle, as returned by `getBoundingClientRect`. */
export interface SelectionToolbarAnchorRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
  width: number;
  height: number;
}

/** One resolved button of the selection mini toolbar. */
export interface SelectionMiniToolbarButton {
  /** Stable item id, also used as the React key. */
  id: string;
  label: string;
  shortcut?: string | null;
  icon: string;
  active?: boolean;
  onActivate: () => void;
}

export interface SelectionMiniToolbarProps {
  anchorRect: SelectionToolbarAnchorRect;
  buttons: readonly SelectionMiniToolbarButton[];
  label: string;
  hint?: string;
  /** Forces the coarse-pointer minimum target size; defaults to the `(pointer: coarse)` media query. */
  coarsePointer?: boolean;
  onDismiss: (returnFocus: boolean) => void;
}

const VIEWPORT_MARGIN = 8;
const ANCHOR_GAP = 8;

/** Position of the toolbar for an anchor, flipped and clamped to the viewport. */
export function computeSelectionToolbarPosition(
  anchor: SelectionToolbarAnchorRect,
  size: { width: number; height: number },
  viewport: { width: number; height: number }
): { left: number; top: number; flipped: boolean } {
  const maxLeft = Math.max(VIEWPORT_MARGIN, viewport.width - size.width - VIEWPORT_MARGIN);
  const center = anchor.left + anchor.width / 2 - size.width / 2;
  const left = Math.min(Math.max(center, VIEWPORT_MARGIN), maxLeft);
  const above = anchor.top - size.height - ANCHOR_GAP;
  if (above >= VIEWPORT_MARGIN) return { left, top: above, flipped: false };
  const below = anchor.bottom + ANCHOR_GAP;
  const maxTop = Math.max(VIEWPORT_MARGIN, viewport.height - size.height - VIEWPORT_MARGIN);
  return { left, top: Math.min(below, maxTop), flipped: true };
}

/** Whether the primary pointer is coarse, when the browser can say. */
export function isCoarsePointer(): boolean {
  return typeof window !== 'undefined' && typeof window.matchMedia === 'function'
    ? window.matchMedia('(pointer: coarse)').matches
    : false;
}

const barStyle: CSSProperties = {
  position: 'fixed',
  zIndex: 60,
  display: 'flex',
  alignItems: 'center',
  gap: 2,
  padding: 4,
  borderRadius: 8,
  border: '1px solid var(--doc-border)',
  backgroundColor: 'var(--doc-surface)',
  boxShadow: '0 4px 16px var(--doc-shadow)',
};

/**
 * Floating toolbar shown above a settled text selection. Presentational:
 * the caller resolves labels, shortcuts and commands, and owns visibility.
 */
export const SelectionMiniToolbar = forwardRef<HTMLDivElement, SelectionMiniToolbarProps>(
  function SelectionMiniToolbar({ anchorRect, buttons, label, hint, coarsePointer, onDismiss }, ref) {
    const rootRef = useRef<HTMLDivElement | null>(null);
    const buttonRefs = useRef<Array<HTMLButtonElement | null>>([]);
    const [activeIndex, setActiveIndex] = useState(0);
    const [size, setSize] = useState<{ width: number; height: number } | null>(null);
    const coarse = coarsePointer ?? isCoarsePointer();

    const setRoot = (element: HTMLDivElement | null) => {
      rootRef.current = element;
      if (typeof ref === 'function') ref(element);
      else if (ref) (ref as { current: HTMLDivElement | null }).current = element;
    };

    useLayoutEffect(() => {
      const element = rootRef.current;
      if (!element) return;
      const measure = () => {
        const rect = element.getBoundingClientRect();
        setSize((current) =>
          current && current.width === rect.width && current.height === rect.height
            ? current
            : { width: rect.width, height: rect.height }
        );
      };
      measure();
      if (typeof ResizeObserver === 'undefined') return;
      const observer = new ResizeObserver(measure);
      observer.observe(element);
      return () => observer.disconnect();
    }, [buttons.length, coarse]);

    const focusButton = (index: number) => {
      const clamped = Math.min(Math.max(index, 0), buttons.length - 1);
      setActiveIndex(clamped);
      buttonRefs.current[clamped]?.focus();
    };

    const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
      switch (event.key) {
        case 'ArrowRight':
          event.preventDefault();
          focusButton(activeIndex + 1 >= buttons.length ? 0 : activeIndex + 1);
          break;
        case 'ArrowLeft':
          event.preventDefault();
          focusButton(activeIndex - 1 < 0 ? buttons.length - 1 : activeIndex - 1);
          break;
        case 'Home':
          event.preventDefault();
          focusButton(0);
          break;
        case 'End':
          event.preventDefault();
          focusButton(buttons.length - 1);
          break;
        case 'Escape':
          event.preventDefault();
          event.stopPropagation();
          onDismiss(true);
          break;
        default:
          break;
      }
    };

    if (buttons.length === 0) return null;
    const viewport =
      typeof window !== 'undefined'
        ? { width: window.innerWidth, height: window.innerHeight }
        : { width: 1024, height: 768 };
    const position = size
      ? computeSelectionToolbarPosition(anchorRect, size, viewport)
      : { left: VIEWPORT_MARGIN, top: VIEWPORT_MARGIN, flipped: false };

    return (
      <div
        ref={setRoot}
        role="toolbar"
        className="docx-selection-toolbar"
        aria-label={label}
        aria-describedby={hint ? 'docx-selection-toolbar-hint' : undefined}
        data-testid="selection-mini-toolbar"
        data-flipped={position.flipped ? 'true' : undefined}
        onKeyDown={onKeyDown}
        style={{
          ...barStyle,
          left: position.left,
          top: position.top,
          visibility: size ? 'visible' : 'hidden',
        }}
      >
        {hint ? (
          <span id="docx-selection-toolbar-hint" hidden>
            {hint}
          </span>
        ) : null}
        {buttons.map((button, index) => (
          <span
            key={button.id}
            ref={(element) => {
              const inner = element?.querySelector('button');
              buttonRefs.current[index] = inner instanceof HTMLButtonElement ? inner : null;
            }}
            className="inline-flex"
            onFocus={() => setActiveIndex(index)}
          >
            <ToolbarButton
              active={button.active}
              title={button.label}
              ariaLabel={button.label}
              tabIndex={index === activeIndex ? 0 : -1}
              shortcut={button.shortcut ?? undefined}
              className={cn(coarse && 'h-11 w-11 min-h-[44px] min-w-[44px]')}
              onClick={() => button.onActivate()}
            >
              <MaterialSymbol name={button.icon} size={coarse ? 20 : 18} />
            </ToolbarButton>
          </span>
        ))}
      </div>
    );
  }
);
