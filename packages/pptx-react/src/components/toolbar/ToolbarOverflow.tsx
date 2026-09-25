import {
  useCallback,
  useEffect,
  useId,
  useLayoutEffect,
  useMemo,
  useReducer,
  useRef,
  useState,
} from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { createPortal } from 'react-dom';
import {
  OverflowMenu,
  type OverflowMenuEntry,
  type OverflowMenuHandle,
} from '../../../../../shared/react-toolbar/OverflowMenu';
import { useToolbarOverflow } from '../../../../../shared/react-toolbar/useToolbarOverflow';
import { pptxCommandController } from '../../commands/createPptxCommandStore';
import { useCommandChromeRef, usePptxCommands } from '../../commands/hooks';
import { useTranslation } from '../../i18n';
import { ToolbarIcon } from '../ui/ToolbarIcon';
import { interactiveButtonStyle, toolbarColors } from '../ui/ToolbarPrimitives';
import {
  ToolbarOverflowContext,
  useOverflowRegistry,
  useOverflowSource,
  type OverflowPrompt,
  type OverflowRegistry,
  type OverflowSource,
} from './overflowRegistry';
import { ToolbarPromptDialog } from './ToolbarPromptDialog';

const MENU_COLORS = {
  surface: toolbarColors.surface,
  text: toolbarColors.text,
  mutedText: toolbarColors.muted,
  border: toolbarColors.border,
  hover: toolbarColors.hover,
  shadow: 'rgba(60, 64, 67, 0.24)',
};

function isInteractive(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement &&
    (target.tagName === 'INPUT' ||
      target.tagName === 'TEXTAREA' ||
      target.tagName === 'SELECT' ||
      target.tagName === 'OPTION')
  );
}

function documentOrder(a: HTMLElement, b: HTMLElement): number {
  if (a === b) return 0;
  return a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
}

export interface ToolbarRailProps {
  children?: ReactNode;
  className?: string;
  style?: CSSProperties;
  testId?: string;
}

/** The toolbar row: host-ordered controls with an accessible "More" overflow menu. */
export function ToolbarRail({ children, className, style, testId }: ToolbarRailProps) {
  const { t } = useTranslation();
  const store = usePptxCommands();
  const itemsRef = useRef<HTMLDivElement>(null);
  const moreRef = useRef<HTMLSpanElement>(null);
  const menuRef = useRef<OverflowMenuHandle>(null);
  const sources = useRef(new Set<OverflowSource>());
  const [sourceVersion, bumpSources] = useReducer((value: number) => value + 1, 0);
  const [, rerender] = useReducer((value: number) => value + 1, 0);
  const [focusMore, setFocusMore] = useState(false);
  const [prompt, setPrompt] = useState<OverflowPrompt | null>(null);
  const overflowing = useRef(false);

  const registry = useMemo<OverflowRegistry>(
    () => ({
      register(source) {
        sources.current.add(source);
        bumpSources();
        return () => {
          sources.current.delete(source);
          bumpSources();
        };
      },
      overflowing: () => overflowing.current,
      changed() {
        if (overflowing.current) rerender();
      },
      prompt: setPrompt,
      anchor: () => moreRef.current?.querySelector('button') ?? null,
    }),
    []
  );

  const sourcesIn = useCallback((unit: HTMLElement) => {
    const found: { element: HTMLElement; source: OverflowSource }[] = [];
    for (const source of sources.current) {
      const element = source.element();
      if (element && unit.contains(element)) found.push({ element, source });
    }
    return found.sort((a, b) => documentOrder(a.element, b.element));
  }, []);

  const { hidden, remeasure } = useToolbarOverflow({
    items: itemsRef,
    more: moreRef,
    canHide: (unit) => unit.getAttribute('role') === 'separator' || sourcesIn(unit).length > 0,
    onFocusHidden: () => setFocusMore(true),
  });

  useLayoutEffect(() => {
    remeasure();
  }, [sourceVersion, remeasure]);

  useLayoutEffect(() => {
    if (!focusMore || hidden.length === 0) return;
    setFocusMore(false);
    menuRef.current?.focus();
  }, [focusMore, hidden]);

  overflowing.current = hidden.length > 0;

  useEffect(() => {
    if (hidden.length === 0) return;
    return store.subscribe(rerender);
  }, [hidden.length, store]);

  const entries: OverflowMenuEntry[] = [];
  hidden.forEach((unit, index) => {
    const unitEntries = sourcesIn(unit).flatMap(({ source }) => source.entries());
    if (unitEntries.length === 0) return;
    const heading = unit.getAttribute('role') === 'group' ? unit.getAttribute('aria-label') : null;
    entries.push({
      kind: 'group',
      id: `unit-${index}`,
      label: heading ?? undefined,
      entries: unitEntries,
    });
  });
  const menuEntries = entries.flatMap((entry, index) =>
    index === 0 ? [entry] : [{ kind: 'separator' as const, id: `separator-${index}` }, entry]
  );

  return (
    <div
      className={className}
      role="toolbar"
      aria-label={t('toolbar.label')}
      data-testid={testId}
      onMouseDown={(event) => {
        if (!isInteractive(event.target)) event.preventDefault();
      }}
      style={{
        display: 'flex',
        alignItems: 'center',
        minWidth: 0,
        minHeight: 36,
        margin: '0 8px 5px',
        padding: '4px 7px',
        borderRadius: 18,
        background: toolbarColors.rail,
        color: toolbarColors.text,
        boxSizing: 'border-box',
        ...style,
      }}
    >
      <ToolbarOverflowContext.Provider value={registry}>
        <div
          ref={itemsRef}
          data-toolbar-items=""
          style={{
            position: 'relative',
            display: 'flex',
            alignItems: 'center',
            flex: '1 1 auto',
            minWidth: 0,
            overflow: 'hidden',
          }}
        >
          {children}
        </div>
      </ToolbarOverflowContext.Provider>
      {hidden.length > 0 && (
        <span ref={moreRef} style={{ display: 'inline-flex', flex: '0 0 auto', marginLeft: 4 }}>
          <OverflowMenu
            ref={menuRef}
            label={t('toolbar.more')}
            trigger={<ToolbarIcon name="more" />}
            entries={menuEntries}
            colors={MENU_COLORS}
            triggerStyle={interactiveButtonStyle(false, false, false)}
            testId="pptx-toolbar-more"
            onItemSelected={(modality) => {
              if (modality === 'pointer') pptxCommandController(store)?.focusEditor();
            }}
          />
        </span>
      )}
      {prompt && (
        <ToolbarPromptDialog
          prompt={prompt}
          onClose={() => {
            setPrompt(null);
            menuRef.current?.focus();
          }}
        />
      )}
    </div>
  );
}

export interface ToolbarOverflowProps {
  /** Label of the menu entry that replaces the content at narrow widths. */
  label: string;
  onSelect(): void;
  disabled?: boolean;
  /** Why the action is disabled. */
  description?: string;
  /** Toggle state for a checkbox entry. */
  checked?: boolean;
  children: ReactNode;
}

/**
 * Gives arbitrary host content a menu entry, so it can move into the overflow
 * menu instead of staying clipped in a narrow toolbar.
 */
export function ToolbarOverflow({
  label,
  onSelect,
  disabled,
  description,
  checked,
  children,
}: ToolbarOverflowProps) {
  const elementRef = useRef<HTMLSpanElement>(null);
  const id = useId();
  useOverflowSource(elementRef, () => [
    {
      kind: 'item',
      id,
      label,
      checked,
      disabled,
      description: disabled ? description : undefined,
      onSelect,
    },
  ]);
  return (
    <span
      ref={elementRef}
      style={{ display: 'inline-flex', alignItems: 'center', flex: '0 0 auto' }}
    >
      {children}
    </span>
  );
}

/**
 * Host content appended to the legacy toolbar. At narrow widths it moves out of
 * the row, and its menu entry opens it in a popup instead.
 */
export function LegacyToolbarContent({ children }: { children: ReactNode }) {
  const { t } = useTranslation();
  const registry = useOverflowRegistry();
  const unitRef = useRef<HTMLSpanElement>(null);
  const popupRef = useRef<HTMLDivElement>(null);
  const chromeRef = useCommandChromeRef();
  const setPopupRef = useCallback(
    (element: HTMLDivElement | null) => {
      popupRef.current = element;
      chromeRef(element);
    },
    [chromeRef]
  );
  const id = useId();
  const label = t('commands.moreControls');
  const [popup, setPopup] = useState<{ anchor: HTMLElement | null; width: number } | null>(null);
  const [position, setPosition] = useState<{ top: number; left: number } | null>(null);
  useOverflowSource(unitRef, () => [
    {
      kind: 'item',
      id,
      label,
      onSelect: () =>
        setPopup({
          anchor: registry?.anchor() ?? null,
          width: unitRef.current?.getBoundingClientRect().width ?? 0,
        }),
    },
  ]);

  const close = useCallback(
    (restore: boolean) => {
      const anchor = popup?.anchor;
      setPopup(null);
      if (restore) anchor?.focus({ preventScroll: true });
    },
    [popup]
  );

  useLayoutEffect(() => {
    const element = popupRef.current;
    if (!popup || !element) {
      setPosition(null);
      return;
    }
    const anchor = popup.anchor?.getBoundingClientRect();
    const width = element.offsetWidth;
    const left = Math.max(8, Math.min((anchor?.right ?? 8) - width, window.innerWidth - width - 8));
    setPosition({ top: (anchor?.bottom ?? 8) + 4, left });
    element
      .querySelector<HTMLElement>('button, input, select, textarea, a[href], [tabindex]')
      ?.focus();
  }, [popup]);

  useEffect(() => {
    if (!popup) return;
    const onPointerDown = (event: MouseEvent) => {
      if (!popupRef.current?.contains(event.target as Node)) close(false);
    };
    const onResize = () => close(false);
    document.addEventListener('mousedown', onPointerDown);
    window.addEventListener('resize', onResize);
    return () => {
      document.removeEventListener('mousedown', onPointerDown);
      window.removeEventListener('resize', onResize);
    };
  }, [popup, close]);

  return (
    <ToolbarOverflowContext.Provider value={null}>
      <span
        ref={unitRef}
        style={{
          display: 'inline-flex',
          alignItems: 'center',
          gap: 1,
          flex: '0 0 auto',
          width: popup ? popup.width : undefined,
        }}
      >
        {popup ? null : children}
      </span>
      {popup &&
        createPortal(
          <div
            ref={setPopupRef}
            role="dialog"
            aria-label={label}
            tabIndex={-1}
            onKeyDown={(event) => {
              if (event.key !== 'Escape') return;
              event.preventDefault();
              event.stopPropagation();
              close(true);
            }}
            onBlur={(event) => {
              const next = event.relatedTarget as Node | null;
              if (next && !popupRef.current?.contains(next)) close(false);
            }}
            style={{
              position: 'fixed',
              top: position?.top ?? 0,
              left: position?.left ?? 0,
              visibility: position ? 'visible' : 'hidden',
              zIndex: 10000,
              display: 'flex',
              flexWrap: 'wrap',
              alignItems: 'center',
              gap: 2,
              maxWidth: 'calc(100vw - 16px)',
              padding: 6,
              border: `1px solid ${toolbarColors.border}`,
              borderRadius: 8,
              background: toolbarColors.surface,
              boxShadow: '0 4px 16px rgba(60, 64, 67, 0.24)',
              boxSizing: 'border-box',
              outline: 'none',
            }}
          >
            {children}
          </div>,
          document.body
        )}
    </ToolbarOverflowContext.Provider>
  );
}
