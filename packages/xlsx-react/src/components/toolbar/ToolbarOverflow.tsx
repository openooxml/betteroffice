import {
  useCallback,
  useContext,
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
import {
  representsUnit,
  useToolbarOverflow,
} from '../../../../../shared/react-toolbar/useToolbarOverflow';
import { xlsxCommandController } from '../../commands/createXlsxCommandStore';
import { useXlsxCommands } from '../../commands/hooks';
import { useTranslation } from '../../i18n';
import { ToolbarIcon } from '../ui/ToolbarIcon';
import { toolbarColors } from '../ui/ToolbarPrimitives';
import {
  ToolbarOverflowContext,
  useOverflowSource,
  type OverflowColorRequest,
  type OverflowPrompt,
  type OverflowRegistry,
  type OverflowSource,
} from './overflowRegistry';
import { openColorPicker } from './ToolbarCommand';
import { ToolbarPromptDialog } from './ToolbarPromptDialog';

const MENU_COLORS = {
  surface: toolbarColors.surface,
  text: toolbarColors.text,
  mutedText: toolbarColors.muted,
  border: toolbarColors.border,
  hover: toolbarColors.hover,
  shadow: 'rgba(60, 64, 67, 0.24)',
};

const MORE_STYLE: CSSProperties = {
  appearance: 'none',
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  width: 28,
  height: 28,
  padding: 0,
  border: 0,
  borderRadius: 4,
  background: 'transparent',
  color: toolbarColors.text,
  cursor: 'pointer',
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
}

/** The formatting rail: host-ordered controls with an accessible "More" overflow menu. */
export function ToolbarRail({ children, className, style }: ToolbarRailProps) {
  const { t } = useTranslation();
  const store = useXlsxCommands();
  const itemsRef = useRef<HTMLDivElement>(null);
  const moreRef = useRef<HTMLSpanElement>(null);
  const menuRef = useRef<OverflowMenuHandle>(null);
  const railRef = useRef<HTMLDivElement>(null);
  const popupRef = useRef<HTMLDivElement>(null);
  const colorRef = useRef<HTMLInputElement>(null);
  const colorRequest = useRef<OverflowColorRequest | null>(null);
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
      pickColor(request) {
        colorRequest.current = request;
        openColorPicker(colorRef.current, request.value);
      },
      popupHost: () => popupRef.current,
      focusMore() {
        if (menuRef.current && moreRef.current?.isConnected) menuRef.current.focus();
        else railRef.current?.focus({ preventScroll: true });
      },
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
    canHide: (unit) =>
      unit.getAttribute('role') === 'separator' ||
      representsUnit(unit, sourcesIn(unit).map(({ element }) => element)),
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
    entries.push({ kind: 'group', id: `unit-${index}`, label: heading ?? undefined, entries: unitEntries });
  });
  const menuEntries = entries.flatMap((entry, index) =>
    index === 0 ? [entry] : [{ kind: 'separator' as const, id: `separator-${index}` }, entry]
  );

  return (
    <div
      ref={railRef}
      className={className}
      role="toolbar"
      aria-label={t('toolbar.actionsLabel')}
      data-testid="xlsx-formatting-toolbar"
      tabIndex={-1}
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
        overflow: 'hidden',
        boxSizing: 'border-box',
        outline: 'none',
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
            flex: 1,
            minWidth: 0,
            overflow: 'hidden',
          }}
        >
          {children}
        </div>
      </ToolbarOverflowContext.Provider>
      {hidden.length > 0 && (
        <span ref={moreRef} style={{ display: 'inline-flex', flex: '0 0 auto' }}>
          <OverflowMenu
            ref={menuRef}
            label={t('toolbar.more')}
            trigger={<ToolbarIcon name="more" />}
            entries={menuEntries}
            colors={MENU_COLORS}
            triggerStyle={MORE_STYLE}
            testId="xlsx-toolbar-more"
            onItemSelected={(modality) => {
              if (modality === 'pointer') xlsxCommandController(store)?.focusEditor();
            }}
          />
        </span>
      )}
      <div ref={popupRef} />
      <input
        ref={colorRef}
        type="color"
        tabIndex={-1}
        aria-hidden="true"
        onChange={(event) => colorRequest.current?.submit(event.target.value)}
        style={{
          position: 'absolute',
          width: 1,
          height: 1,
          padding: 0,
          border: 0,
          opacity: 0,
          pointerEvents: 'none',
        }}
      />
      {prompt && (
        <ToolbarPromptDialog
          prompt={prompt}
          onClose={() => {
            setPrompt(null);
            registry.focusMore();
          }}
        />
      )}
    </div>
  );
}

/** @experimental */
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
 * @experimental
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
    <span ref={elementRef} style={{ display: 'inline-flex', alignItems: 'center', flex: '0 0 auto' }}>
      {children}
    </span>
  );
}

const PANEL_FOCUSABLE =
  'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [href], [tabindex]:not([tabindex="-1"])';

/**
 * Legacy content appended to the default controls. At narrow widths it moves,
 * still mounted, into a dialog opened from the "More" menu.
 */
export function ToolbarLegacyContent({ children }: { children: ReactNode }) {
  const { t } = useTranslation();
  const registry = useContext(ToolbarOverflowContext);
  const unitRef = useRef<HTMLSpanElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const id = useId();
  const [host] = useState(() =>
    typeof document === 'undefined' ? null : document.createElement('span')
  );
  const [open, setOpen] = useState(false);
  const label = t('commands.moreTools');
  useOverflowSource(unitRef, () => [{ kind: 'item', id, label, onSelect: () => setOpen(true) }]);

  useLayoutEffect(() => {
    const unit = unitRef.current;
    const target = open ? panelRef.current : unit;
    if (!host || !unit || !target) return;
    if (open && !unit.style.minWidth) unit.style.minWidth = `${unit.getBoundingClientRect().width}px`;
    if (!open) unit.style.minWidth = '';
    if (host.parentNode !== target) target.appendChild(host);
    if (open) panelRef.current?.querySelector<HTMLElement>(PANEL_FOCUSABLE)?.focus({ preventScroll: true });
  }, [open, host]);

  useEffect(() => {
    if (!open) return;
    const close = (event: Event) => {
      if (event.target instanceof Node && panelRef.current?.contains(event.target)) return;
      setOpen(false);
    };
    const resize = () => setOpen(false);
    document.addEventListener('mousedown', close);
    window.addEventListener('resize', resize);
    return () => {
      document.removeEventListener('mousedown', close);
      window.removeEventListener('resize', resize);
    };
  }, [open]);

  const popupHost = registry?.popupHost() ?? null;
  const moreRect = open ? popupHost?.parentElement?.getBoundingClientRect() : undefined;
  return (
    <>
      <span
        ref={unitRef}
        style={{ display: 'inline-flex', alignItems: 'center', flex: '0 0 auto' }}
      />
      {host &&
        createPortal(
          <ToolbarOverflowContext.Provider value={null}>{children}</ToolbarOverflowContext.Provider>,
          host
        )}
      {open &&
        popupHost &&
        createPortal(
          <div
            ref={panelRef}
            role="dialog"
            aria-label={label}
            onKeyDown={(event) => {
              if (event.key !== 'Escape') return;
              event.preventDefault();
              event.stopPropagation();
              setOpen(false);
              registry?.focusMore();
            }}
            onBlur={(event) => {
              const next = event.relatedTarget as Node | null;
              if (next && !panelRef.current?.contains(next)) setOpen(false);
            }}
            style={{
              position: 'fixed',
              top: (moreRect?.bottom ?? 0) + 4,
              right: 8,
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
            }}
          />,
          popupHost
        )}
    </>
  );
}
