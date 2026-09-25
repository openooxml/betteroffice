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
import {
  OverflowMenu,
  type OverflowMenuEntry,
  type OverflowMenuHandle,
} from '../../../../../shared/react-toolbar/OverflowMenu';
import { useToolbarOverflow } from '../../../../../shared/react-toolbar/useToolbarOverflow';
import { docxCommandController } from '../../commands/createDocxCommandStore';
import { useDocxCommands } from '../../commands/hooks';
import { useTranslation } from '../../i18n';
import { cn } from '../../lib/utils';
import { MaterialSymbol } from '../ui/MaterialSymbol';
import {
  ToolbarOverflowContext,
  useOverflowSource,
  type OverflowPrompt,
  type OverflowRegistry,
  type OverflowSource,
} from './overflowRegistry';
import { ToolbarPromptDialog } from './ToolbarPromptDialog';

const MENU_COLORS = {
  surface: 'var(--doc-surface)',
  text: 'var(--doc-text)',
  mutedText: 'var(--doc-text-muted)',
  border: 'var(--doc-border)',
  hover: 'var(--doc-bg-hover)',
  shadow: 'var(--doc-shadow)',
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
  label?: string;
  testId?: string;
}

/** The toolbar row: host-ordered controls with an accessible "More" overflow menu. */
export function ToolbarRail({ children, className, style, label, testId }: ToolbarRailProps) {
  const { t } = useTranslation();
  const store = useDocxCommands();
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
      className={cn(
        'flex items-center px-2 py-1 bg-muted rounded-full min-h-[36px] mx-2 mb-1 min-w-0',
        className
      )}
      style={style}
      role="toolbar"
      aria-label={label ?? t('toolbar.ariaLabel')}
      data-testid={testId}
      onMouseDown={(event) => {
        if (!isInteractive(event.target)) event.preventDefault();
      }}
    >
      <ToolbarOverflowContext.Provider value={registry}>
        <div
          ref={itemsRef}
          className="relative flex min-w-0 flex-1 items-center overflow-hidden"
          data-toolbar-items=""
        >
          {children}
        </div>
      </ToolbarOverflowContext.Provider>
      {hidden.length > 0 && (
        <span ref={moreRef} className="inline-flex flex-shrink-0">
          <OverflowMenu
            ref={menuRef}
            label={t('commands.moreActions')}
            trigger={<MaterialSymbol name="more_vert" size={18} />}
            entries={menuEntries}
            colors={MENU_COLORS}
            triggerClassName="oox-toolbar-toggle inline-flex h-7 w-7 items-center justify-center rounded-md text-muted-foreground"
            testId="toolbar-more"
            onItemSelected={(modality) => {
              if (modality === 'pointer') docxCommandController(store)?.focusEditor();
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
    <span ref={elementRef} className="inline-flex flex-shrink-0 items-center">
      {children}
    </span>
  );
}
