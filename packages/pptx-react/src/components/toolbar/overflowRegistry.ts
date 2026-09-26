import { createContext, useContext, useLayoutEffect, useRef } from 'react';
import type { RefObject } from 'react';
import type { OverflowMenuEntry } from '../../../../../shared/react-toolbar/OverflowMenu';

/** A control's representation in the overflow menu. */
export interface OverflowSource {
  element(): HTMLElement | null;
  entries(): readonly OverflowMenuEntry[];
}

/** A value an overflow action asks for in a dialog, such as a custom color. */
export interface OverflowPrompt {
  title: string;
  label: string;
  placeholder?: string;
  initialValue?: string;
  valid(value: string): boolean;
  submit(value: string): void;
}

export interface OverflowRegistry {
  register(source: OverflowSource): () => void;
  /** Whether the menu currently holds hidden controls. */
  overflowing(): boolean;
  /** A registered source now presents different entries. */
  changed(): void;
  /** Asks for a value in an accessible dialog, then returns focus to the menu button. */
  prompt(request: OverflowPrompt): void;
  /** The "More" button, where popups of hidden controls open. */
  anchor(): HTMLElement | null;
}

export const ToolbarOverflowContext = createContext<OverflowRegistry | null>(null);

function findItem(
  entries: readonly OverflowMenuEntry[],
  id: string
): Extract<OverflowMenuEntry, { kind: 'item' }> | null {
  for (const entry of entries) {
    if (entry.kind === 'item' && entry.id === id) return entry;
    if (entry.kind === 'group' || entry.kind === 'submenu') {
      const found = findItem(entry.entries, id);
      if (found) return found;
    }
  }
  return null;
}

/** Items that act through the source's latest entries, so a rendered menu never runs a stale action. */
function bindLatest(
  entries: readonly OverflowMenuEntry[],
  latest: () => readonly OverflowMenuEntry[]
): OverflowMenuEntry[] {
  return entries.map((entry) => {
    if (entry.kind === 'item') {
      return {
        ...entry,
        onSelect: () => {
          const current = findItem(latest(), entry.id);
          if (current && !current.disabled) current.onSelect();
        },
      };
    }
    if (entry.kind === 'group' || entry.kind === 'submenu') {
      return { ...entry, entries: bindLatest(entry.entries, latest) };
    }
    return entry;
  });
}

function signature(entries: readonly OverflowMenuEntry[]): string {
  return JSON.stringify(entries, (key, value) =>
    key === 'icon' || typeof value === 'function' ? undefined : value
  );
}

/** Registers the overflow presentation of a toolbar control while it is mounted. */
export function useOverflowSource(
  element: RefObject<HTMLElement | null>,
  entries: (() => readonly OverflowMenuEntry[]) | null
): void {
  const registry = useContext(ToolbarOverflowContext);
  const entriesRef = useRef(entries);
  entriesRef.current = entries;
  const published = useRef<string | null>(null);
  const enabled = entries !== null;
  useLayoutEffect(() => {
    if (!registry || !enabled) return;
    const latest = () => entriesRef.current?.() ?? [];
    published.current = null;
    return registry.register({
      element: () => element.current,
      entries: () => bindLatest(latest(), latest),
    });
  }, [registry, element, enabled]);
  useLayoutEffect(() => {
    if (!registry || !enabled) return;
    if (!registry.overflowing()) {
      published.current = null;
      return;
    }
    const next = signature(entriesRef.current?.() ?? []);
    if (published.current !== next) registry.changed();
    published.current = next;
  });
}

/** The surrounding toolbar's overflow registry, if any. */
export function useOverflowRegistry(): OverflowRegistry | null {
  return useContext(ToolbarOverflowContext);
}
