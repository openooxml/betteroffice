import type { DisplayListQueries } from '@betteroffice/docx/layout/render';

export interface DisplayWindow {
  read(): readonly [number, number];
  subscribe(listener: () => void): () => void;
}

const displayWindows = new WeakMap<DisplayListQueries, DisplayWindow>();

/** The live page window used to build this query facade's display. */
export function bindDisplayWindow(queries: DisplayListQueries, window: DisplayWindow): void {
  displayWindows.set(queries, window);
}

export function displayWindowOf(queries: DisplayListQueries | null): DisplayWindow | null {
  return queries ? (displayWindows.get(queries) ?? null) : null;
}
