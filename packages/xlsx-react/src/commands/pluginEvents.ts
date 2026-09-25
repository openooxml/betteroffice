import type { XlsxCommandStore } from './types';

const claims = new WeakMap<Event, XlsxCommandStore>();

/**
 * Marks an event raised inside a plugin contribution, portals included, with that plugin's
 * store, which identifies the editor the event belongs to.
 */
export function claimPluginEvent(event: Event, store: XlsxCommandStore): void {
  if (!claims.has(event)) claims.set(event, store);
}

/** The store of the plugin whose contribution raised `event`; null for other input. */
export function pluginEventStore(event: Event): XlsxCommandStore | null {
  return claims.get(event) ?? null;
}

/** Marks plugin chrome the editor draws, such as docks and the overlay layer. */
export const PLUGIN_CHROME = { 'data-xlsx-plugin-chrome': '' } as const;

/** Whether `target` sits in plugin chrome, whose input never reaches the grid. */
export function inPluginChrome(target: EventTarget | null): boolean {
  return target instanceof Element && target.closest('[data-xlsx-plugin-chrome]') !== null;
}
