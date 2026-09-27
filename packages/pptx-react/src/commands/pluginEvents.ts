import type { PptxCommandStore } from './types';

const claims = new WeakMap<Event, PptxCommandStore>();

/**
 * Marks a keyboard event raised inside a plugin contribution, portals included, with that
 * plugin's store, which identifies the editor the event belongs to.
 */
export function claimPluginEvent(event: Event, store: PptxCommandStore): void {
  if (!claims.has(event)) claims.set(event, store);
}

/** The store of the plugin whose contribution raised `event`; null for other input. */
export function pluginEventStore(event: Event): PptxCommandStore | null {
  return claims.get(event) ?? null;
}

/** Marks plugin chrome the editor draws, such as docks and the overlay layer. */
export const PLUGIN_CHROME = { 'data-pptx-plugin-chrome': '' } as const;

/** Whether `target` sits in plugin chrome, whose keys never edit the slide's text. */
export function inPluginChrome(target: EventTarget | null): boolean {
  return target instanceof Element && target.closest('[data-pptx-plugin-chrome]') !== null;
}
