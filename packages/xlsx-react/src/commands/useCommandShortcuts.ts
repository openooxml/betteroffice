import { useEffect } from 'react';
import type { RefObject } from 'react';
import type { XlsxCommandController } from './createXlsxCommandStore';
import { commandForEvent, matchesChord } from './descriptors';
import { pluginEventStore } from './pluginEvents';
import type { XlsxCommandId } from './types';

/** Commands whose shortcuts also apply while a text field has focus. */
const FIELD_COMMANDS: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>(['save']);
const REPEATABLE: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>(['undo', 'redo']);

function isTextField(target: EventTarget | null): boolean {
  const element = target instanceof Element ? target : null;
  const field = element?.closest('input, textarea, select, [contenteditable]');
  return field != null && field.getAttribute('contenteditable') !== 'false';
}

/**
 * The editor's keyboard shortcuts, dispatched from the command descriptors,
 * then those of contributed commands. Only events inside this editor, its
 * plugins' contributions (portals included) or chrome registered for it are
 * handled, so several editors on one page never answer the same keystroke.
 * Composition, prevented events and text fields keep their own keys; the
 * in-cell editor and formula bar keep native text undo.
 */
export function useCommandShortcuts(
  controller: XlsxCommandController,
  root: RefObject<HTMLElement | null>
): void {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
      const target = event.target;
      const plugin = pluginEventStore(event);
      const owned = plugin
        ? controller.ownsStore(plugin)
        : (target instanceof Node && (root.current?.contains(target) ?? false)) ||
          controller.ownsChrome(target);
      if (!owned) return;
      const binding = commandForEvent(event);
      if (!binding) {
        const contributed = controller
          .pluginShortcuts()
          .find(({ chord }) => matchesChord(chord, event));
        if (!contributed || isTextField(target)) return;
        event.preventDefault();
        if (!event.repeat) void controller.store.execute(contributed.id, null);
        return;
      }
      if (isTextField(target) && !FIELD_COMMANDS.has(binding.id)) return;
      event.preventDefault();
      if (event.repeat && !REPEATABLE.has(binding.id)) return;
      void controller.store.execute(binding.id, binding.shortcut.args as never);
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, [controller, root]);
}
