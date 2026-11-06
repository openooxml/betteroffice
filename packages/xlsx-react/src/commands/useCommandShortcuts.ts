import { useEffect, useLayoutEffect, useRef } from 'react';
import type { RefObject } from 'react';
import type { ShortcutIssue } from '../../../../shared/keyboard/resolve';
import type { XlsxCommandController } from './createXlsxCommandStore';
import { matchesChord } from './descriptors';
import { pluginEventStore } from './pluginEvents';
import type { XlsxShortcuts } from './shortcuts';
import type { XlsxCommandId } from './types';

/** Commands whose shortcuts also apply while a text field has focus. */
const FIELD_COMMANDS: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>(['save']);
const REPEATABLE: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>(['undo', 'redo']);

/** Host shortcut settings of one editor. */
export interface XlsxShortcutOptions {
  shortcuts?: XlsxShortcuts;
  onShortcutIssue?: (issue: ShortcutIssue) => void;
}

function isTextField(target: EventTarget | null): boolean {
  const element = target instanceof Element ? target : null;
  const field = element?.closest('input, textarea, select, [contenteditable]');
  return field != null && field.getAttribute('contenteditable') !== 'false';
}

function reportIssue(issue: ShortcutIssue): void {
  if (issue.code === 'displaced-default') return;
  if (typeof process !== 'undefined' && process.env?.NODE_ENV === 'production') return;
  console.warn(`Shortcut ${issue.chord ?? ''} of ${issue.id} was not applied: ${issue.code}`);
}

/**
 * The editor's keyboard shortcuts, dispatched from the effective bindings,
 * then those of contributed commands. Only events inside this editor, its
 * plugins' contributions (portals included) or chrome registered for it are
 * handled, so several editors on one page never answer the same keystroke.
 * Composition, prevented events and text fields keep their own keys; the
 * in-cell editor and formula bar keep native text undo.
 */
export function useCommandShortcuts(
  controller: XlsxCommandController,
  root: RefObject<HTMLElement | null>,
  options: XlsxShortcutOptions = {}
): void {
  const { shortcuts, onShortcutIssue } = options;
  const shortcutsRef = useRef(shortcuts);
  shortcutsRef.current = shortcuts;
  const issueRef = useRef(onShortcutIssue);
  issueRef.current = onShortcutIssue;
  const shortcutsKey = shortcuts === undefined ? undefined : JSON.stringify(shortcuts);
  const reported = useRef(new Set<string>());

  useLayoutEffect(() => {
    for (const issue of controller.setShortcuts(shortcutsRef.current)) {
      const key = JSON.stringify(issue);
      if (reported.current.has(key)) continue;
      reported.current.add(key);
      (issueRef.current ?? reportIssue)(issue);
    }
  }, [controller, shortcutsKey]);

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
      const binding = controller.shortcutAt(event);
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
      const state = controller.store.getState(binding.id, binding.args as never);
      if (!state.enabled && state.disabledReason.code === 'unsupported-command') return;
      event.preventDefault();
      if (event.repeat && !REPEATABLE.has(binding.id)) return;
      void controller.store.execute(binding.id, binding.args as never);
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, [controller, root]);
}
