import { useEffect, useLayoutEffect, useRef } from 'react';
import type { RefObject } from 'react';
import type { ShortcutIssue } from '../../../../shared/keyboard/resolve';
import {
  armClipboardGesture,
  disarmClipboardGesture,
  releaseKeyFromInput,
} from '../../../../shared/keyboard/nativeGesture';
import { pickChord, type PptxCommandController } from './createPptxCommandStore';
import { normalizeChord } from './descriptors';
import { pluginEventStore } from './pluginEvents';
import { CLIPBOARD_IDS, inputKeepsChord, type PptxShortcuts } from './shortcuts';
import type { PptxCommandId } from './types';

/** Commands whose shortcuts also work while a text field inside the editor has focus. */
const FIELD_COMMANDS: ReadonlySet<PptxCommandId> = new Set(['save']);
const REPEATABLE: ReadonlySet<PptxCommandId> = new Set(['undo', 'redo']);

/** The browser's own clipboard gesture for each default clipboard chord. */
const NATIVE_CHORDS: ReadonlyMap<string, string> = new Map([
  ['Mod+c', 'copy'],
  ['Mod+x', 'cut'],
  ['Mod+v', 'paste'],
  ['Mod+Shift+v', 'pasteAsPlainText'],
]);

function isTextField(target: EventTarget | null): boolean {
  const element = target instanceof Element ? target : null;
  const field = element?.closest('input, textarea, select, [contenteditable]');
  return Boolean(field) && field!.getAttribute('contenteditable') !== 'false';
}

function reportIssue(issue: ShortcutIssue): void {
  if (issue.code === 'displaced-default') return;
  if (typeof process !== 'undefined' && process.env?.NODE_ENV === 'production') return;
  console.warn(`Shortcut ${issue.chord ?? ''} of ${issue.id} was not applied: ${issue.code}`);
}

/**
 * The editor's keyboard shortcuts, dispatched from the effective bindings.
 * Only events inside this editor, its plugins' contributions (portals
 * included) or chrome registered for it are handled, so several editors on one
 * page never answer the same keystroke. Composition, prevented events and
 * suspended periods, such as a slideshow, pass through. Clipboard commands on
 * their default chords stay the browser's own gesture; a chord the host moved
 * away from the text input's built-in action is released from it.
 */
export function useCommandShortcuts({
  commands,
  containerRef,
  editorInput,
  suspended,
  shortcuts,
  onShortcutIssue,
  unsupported,
}: {
  commands: PptxCommandController;
  containerRef: RefObject<HTMLElement | null>;
  editorInput?: RefObject<HTMLElement | null>;
  suspended: () => boolean;
  shortcuts?: PptxShortcuts;
  onShortcutIssue?: (issue: ShortcutIssue) => void;
  /** Ids this editor cannot run, refused from the host map. */
  unsupported?: ReadonlySet<string>;
}) {
  const suspendedRef = useRef(suspended);
  suspendedRef.current = suspended;
  const shortcutsRef = useRef(shortcuts);
  shortcutsRef.current = shortcuts;
  const issueRef = useRef(onShortcutIssue);
  issueRef.current = onShortcutIssue;
  const shortcutsKey = shortcuts === undefined ? undefined : JSON.stringify(shortcuts);
  const reported = useRef(new Set<string>());

  useLayoutEffect(() => {
    for (const issue of commands.setShortcuts(shortcutsRef.current, unsupported)) {
      const key = JSON.stringify(issue);
      if (reported.current.has(key)) continue;
      reported.current.add(key);
      (issueRef.current ?? reportIssue)(issue);
    }
  }, [commands, shortcutsKey, unsupported]);

  useEffect(() => {
    const owns = (event: KeyboardEvent): boolean => {
      const plugin = pluginEventStore(event);
      if (plugin) return commands.ownsStore(plugin);
      const target = event.target;
      return (
        (target instanceof Node && (containerRef.current?.contains(target) ?? false)) ||
        commands.ownsChrome(target)
      );
    };
    const isInput = (target: EventTarget | null): target is HTMLElement =>
      target !== null && target === editorInput?.current;

    const handleInputChord = (event: KeyboardEvent) => {
      if (!owns(event)) return;
      const target = event.target;
      if (isInput(target)) disarmClipboardGesture(target);
      if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
      if (!event.ctrlKey && !event.metaKey && !event.altKey && !/^F\d/.test(event.key)) return;
      const pressed = commands.pressedChord(event);
      if (!pressed) {
        const contributed = commands.pluginShortcuts()
          .map(({ chord }) => normalizeChord(chord) ?? chord)
          .filter((chord) => !inputKeepsChord(chord));
        if (pickChord(contributed, event) !== null) releaseKeyFromInput(event);
        return;
      }
      const { owner, fallback } = pressed;
      if (owner === fallback) return;
      releaseKeyFromInput(event);
      if (isInput(target) && fallback !== undefined && CLIPBOARD_IDS.has(fallback)) {
        armClipboardGesture(target, 'ignore');
      }
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
      if (!owns(event) || suspendedRef.current()) return;
      const target = event.target;
      const field = isTextField(target) && !isInput(target);
      const binding = commands.shortcutAt(event);
      if (!binding) {
        const list = commands.pluginShortcuts();
        const chord = pickChord(list.map(({ chord }) => normalizeChord(chord) ?? chord), event);
        const contributed = list.find(({ chord: own }) => (normalizeChord(own) ?? own) === chord);
        if (!contributed || field) return;
        event.preventDefault();
        if (isInput(target)) disarmClipboardGesture(target);
        if (!event.repeat) void commands.store.execute(contributed.id, null);
        return;
      }
      const { id, chord } = binding;
      if (NATIVE_CHORDS.get(chord) === id) return;
      if (field && !FIELD_COMMANDS.has(id)) return;
      const state = commands.store.getState(id, null as never);
      if (!state.enabled && (state.disabledReason.code === 'host-disabled' ||
        state.disabledReason.code === 'unsupported-command')) return;
      event.preventDefault();
      if (isInput(target)) disarmClipboardGesture(target);
      if (event.repeat && !REPEATABLE.has(id)) return;
      void commands.store.execute(id, null as never);
    };

    document.addEventListener('keydown', handleInputChord, true);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleInputChord, true);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [commands, containerRef, editorInput]);
}
