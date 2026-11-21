import { useEffect, useLayoutEffect, useRef } from 'react';
import type { useTableSelection } from '../../../hooks/useTableSelection';
import { isMacPlatform, matchesChord } from '../../../commands/descriptors';
import type { DocxCommandController } from '../../../commands/createDocxCommandStore';
import type { DocxShortcuts } from '../../../commands/shortcuts';
import type { DocxCommandId } from '../../../commands/types';
import type { ShortcutIssue } from '../../../../../../shared/keyboard/resolve';
import {
  armClipboardGesture,
  disarmClipboardGesture,
  releaseKeyFromInput,
} from '../../../../../../shared/keyboard/nativeGesture';
import type { PagedEditorRef } from '../PagedEditor';

/** Commands whose shortcuts also work while a text field inside the editor has focus. */
const FIELD_COMMANDS: ReadonlySet<DocxCommandId> = new Set(['save', 'open', 'find', 'replace']);
const REPEATABLE: ReadonlySet<DocxCommandId> = new Set(['undo', 'redo']);
const FIND_REPLACE: ReadonlySet<DocxCommandId> = new Set(['find', 'replace']);

const CLIPBOARD_ACTIONS: ReadonlySet<string> = new Set([
  'copy',
  'cut',
  'paste',
  'pasteAsPlainText',
]);
const PLAIN_PASTE_CHORDS: ReadonlySet<string> = new Set(['Mod+v', 'Mod+Shift+v']);

const NATIVE_CHORDS: ReadonlyMap<string, DocxCommandId> = new Map<string, DocxCommandId>([
  ['Mod+a', 'selectAll'],
  ['Mod+c', 'copy'],
  ['Mod+x', 'cut'],
  ['Mod+v', 'paste'],
]);

/** Whether `id` on `chord` is the browser's own gesture, which needs no permission. */
function isNativeGesture(id: DocxCommandId, chord: string): boolean {
  if (id === 'pasteAsPlainText') return PLAIN_PASTE_CHORDS.has(chord);
  return NATIVE_CHORDS.get(chord) === id;
}

function isEditorInput(target: EventTarget | null): target is Element {
  return target instanceof Element && target.matches('.paged-editor__yrs-input, .paged-editor__input');
}

function reportIssue(issue: ShortcutIssue): void {
  if (issue.code === 'displaced-default') return;
  if (typeof process !== 'undefined' && process.env?.NODE_ENV === 'production') return;
  console.warn(`Shortcut ${issue.chord ?? ''} of ${issue.id} was not applied: ${issue.code}`);
}

function isForeignTextField(target: EventTarget | null): boolean {
  const element = target instanceof Element ? target : null;
  const field = element?.closest('input, textarea, select, [contenteditable]');
  if (!field || field.matches('.paged-editor__yrs-input, .paged-editor__input')) return false;
  return field.getAttribute('contenteditable') !== 'false';
}

/**
 * The editor's keyboard shortcuts, dispatched from the command descriptors.
 * Only events targeting this editor, its input or chrome registered for it are
 * handled, so several editors on one page never answer the same keystroke.
 * Composition, prevented events and host opt-outs pass through.
 */
export function useKeyboardShortcuts({
  commands,
  pagedEditorRef,
  containerRef,
  disableFindReplaceShortcuts,
  tableSelection,
  shortcuts,
  onShortcutIssue,
  onError,
}: {
  commands: DocxCommandController;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  containerRef?: React.RefObject<HTMLElement | null>;
  disableFindReplaceShortcuts: boolean;
  tableSelection?: ReturnType<typeof useTableSelection>;
  shortcuts?: DocxShortcuts;
  onShortcutIssue?: (issue: ShortcutIssue) => void;
  onError?: (error: Error, info: { recoverable: true }) => void;
}) {
  const errorRef = useRef(onError);
  errorRef.current = onError;
  const optionsRef = useRef({ disableFindReplaceShortcuts, tableSelection, shortcuts });
  optionsRef.current = { disableFindReplaceShortcuts, tableSelection, shortcuts };
  const issueRef = useRef(onShortcutIssue);
  issueRef.current = onShortcutIssue;
  const shortcutsKey = shortcuts === undefined ? undefined : JSON.stringify(shortcuts);
  const reportedIssues = useRef(new Set<string>());

  useLayoutEffect(() => {
    for (const issue of commands.setShortcuts(optionsRef.current.shortcuts)) {
      const key = JSON.stringify(issue);
      if (reportedIssues.current.has(key)) continue;
      reportedIssues.current.add(key);
      (issueRef.current ?? reportIssue)(issue);
    }
  }, [commands, shortcutsKey]);

  useEffect(() => {
    const owns = (target: EventTarget | null): boolean =>
      (pagedEditorRef.current?.isFocused() ?? false) ||
      (target instanceof Node && (containerRef?.current?.contains(target) ?? false)) ||
      commands.ownsChrome(target);

    const handleInputChord = (event: KeyboardEvent) => {
      if (!isEditorInput(event.target) || !owns(event.target)) return;
      disarmClipboardGesture(event.target);
      if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
      if (!event.ctrlKey && !event.metaKey && !event.altKey && !/^F\d/.test(event.key)) return;
      const mac = isMacPlatform();
      const pressed = commands.pressedChord(event, mac) ?? commands.pressedChord(event, !mac);
      if (!pressed) {
        const contributed = commands.pluginShortcuts();
        if (contributed.some(({ chord }) => matchesChord(chord, event, mac))) releaseKeyFromInput(event);
        return;
      }
      const { chord, owner, fallback } = pressed;
      if (owner !== fallback) releaseKeyFromInput(event);
      if (fallback === undefined || !CLIPBOARD_ACTIONS.has(fallback)) return;
      if (owner === 'pasteAsPlainText' && PLAIN_PASTE_CHORDS.has(chord)) {
        armClipboardGesture(event.target, 'plain');
      } else if (owner !== fallback) {
        armClipboardGesture(event.target, 'ignore');
      }
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
      if (!owns(event.target)) return;
      const { disableFindReplaceShortcuts: noFind, tableSelection: table, shortcuts: map } =
        optionsRef.current;
      const plain = !event.ctrlKey && !event.metaKey && !event.shiftKey && !event.altKey;
      if (plain && (event.key === 'Delete' || event.key === 'Backspace')) {
        if (table && table.state.tableIndex !== null && !isForeignTextField(event.target)) {
          event.preventDefault();
          table.handleAction('deleteTable');
        }
        return;
      }
      const mac = isMacPlatform();
      const binding = commands.shortcutAt(event, mac);
      if (!binding) {
        const contributed = commands
          .pluginShortcuts()
          .find(({ chord }) => matchesChord(chord, event, mac));
        if (!contributed || isForeignTextField(event.target)) return;
        event.preventDefault();
        if (isEditorInput(event.target)) disarmClipboardGesture(event.target);
        if (!event.repeat) void commands.store.execute(contributed.id, null);
        return;
      }
      const { id, args, chord } = binding;
      if (isNativeGesture(id, chord)) return;
      if (isForeignTextField(event.target) && !FIELD_COMMANDS.has(id)) return;
      if (noFind && FIND_REPLACE.has(id) && map?.[id as 'find'] === undefined) return;
      const state = commands.store.getState(id, args as never);
      if (!state.enabled && state.disabledReason.code === 'host-disabled') return;
      event.preventDefault();
      if (isEditorInput(event.target)) disarmClipboardGesture(event.target);
      if (event.repeat && !REPEATABLE.has(id)) return;
      void commands.store.execute(id, args as never).then((result) => {
        if (!result.ok && (CLIPBOARD_ACTIONS.has(id) || id === 'selectAll')) {
          errorRef.current?.(new Error(result.failure.message), { recoverable: true });
        }
      });
    };

    document.addEventListener('keydown', handleInputChord, true);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleInputChord, true);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [commands, containerRef, pagedEditorRef]);
}
