import { useEffect, useRef } from 'react';
import type { useTableSelection } from '../../../hooks/useTableSelection';
import { DOCX_COMMAND_DESCRIPTORS, matchesChord } from '../../../commands/descriptors';
import type { DocxCommandController } from '../../../commands/createDocxCommandStore';
import type { DocxCommandId, DocxCommandShortcut } from '../../../commands/types';
import type { PagedEditorRef } from '../PagedEditor';

const BINDINGS = Object.values(DOCX_COMMAND_DESCRIPTORS).flatMap((descriptor) =>
  (descriptor.shortcuts as readonly DocxCommandShortcut[]).map((shortcut) => ({
    id: descriptor.id as DocxCommandId,
    shortcut,
  }))
);

/** Commands whose shortcuts also work while a text field inside the editor has focus. */
const FIELD_COMMANDS: ReadonlySet<DocxCommandId> = new Set(['save', 'open', 'find', 'replace']);
const REPEATABLE: ReadonlySet<DocxCommandId> = new Set(['undo', 'redo']);
const FIND_REPLACE: ReadonlySet<DocxCommandId> = new Set(['find', 'replace']);

function isForeignTextField(target: EventTarget | null): boolean {
  const element = target instanceof Element ? target : null;
  const field = element?.closest('input, textarea, select, [contenteditable]');
  if (!field || field.matches('.paged-editor__yrs-input')) return false;
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
}: {
  commands: DocxCommandController;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  containerRef?: React.RefObject<HTMLElement | null>;
  disableFindReplaceShortcuts: boolean;
  tableSelection: ReturnType<typeof useTableSelection>;
}) {
  const optionsRef = useRef({ disableFindReplaceShortcuts, tableSelection });
  optionsRef.current = { disableFindReplaceShortcuts, tableSelection };

  useEffect(() => {
    const owns = (target: EventTarget | null): boolean =>
      (pagedEditorRef.current?.isFocused() ?? false) ||
      (target instanceof Node && (containerRef?.current?.contains(target) ?? false)) ||
      commands.ownsChrome(target);

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
      if (!owns(event.target)) return;
      const { disableFindReplaceShortcuts: noFind, tableSelection: table } = optionsRef.current;
      const plain = !event.ctrlKey && !event.metaKey && !event.shiftKey && !event.altKey;
      if (plain && (event.key === 'Delete' || event.key === 'Backspace')) {
        if (table.state.tableIndex !== null && !isForeignTextField(event.target)) {
          event.preventDefault();
          table.handleAction('deleteTable');
        }
        return;
      }
      const binding = BINDINGS.find(({ shortcut }) => matchesChord(shortcut.chord, event));
      if (!binding) {
        const contributed = commands
          .pluginShortcuts()
          .find(({ chord }) => matchesChord(chord, event));
        if (!contributed || isForeignTextField(event.target)) return;
        event.preventDefault();
        if (!event.repeat) void commands.store.execute(contributed.id, null);
        return;
      }
      const { id, shortcut } = binding;
      if (isForeignTextField(event.target) && !FIELD_COMMANDS.has(id)) return;
      if (noFind && FIND_REPLACE.has(id)) return;
      const state = commands.store.getState(id, shortcut.args as never);
      if (!state.enabled && state.disabledReason.code === 'host-disabled') return;
      event.preventDefault();
      if (event.repeat && !REPEATABLE.has(id)) return;
      void commands.store.execute(id, shortcut.args as never);
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [commands, containerRef, pagedEditorRef]);
}
