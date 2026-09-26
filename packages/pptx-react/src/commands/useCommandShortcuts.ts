import { useEffect, useRef } from 'react';
import type { RefObject } from 'react';
import type { PptxCommandController } from './createPptxCommandStore';
import { matchesChord, PPTX_COMMAND_DESCRIPTORS } from './descriptors';
import type { PptxCommandId, PptxCommandShortcut } from './types';

const BINDINGS = Object.values(PPTX_COMMAND_DESCRIPTORS).flatMap((descriptor) =>
  (descriptor.shortcuts as readonly PptxCommandShortcut[]).map((shortcut) => ({
    id: descriptor.id as PptxCommandId,
    shortcut,
  }))
);

/** Commands whose shortcuts also work while a text field inside the editor has focus. */
const FIELD_COMMANDS: ReadonlySet<PptxCommandId> = new Set(['save']);
const REPEATABLE: ReadonlySet<PptxCommandId> = new Set(['undo', 'redo']);

function isTextField(target: EventTarget | null): boolean {
  const element = target instanceof Element ? target : null;
  const field = element?.closest('input, textarea, select, [contenteditable]');
  return Boolean(field) && field!.getAttribute('contenteditable') !== 'false';
}

/**
 * The editor's keyboard shortcuts, dispatched from the command descriptors.
 * Only events inside this editor or chrome registered for it are handled, so
 * several editors on one page never answer the same keystroke. Composition,
 * prevented events and suspended periods, such as a slideshow, pass through.
 */
export function useCommandShortcuts({
  commands,
  containerRef,
  suspended,
}: {
  commands: PptxCommandController;
  containerRef: RefObject<HTMLElement | null>;
  suspended: () => boolean;
}) {
  const suspendedRef = useRef(suspended);
  suspendedRef.current = suspended;

  useEffect(() => {
    const owns = (target: EventTarget | null): boolean =>
      (target instanceof Node && (containerRef.current?.contains(target) ?? false)) ||
      commands.ownsChrome(target);

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
      if (!owns(event.target) || suspendedRef.current()) return;
      const binding = BINDINGS.find(({ shortcut }) => matchesChord(shortcut.chord, event));
      if (!binding) return;
      const { id, shortcut } = binding;
      if (isTextField(event.target) && !FIELD_COMMANDS.has(id)) return;
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
  }, [commands, containerRef]);
}
