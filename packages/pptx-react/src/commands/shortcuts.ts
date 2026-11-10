import { isMacPlatform } from '../../../../shared/keyboard/format';
import { inputKeepsChord } from './shortcutKeys';

export { CLIPBOARD_IDS, inputKeepsChord } from './shortcutKeys';
import {
  resolveShortcuts,
  type ResolvedShortcuts,
  type ShortcutBinding,
  type ShortcutDefault,
  type ShortcutIssue,
} from '../../../../shared/keyboard/resolve';
import { isPluginCommandId, isPptxCommandId, PPTX_COMMAND_DESCRIPTORS } from './descriptors';
import type { PptxCommandArgs, PptxCommandId, PptxCommandShortcut, PptxPluginCommandId } from './types';

type PptxNullArgsCommandId = {
  [K in PptxCommandId]: null extends PptxCommandArgs[K] ? K : never;
}[PptxCommandId];

/**
 * Every id a host can bind: commands that run without arguments and plugin commands.
 * @experimental
 */
export type PptxShortcutId = PptxNullArgsCommandId | PptxPluginCommandId;

/**
 * Host changes to the editor's shortcuts. An omitted id keeps its default; chords replace all of
 * its defaults; `null` disables it and leaves the keystroke to the host or browser. A chord taken
 * from another command's default wins over it. `cut`, `copy`, `paste` and `pasteAsPlainText` on a
 * chord other than their default use the async Clipboard API, which browsers may gate behind a
 * prompt; the non-worker editor and the viewer have no clipboard and report them as `unknown-id`.
 * Text navigation and editing keys, select all and the context menu key cannot be bound.
 * @experimental
 */
export type PptxShortcuts = Readonly<Partial<Record<PptxShortcutId, ShortcutBinding>>>;

function shortcutDefaults(): ShortcutDefault[] {
  return Object.values(PPTX_COMMAND_DESCRIPTORS).flatMap((descriptor): ShortcutDefault[] => {
    const shortcuts = descriptor.shortcuts as readonly PptxCommandShortcut[];
    return shortcuts.length === 0
      ? []
      : [{ id: descriptor.id, chords: shortcuts.map(({ chord }) => chord) }];
  });
}

const NULL_ARGS: Readonly<Record<PptxNullArgsCommandId, true>> = {
  cut: true, copy: true, paste: true, pasteAsPlainText: true, deleteObject: true, deleteSlide: true,
  duplicateSlide: true, bold: true, italic: true, underline: true, strikethrough: true, insertImage: true,
  save: true, exportPng: true, slideshow: true, undo: true, redo: true, proposalsPanel: true,
};

function isBindable(id: string): boolean {
  return isPluginCommandId(id) || (isPptxCommandId(id) && Object.hasOwn(NULL_ARGS, id));
}

/**
 * The effective bindings a host map produces and what was wrong with it, for checking a map
 * outside the editor.
 * @experimental
 */
export function resolvePptxShortcuts(
  shortcuts: PptxShortcuts | undefined,
  options: { mac?: boolean; unsupported?: ReadonlySet<string> } = {}
): ResolvedShortcuts {
  const { unsupported } = options;
  const refused: ShortcutIssue[] = [];
  const map = shortcuts && unsupported
    ? Object.fromEntries(Object.entries(shortcuts).filter(([id, binding]) => {
      if (binding === undefined || !unsupported.has(id)) return true;
      refused.push({ code: 'unknown-id', id });
      return false;
    }))
    : shortcuts;
  const resolved = resolveShortcuts(shortcutDefaults(), map, {
    mac: options.mac ?? isMacPlatform(),
    isBindable,
    isReserved: (chord) => inputKeepsChord(chord),
  });
  return refused.length ? { ...resolved, issues: Object.freeze([...refused, ...resolved.issues]) } : resolved;
}

