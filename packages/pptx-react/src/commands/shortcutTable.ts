import type { ResolvedShortcuts } from '../../../../shared/keyboard/resolve';
import type { PptxCommandController } from './createPptxCommandStore';
import { normalizeChord, PPTX_COMMAND_DESCRIPTORS } from './descriptors';
import type { PptxPluginCommandId } from './types';

let defaultTable: ResolvedShortcuts | null = null;

/** The descriptors' own chords, canonical; they name `Mod`, so one table serves every platform. */
export function defaultShortcutTable(): ResolvedShortcuts {
  if (defaultTable) return defaultTable;
  const chords = new Map<string, readonly string[]>();
  const owners = new Map<string, string>();
  for (const descriptor of Object.values(PPTX_COMMAND_DESCRIPTORS)) {
    const held = descriptor.shortcuts.map(({ chord }) => normalizeChord(chord)!);
    if (held.length === 0) continue;
    chords.set(descriptor.id, Object.freeze(held));
    for (const chord of held) owners.set(chord, descriptor.id);
  }
  defaultTable = { chords, owners, issues: Object.freeze([]) };
  return defaultTable;
}

/** The effective shortcuts: canonical chords by id and the id holding each chord. */
export function shortcutTable(commands: PptxCommandController): ResolvedShortcuts {
  return commands.hostShortcuts() ?? defaultShortcutTable();
}

/** The shortcut id that holds a canonical chord. */
export function shortcutOwner(commands: PptxCommandController, chord: string): string | undefined {
  return shortcutTable(commands).owners.get(chord);
}

/** Whether the host moved or disabled the default holder of a canonical chord. */
export function releasedChord(commands: PptxCommandController, chord: string): boolean {
  const fallback = defaultShortcutTable().owners.get(chord);
  return fallback !== undefined && shortcutOwner(commands, chord) !== fallback;
}

/** Chords the host set for a contributed command; undefined keeps the plugin's own. */
export function hostPluginChords(
  commands: PptxCommandController,
  id: PptxPluginCommandId
): readonly string[] | undefined {
  return commands.hostShortcuts()?.chords.get(id);
}

/** Calls `listener` after each change of the shortcut map. */
export function subscribeShortcuts(commands: PptxCommandController, listener: () => void): () => void {
  let seen = commands.shortcutRevision();
  return commands.store.subscribe(() => {
    const revision = commands.shortcutRevision();
    if (revision === seen) return;
    seen = revision;
    listener();
  });
}
