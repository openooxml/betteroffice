import type { TranslationKey } from '@betteroffice/pptx-i18n';
import { ariaChord, formatChord, isMacPlatform, type ChordLabels } from '../../../../shared/keyboard/format';
import { defaultArgs } from './descriptors';
import type { PptxCommandArgs, PptxCommandId, PptxCommandShortcut, PptxCommandStore } from './types';

type Translate = (key: TranslationKey) => string;

/** A chord as people read it, with the key names in the editor's language, and its aria form. */
export function chordShortcut(
  chord: string,
  t: Translate,
  mac = isMacPlatform()
): { shortcut: string; ariaShortcut: string } {
  const labels: ChordLabels = Object.fromEntries(
    ['ctrl', 'alt', 'shift', 'delete', 'escape', 'pageup', 'pagedown'].map((key) => [key, t(`keys.${key}` as TranslationKey)])
  );
  return { shortcut: formatChord(chord, mac, labels), ariaShortcut: ariaChord(chord, mac) };
}

/** The chord `store` runs `id` with, in both forms; empty when it has none. */
export function storeShortcut<K extends PptxCommandId>(
  store: PptxCommandStore,
  id: K,
  t: Translate,
  args?: PptxCommandArgs[K],
  mac?: boolean
): { shortcut?: string; ariaShortcut?: string } {
  const key = JSON.stringify(args ?? defaultArgs(id));
  const shortcuts = store.getDescriptor(id).shortcuts as readonly PptxCommandShortcut[];
  const chord = shortcuts.find((shortcut) => JSON.stringify(shortcut.args) === key)?.chord;
  return chord ? chordShortcut(chord, t, mac) : {};
}
