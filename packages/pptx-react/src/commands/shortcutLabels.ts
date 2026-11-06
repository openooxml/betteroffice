import type { TranslationKey } from '@betteroffice/pptx-i18n';
import { ariaChord, formatChord, isMacPlatform, type ChordLabels } from '../../../../shared/keyboard/chords';
import { defaultArgs } from './descriptors';
import type { PptxCommandArgs, PptxCommandId, PptxCommandShortcut, PptxCommandStore } from './types';

type Translate = (key: TranslationKey) => string;

/** Key names in the editor's language. */
export function chordLabels(t: Translate): ChordLabels {
  return {
    ctrl: t('keys.ctrl'),
    alt: t('keys.alt'),
    shift: t('keys.shift'),
    delete: t('keys.del'),
    escape: t('keys.escape'),
    pageup: t('keys.pageUp'),
    pagedown: t('keys.pageDown'),
  };
}

function firstChord<K extends PptxCommandId>(
  store: PptxCommandStore,
  id: K,
  args: PptxCommandArgs[K] | undefined
): string | null {
  const key = JSON.stringify(args ?? defaultArgs(id));
  const shortcuts = store.getDescriptor(id).shortcuts as readonly PptxCommandShortcut[];
  return shortcuts.find((shortcut) => JSON.stringify(shortcut.args) === key)?.chord ?? null;
}

/** The chord `store` runs `id` with, as people read it, or null when it has none. */
export function storeShortcutLabel<K extends PptxCommandId>(
  store: PptxCommandStore,
  id: K,
  t: Translate,
  args?: PptxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const chord = firstChord(store, id, args);
  return chord ? formatChord(chord, mac, chordLabels(t)) : null;
}

/** The same chord in `aria-keyshortcuts` form. */
export function storeAriaShortcut<K extends PptxCommandId>(
  store: PptxCommandStore,
  id: K,
  args?: PptxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const chord = firstChord(store, id, args);
  return chord ? ariaChord(chord, mac) : null;
}
