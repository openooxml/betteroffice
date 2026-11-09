import type { TFunction } from '@betteroffice/docx-i18n';
import {
  ariaChord,
  formatChord,
  isMacPlatform,
  parseChord,
  type ChordLabels,
} from '../../../../shared/keyboard/chords';
import {
  resolveShortcuts,
  type ResolvedShortcuts,
  type ShortcutBinding,
  type ShortcutDefault,
} from '../../../../shared/keyboard/resolve';
import {
  DOCX_COMMAND_DESCRIPTORS,
  isDocxCommandId,
  isPluginCommandId,
  shortcutMatchesPlatform,
} from './descriptors';
import { acceptsNullArgs } from './evaluate';
import type {
  DocxCommandArgs,
  DocxCommandId,
  DocxCommandShortcut,
  DocxCommandStore,
  DocxPluginCommandId,
} from './types';

/**
 * Ids of the alignment shortcuts, one per value, so each chord moves on its own.
 * @experimental
 */
export type DocxShortcutVariantId =
  | 'alignment.left'
  | 'alignment.center'
  | 'alignment.right'
  | 'alignment.justify';

type DocxNullArgsCommandId = {
  [K in DocxCommandId]: null extends DocxCommandArgs[K] ? K : never;
}[DocxCommandId];

/**
 * Every id a host can bind: commands that run without arguments, alignment variants and plugin
 * commands.
 * @experimental
 */
export type DocxShortcutId = DocxNullArgsCommandId | DocxShortcutVariantId | DocxPluginCommandId;

/**
 * Host changes to the editor's shortcuts. An omitted id keeps its default; chords replace all of
 * its defaults; `null` disables it and leaves the keystroke to the host or browser. A chord taken
 * from another command's default wins over it. `cut`, `copy` and `paste` on a chord other than
 * their default use the async Clipboard API, which browsers may gate behind a prompt.
 * @experimental
 */
export type DocxShortcuts = Readonly<Partial<Record<DocxShortcutId, ShortcutBinding>>>;

const VARIANT_VALUES: Readonly<Record<DocxShortcutVariantId, string>> = {
  'alignment.left': 'left',
  'alignment.center': 'center',
  'alignment.right': 'right',
  'alignment.justify': 'both',
};
const VARIANT_IDS = Object.keys(VARIANT_VALUES) as DocxShortcutVariantId[];

/** Keys the text input moves or edits with, whatever the modifiers. */
const INPUT_KEYS: ReadonlySet<string> = new Set([
  'arrowup',
  'arrowdown',
  'arrowleft',
  'arrowright',
  'home',
  'end',
  'pageup',
  'pagedown',
  'backspace',
  'delete',
  'enter',
  'tab',
]);

function isVariantId(id: string): id is DocxShortcutVariantId {
  return Object.hasOwn(VARIANT_VALUES, id);
}

function shortcutDefaults(mac: boolean): ShortcutDefault[] {
  return Object.values(DOCX_COMMAND_DESCRIPTORS).flatMap((descriptor): ShortcutDefault[] => {
    const shortcuts = (descriptor.shortcuts as readonly DocxCommandShortcut[]).filter((shortcut) =>
      shortcutMatchesPlatform(shortcut, mac)
    );
    if (shortcuts.length === 0) return [];
    if (descriptor.id !== 'alignment') {
      return [{ id: descriptor.id, chords: shortcuts.map(({ chord }) => chord) }];
    }
    return VARIANT_IDS.map((id) => ({
      id,
      chords: shortcuts
        .filter(({ args }) => (args as { value: string }).value === VARIANT_VALUES[id])
        .map(({ chord }) => chord),
    }));
  });
}

function isBindable(id: string): boolean {
  return isPluginCommandId(id) || (isDocxCommandId(id) && acceptsNullArgs(id));
}

/**
 * The effective bindings a host map produces and what was wrong with it, for checking a map
 * outside the editor.
 * @experimental
 */
export function resolveDocxShortcuts(
  shortcuts: DocxShortcuts | undefined,
  options: { mac?: boolean } = {}
): ResolvedShortcuts {
  return resolveShortcuts(shortcutDefaults(options.mac ?? isMacPlatform()), shortcuts, {
    mac: options.mac ?? isMacPlatform(),
    isBindable,
    isReserved: (chord) => INPUT_KEYS.has(parseChord(chord).key),
  });
}

/** The command and arguments a shortcut id runs. */
export function shortcutTarget(id: string): { id: DocxCommandId; args: unknown } {
  return isVariantId(id)
    ? { id: 'alignment', args: { value: VARIANT_VALUES[id] } }
    : { id: id as DocxCommandId, args: null };
}

/** The bindings of `id` in `resolved`, in descriptor form. */
export function effectiveCommandShortcuts(
  resolved: ResolvedShortcuts,
  id: DocxCommandId
): readonly DocxCommandShortcut[] {
  const ids = id === 'alignment' ? VARIANT_IDS : [id];
  return ids.flatMap((shortcutId) =>
    (resolved.chords.get(shortcutId) ?? []).map((chord) => ({
      chord,
      args: shortcutTarget(shortcutId).args as never,
    }))
  );
}

/** Key names in the editor's language. */
export function chordLabels(t: TFunction): ChordLabels {
  return {
    ctrl: t('dialogs.keyboardShortcuts.keys.ctrl'),
    alt: t('dialogs.keyboardShortcuts.keys.alt'),
    shift: t('dialogs.keyboardShortcuts.keys.shift'),
    delete: t('dialogs.keyboardShortcuts.keys.del'),
    escape: t('dialogs.keyboardShortcuts.keys.escape'),
    pageup: t('dialogs.keyboardShortcuts.keys.pageUp'),
    pagedown: t('dialogs.keyboardShortcuts.keys.pageDown'),
  };
}

function firstChord<K extends DocxCommandId>(
  store: DocxCommandStore,
  id: K,
  args: DocxCommandArgs[K] | undefined
): string | null {
  const key = JSON.stringify(args ?? null);
  const shortcuts = store.getDescriptor(id).shortcuts as readonly DocxCommandShortcut[];
  return shortcuts.find((shortcut) => JSON.stringify(shortcut.args) === key)?.chord ?? null;
}

/** The chord `store` runs `id` with, as people read it, or null when it has none. */
export function storeShortcutLabel<K extends DocxCommandId>(
  store: DocxCommandStore,
  id: K,
  t: TFunction,
  args?: DocxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const chord = firstChord(store, id, args);
  return chord ? formatChord(chord, mac, chordLabels(t)) : null;
}

/** The same chord in `aria-keyshortcuts` form. */
export function storeAriaShortcut<K extends DocxCommandId>(
  store: DocxCommandStore,
  id: K,
  args?: DocxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const chord = firstChord(store, id, args);
  return chord ? ariaChord(chord, mac) : null;
}
